import type { IStyleData } from '@univerjs/core'
import type { CellFormatState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { ooxmlTextRotationToUniver } from '../edit-journal'
import { normalizeHexColor, toSelectionFormat } from '../selection-format'

/**
 * The daemon read family's own view of a cell's explicit formatting, and the
 * composition that builds it from a live cell. Kept here (not on the shared
 * `selection-format` / `CellFormatState` the rest of the app uses) so the read
 * never has to change the types other features depend on: `border` reports what
 * is on screen — per edge, with a line style and color — rather than the
 * write-side border patch `CellFormatState` carries.
 */
export interface ReadBorderEdge {
  /** an OOXML line-style name, e.g. "thin" / "medium" */
  readonly style: string
  /** the edge's color as plain hex, when it has one */
  readonly color?: string
}

/** the edges a read cell borders, keyed by side; an edge it does not border is absent */
export interface ReadCellBorders {
  readonly top?: ReadBorderEdge
  readonly right?: ReadBorderEdge
  readonly bottom?: ReadBorderEdge
  readonly left?: ReadBorderEdge
}

/** a read cell's explicit formatting: the shared CellFormatState with per-edge `border` */
export type ReadCellFormat = Omit<CellFormatState, 'border'> & {
  readonly border?: ReadCellBorders
}

/**
 * The scopes a read can be limited to: the five property groups (one per Format
 * Cells tab), plus `all` for the whole cell. A read that names a category reports
 * only that group's properties, so an agent can ask for just the font, or just
 * the borders, instead of the whole cell.
 */
export const FORMAT_CATEGORIES = ['font', 'fill', 'border', 'alignment', 'number', 'all'] as const
export type FormatCategory = (typeof FORMAT_CATEGORIES)[number]

/**
 * What a read reports when the caller names no category: the groups an agent
 * reaches for most (text, fill, alignment, number), leaving the per-edge
 * `border` — the one group that expands per side, with a style and color — as an
 * explicit ask.
 */
export const DEFAULT_FORMAT_CATEGORIES = ['font', 'fill', 'alignment', 'number'] as const

/** the wire keys of `ReadCellFormat` each category covers (`all` = every key) */
export const FORMAT_CATEGORY_KEYS = {
  font: ['fontFamily', 'fontSize', 'bold', 'italic', 'underline', 'strikethrough', 'fontColor'],
  fill: ['fillColor'],
  border: ['border'],
  alignment: ['horizontalAlign', 'verticalAlign', 'wrapText', 'textRotation', 'indent'],
  number: ['numberFormat'],
  all: [
    'fontFamily',
    'fontSize',
    'bold',
    'italic',
    'underline',
    'strikethrough',
    'fontColor',
    'fillColor',
    'border',
    'horizontalAlign',
    'verticalAlign',
    'wrapText',
    'textRotation',
    'indent',
    'numberFormat',
  ],
} as const satisfies Record<FormatCategory, readonly (keyof ReadCellFormat)[]>

/** the keys of `ReadCellFormat` that a single category covers */
type CategoryKeys<C extends FormatCategory> = (typeof FORMAT_CATEGORY_KEYS)[C][number]

/** a read cell's formatting scoped to one Format Cells tab, each a slice of `ReadCellFormat` */
export type ReadFontFormat = Pick<ReadCellFormat, CategoryKeys<'font'>>
export type ReadFillFormat = Pick<ReadCellFormat, CategoryKeys<'fill'>>
export type ReadBorderFormat = Pick<ReadCellFormat, CategoryKeys<'border'>>
export type ReadAlignmentFormat = Pick<ReadCellFormat, CategoryKeys<'alignment'>>
export type ReadNumberFormat = Pick<ReadCellFormat, CategoryKeys<'number'>>

/**
 * The compact, wire-facing view of a cell's explicit formatting: only the
 * properties actually set appear, so an unstyled cell yields `{}` (and the read
 * omits it). Borders come straight off the style data — per edge, with a line
 * style and color — since the echo has no room for them.
 */
export function formatStateOf(style: IStyleData, pattern: string): ReadCellFormat {
  const echo = toSelectionFormat(style, pattern)
  const format: Record<string, unknown> = {}
  if (echo.bold) format.bold = true
  if (echo.italic) format.italic = true
  if (echo.underline) format.underline = true
  if (echo.strike) format.strikethrough = true
  if (echo.fontFamily) format.fontFamily = echo.fontFamily
  if (echo.fontSize) format.fontSize = echo.fontSize
  if (echo.fontColor) format.fontColor = echo.fontColor
  if (echo.fillColor) format.fillColor = echo.fillColor
  if (echo.numberFormat && echo.numberFormat !== 'General') format.numberFormat = echo.numberFormat
  if (
    echo.horizontalAlignment === 'left' ||
    echo.horizontalAlignment === 'center' ||
    echo.horizontalAlignment === 'right'
  ) {
    format.horizontalAlign = echo.horizontalAlignment
  }
  if (
    echo.verticalAlignment === 'top' ||
    echo.verticalAlignment === 'center' ||
    echo.verticalAlignment === 'bottom'
  ) {
    format.verticalAlign = echo.verticalAlignment
  }
  if (echo.wrap) format.wrapText = true
  if (echo.textRotation) {
    // the echo speaks OOXML (0–180, 255 stacked); the read carries the DSL's
    // own vocabulary (degrees −90..90, or 'vertical')
    const rotation = ooxmlTextRotationToUniver(echo.textRotation)
    if (rotation) format.textRotation = rotation.v ? 'vertical' : rotation.a
  }
  if (echo.indent) format.indent = echo.indent
  const border = borderStateOf(style)
  if (border) format.border = border
  return format as ReadCellFormat
}

/** the minimal cell view the composition reads: its resolved style and number format */
export interface StyledRange {
  getCellStyleData(): IStyleData | null
  getNumberFormat(): string
}

/** Univer BorderStyleTypes value → the OOXML line-style name a read reports */
const BORDER_STYLE_NAMES: Record<number, string> = {
  1: 'thin',
  2: 'hair',
  3: 'dotted',
  4: 'dashed',
  5: 'dashDot',
  6: 'dashDotDot',
  7: 'double',
  8: 'medium',
  9: 'mediumDashed',
  10: 'mediumDashDot',
  11: 'mediumDashDotDot',
  12: 'slantDashDot',
  13: 'thick',
}

/** Univer's border edge keys, in a stable top/right/bottom/left read order */
const BORDER_SIDES = [
  ['t', 'top'],
  ['r', 'right'],
  ['b', 'bottom'],
  ['l', 'left'],
] as const

/**
 * The per-edge borders a cell's style data carries: an edge with a line style
 * becomes `{ style, color }` (color as plain hex); an edge without one is left
 * out, and a cell with no borders at all yields `undefined`.
 */
export function borderStateOf(style: IStyleData): ReadCellBorders | undefined {
  const bd = style.bd as unknown as Record<string, unknown> | undefined
  if (!bd) return undefined
  const borders: Record<string, ReadBorderEdge> = {}
  for (const [key, side] of BORDER_SIDES) {
    const edge = bd[key]
    if (typeof edge !== 'object' || edge === null) continue
    const { s, cl } = edge as { s?: unknown; cl?: unknown }
    const style0 = typeof s === 'number' ? BORDER_STYLE_NAMES[s] : undefined
    if (style0 === undefined) continue
    const color = normalizeHexColor((cl as { rgb?: string } | undefined)?.rgb ?? null)
    borders[side] = { style: style0, ...(color === null ? {} : { color }) }
  }
  return Object.keys(borders).length > 0 ? (borders as ReadCellBorders) : undefined
}

/**
 * The composition the daemon's format read runs per cell: the cell's resolved
 * style turned into the compact wire view above. When `categories` names scopes,
 * only those groups' properties are kept (a cell with no formatting still yields
 * `{}`); omitted or containing `all` reports the whole cell.
 */
export function cellFormatStateOf(
  range: StyledRange,
  categories?: readonly FormatCategory[],
): ReadCellFormat {
  let pattern = ''
  try {
    pattern = range.getNumberFormat() ?? ''
  } catch {
    // number-format resolution can fail on never-touched cells; treat as General
  }
  const format = formatStateOf(range.getCellStyleData() ?? {}, pattern)
  if (!categories || categories.includes('all')) return format
  const keys = new Set(categories.flatMap((category) => FORMAT_CATEGORY_KEYS[category]))
  const scoped: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(format)) {
    if (keys.has(key as keyof ReadCellFormat)) scoped[key] = value
  }
  return scoped as ReadCellFormat
}
