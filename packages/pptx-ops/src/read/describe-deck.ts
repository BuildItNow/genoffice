import { placeholderSource } from '@genoffice/agent-core/redact-core'
import {
  builtinTableStyleName,
  getSlideLinks,
  getSlideNotes,
  listSlideLayouts,
  type LinkTarget,
  type OpenedPptx,
  type Paragraph,
  type Slide,
  type SlideElement,
} from '@genoffice/pptx-engine'
import { elementDurableId, listSlideAnimations, slideDurableId, type AnimationEntry } from '../ops'
import { clipText } from './clip'
import { fromEmu, type ReadUnit } from './units'

const EMU_PER_INCH = 914400

/** Displayed style of the first run and the inheritance layer each value came from. */
export interface EffectiveStyle {
  fontSizePt?: number
  fontFamily?: string
  color?: string
  bold: boolean
  italic: boolean
  align?: string
  src: Record<string, string>
}

export interface ElementSummary {
  id: string | null
  type: string
  kind?: string
  name?: string
  placeholder?: string
  box: { x: number; y: number; cx: number; cy: number; unit?: ReadUnit }
  text?: string
  /** the text was clipped to the preview length; `--full` or `--max-chars` returns the rest */
  truncated?: true
  effective?: EffectiveStyle
  /** whole-element hyperlink: url, slide jump or named show action (`setLink` writes it) */
  link?: LinkTarget
  rows?: number
  cols?: number
  style?: { id: string; name: string | null; flags: string[] }
  children?: ElementSummary[]
}

export interface LayoutSummary {
  index: number
  name: string
  type: string
  placeholders: Array<{ type: string; idx: string }>
}
export const PREVIEW_CHARS = 300

export interface SlideSummary {
  index: number
  id: string
  elements: ElementSummary[]
  /** the page timeline in play order; absent when the slide has no animations */
  animations?: AnimationEntry[]
  notes?: string
}

export interface DeckSummary {
  slides: number
  size: { cx: number; cy: number; unit?: ReadUnit; inches: { width: number; height: number } }
  emu_per_inch: number
  layouts?: LayoutSummary[]
  pages: SlideSummary[]
}

/**
 * Structure an agent needs to target ops: durable ids, geometry in EMU, text previews.
 * `full` drops the clipping (whole text, every table row) and adds speaker notes, for
 * reading a deck as source material rather than targeting it.
 *
 * Shared by the genoffice CLI (`slides read`) and the app's MCP `read_deck`, so a
 * deck reads identically whether it came from disk or from an open tab.
 */
export function describeDeck(
  opened: OpenedPptx,
  only?: number,
  full = false,
  maxChars = full ? Infinity : PREVIEW_CHARS,
  layouts = false,
  unit: ReadUnit = 'emu',
): DeckSummary {
  const { deck } = opened
  const pages = deck.slides
    .map((slide, index) => ({ slide, index }))
    .filter(({ index }) => only === undefined || index === only)
    .map(({ slide, index }) => summarizeSlide(opened, slide, index, full, maxChars))
  if (unit !== 'emu')
    for (const page of pages) for (const el of page.elements) convertBoxes(el, unit)
  return {
    slides: deck.slides.length,
    size: {
      cx: fromEmu(deck.size.cx, unit),
      cy: fromEmu(deck.size.cy, unit),
      ...(unit !== 'emu' ? { unit } : {}),
      inches: {
        width: round(deck.size.cx / EMU_PER_INCH),
        height: round(deck.size.cy / EMU_PER_INCH),
      },
    },
    emu_per_inch: EMU_PER_INCH,
    ...(layouts ? { layouts: describeLayouts(opened) } : {}),
    pages,
  }
}

/** Layouts an `addSlideWithLayout` / `setSlideLayout` op can name, with their content placeholders. */
export function describeLayouts(opened: OpenedPptx): LayoutSummary[] {
  return listSlideLayouts(opened.archive).map((l, index) => ({
    index,
    name: l.name,
    type: l.layoutType,
    placeholders: l.placeholders.map((ph) => ({ type: ph.type || 'body', idx: ph.idx })),
  }))
}

function convertBoxes(el: ElementSummary, unit: ReadUnit): void {
  const { x, y, cx, cy } = el.box
  el.box = {
    x: fromEmu(x, unit),
    y: fromEmu(y, unit),
    cx: fromEmu(cx, unit),
    cy: fromEmu(cy, unit),
    unit,
  }
  for (const child of el.children ?? []) convertBoxes(child, unit)
}

function summarizeSlide(
  opened: OpenedPptx,
  slide: Slide,
  index: number,
  full: boolean,
  maxChars: number,
): SlideSummary {
  const links = new Map(getSlideLinks(opened, index).map((l) => [l.elementId, l.target]))
  const out: SlideSummary = {
    index,
    id: slideDurableId(slide),
    elements: slide.elements.map((el) => summarizeElement(el, full, maxChars, links)),
  }
  const animations = listSlideAnimations(slide)
  if (animations.length) out.animations = animations
  if (full) {
    const notes = getSlideNotes(opened.archive, slide.path)
    if (notes) out.notes = notes
  }
  return out
}

function summarizeElement(
  el: SlideElement,
  full: boolean,
  maxChars: number,
  links: Map<string, LinkTarget>,
): ElementSummary {
  const { offset } = el.transform
  const out: ElementSummary = {
    id: elementDurableId(el),
    type: el.type,
    box: { x: offset.x, y: offset.y, cx: offset.cx, cy: offset.cy },
  }
  if (el.name) out.name = el.name
  if (el.placeholder) out.placeholder = el.placeholder
  const link = links.get(el.id)
  if (link) out.link = link
  switch (el.type) {
    case 'text':
    case 'shape': {
      const { text, truncated } = clipText(paragraphsText(el.text?.paragraphs ?? []), maxChars)
      if (text) out.text = text
      if (truncated) out.truncated = true
      const effective = effectiveStyle(el.text?.paragraphs ?? [])
      if (effective) out.effective = effective
      break
    }
    case 'table': {
      out.rows = el.rows.length
      out.cols = el.colWidths.length
      if (el.styleId) {
        out.style = {
          id: el.styleId,
          name: builtinTableStyleName(el.styleId) ?? null,
          flags: Object.entries(el.styleFlags ?? {})
            .filter(([, on]) => on)
            .map(([k]) => k),
        }
      }
      let clipped = !full && el.rows.length > 3
      out.text = (full ? el.rows : el.rows.slice(0, 3))
        .map((r) =>
          r
            .map((c) => {
              const cell = clipText(paragraphsText(c.text?.paragraphs ?? []), maxChars)
              clipped ||= cell.truncated
              return cell.text
            })
            .join(' | '),
        )
        .join('\n')
      if (clipped) out.truncated = true
      break
    }
    case 'group':
      out.children = el.children.map((child) => summarizeElement(child, full, maxChars, links))
      break
    case 'picture':
      // a withheld picture reads as its marker, so the reader can see something
      // is there without handing the model a reference it could fetch. Its
      // target never reaches this summary either way (only `kind` would).
      if (el.redact) out.text = placeholderSource(el.redact)
      break
    case 'passthrough':
      out.kind = el.kind
      break
  }
  return out
}

function effectiveStyle(paragraphs: Paragraph[]): EffectiveStyle | undefined {
  const paragraph = paragraphs.find((p) => p.runs.some((r) => r.text))
  const run = paragraph?.runs.find((r) => r.text)
  if (!paragraph || !run) return undefined
  const src: Record<string, string> = { ...run.styleSrc }
  if (paragraph.alignSrc) src.align = paragraph.alignSrc
  return {
    ...(run.fontSize != null ? { fontSizePt: run.fontSize } : {}),
    ...(run.fontFamily ? { fontFamily: run.fontFamily } : {}),
    ...(run.color ? { color: run.color } : {}),
    bold: !!run.bold,
    italic: !!run.italic,
    ...(paragraph.align ? { align: paragraph.align } : {}),
    src,
  }
}

/** A withheld run reads as its marker; the words it covers never leave the model. */
function paragraphsText(paragraphs: { runs: { text: string; redact?: string }[] }[]): string {
  return paragraphs
    .map((p) => p.runs.map((r) => (r.redact ? placeholderSource(r.redact) : r.text)).join(''))
    .join('\n')
}

function round(n: number): number {
  return Math.round(n * 100) / 100
}
