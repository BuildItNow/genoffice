import type { ApplyOutcome, CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { FormatCategory, ReadCellFormat } from './format-state'

/**
 * The shapes the daemon's `office_sheet_*` commands exchange with the renderer.
 *
 * The shell strips `tabId` before forwarding (it addresses the tab itself), so a
 * payload below carries only what the renderer needs to run the command. Each
 * `*Payload` is what arrives on the command, each `*Result` is what the command
 * resolves to; both are grouped here so the dispatcher and its tests share one
 * definition.
 */

/** office_sheet_info — no renderer arguments (it reports the whole workbook) */
export type SheetInfoPayload = Record<string, never>
/** one worksheet of the workbook, as office_sheet_info lists it */
export interface SheetSummary {
  id: string
  name: string
  /** rows in use (the sheet's data extent) */
  rows: number
  /** columns in use (the sheet's data extent) */
  columns: number
}
/** office_sheet_info — the file's worksheets and which one is active */
export interface SheetInfoResult {
  /** the name of the sheet the view is currently on */
  current: string
  /** every worksheet in the file, in tab order */
  sheets: SheetSummary[]
}

/** office_sheet_read — one worksheet's structural report plus, optionally, a range */
export interface SheetReadPayload {
  /** worksheet name; omitted means the tab's active sheet */
  sheet?: string
  /** A1 area to read; omitted returns the structural report alone */
  range?: string
  /** page a tall range: stop after this many rows */
  maxRows?: number
}

/** the renderer's range read (`readSheetRange`): a value grid paged by rows */
export interface RangeReadResult {
  /** the worksheet's name, when it could be resolved */
  sheet?: string
  /** the sub-range actually returned this call */
  range: string
  /** rows of the sub-range; null cells are empty */
  rows: unknown[][]
  /** formulas in the sub-range, keyed by A1 address (`=…`) */
  formulas: Record<string, string>
  /** the merged areas intersecting this range, as A1 ranges (`[]` when none) */
  merges: string[]
  rowsShown: number
  rowsTotal: number
  /** true when the range was paged — more rows remain */
  truncated: boolean
  /** how to continue, when `truncated` */
  note?: string
}

/** frozen-pane state, exactly as Univer's `getFreeze()` reports it (all-zero = none) */
export interface SheetFreeze {
  /** count of fixed columns (held from the left) */
  xSplit: number
  /** count of fixed rows (held from the top) */
  ySplit: number
  /** the first scrollable row of the viewport */
  startRow: number
  /** the first scrollable column of the viewport */
  startColumn: number
}

/** one worksheet distilled to plain data, as the bridge hands it to the commands */
export interface SheetMeta {
  id: string
  name: string
  /** rows in use (the sheet's data extent) */
  rows: number
  /** columns in use (the sheet's data extent) */
  columns: number
  /** frozen panes (the full `getFreeze()` state; all-zero when none) */
  freeze: SheetFreeze
}

/** office_sheet_read — the target sheet's descriptor plus, optionally, a range grid */
export interface SheetReadResult extends SheetMeta {
  /** when a `range` was requested: that area's value grid, paged by `maxRows` */
  range?: RangeReadResult
}

/** office_sheet_read_formats — the explicit formatting over an A1 area, paged like the value read */
export interface SheetReadFormatsPayload {
  /** worksheet name; omitted means the tab's active sheet */
  sheet?: string
  /** the A1 area whose formatting to read (required) */
  range: string
  /** page a tall area: stop after this many rows */
  maxRows?: number
  /** narrow each cell to the named Format Cells groups (font/fill/border/alignment/number); omit for the common font/fill/alignment/number set, or pass `all` for the whole cell */
  category?: FormatCategory[]
}

/** the renderer's format read (`readSheetRangeFormats`): a per-cell format grid paged by rows */
export interface RangeFormatsResult {
  /** the worksheet's name, when it could be resolved */
  sheet?: string
  /** the sub-range actually returned this call */
  range: string
  /** rows of the sub-range; each cell is the explicit format, `null` when the cell has none */
  rows: (ReadCellFormat | null)[][]
  /** the merged areas intersecting this range, as A1 ranges (`[]` when none) */
  merges: string[]
  rowsShown: number
  rowsTotal: number
  /** true when the range was paged — more rows remain */
  truncated: boolean
  /** how to continue, when `truncated` */
  note?: string
}

/** office_sheet_read_formats — the target sheet's descriptor plus, nested, the area's formatting */
export interface SheetReadFormatsResult extends SheetMeta {
  /** the area's per-cell formatting, paged by `maxRows` */
  range: RangeFormatsResult
}

/** office_sheet_search — find text across a worksheet (or the whole workbook) */
export interface SheetSearchPayload {
  /** worksheet name; omitted searches every worksheet */
  sheet?: string
  query: string
  /** scope the search to an A1 area */
  range?: string
  regex?: boolean
  lookIn?: 'values' | 'formulas' | 'both'
  /** match case-sensitively (default is case-insensitive) */
  matchCase?: boolean
  maxResults?: number
}
export interface SheetSearchMatch {
  /** the worksheet the hit is on */
  sheet: string
  address: string
  value: CellScalar
  formula?: string
}
export interface SheetSearchResult {
  total: number
  matches: SheetSearchMatch[]
  /** the scan stopped early — more matches may exist */
  truncated: boolean
  /** worksheets whose background indexing had not finished */
  incompleteSheets?: string[]
}

/** office_sheet_apply_ops — a workbook DSL batch applied to one worksheet */
export interface SheetApplyOpsPayload {
  /** worksheet name seeded into ops that omit a sheet; omitted means the tab's active sheet */
  sheet?: string
  /** workbook DSL op objects, A1-addressed */
  ops: unknown[]
  /** plan the batch and report what would change, without modifying the grid */
  dryRun?: boolean
}
/** office_sheet_apply_ops — the apply outcome, plus the plan (dryRun) or settled values */
export interface SheetApplyOpsResult extends ApplyOutcome {
  readonly dryRun?: boolean
  readonly structuralChanges?: readonly string[]
  readonly formatChanges?: readonly string[]
  readonly cellChanges?: readonly { address: string; after: CellScalar }[]
  readonly cellChangeCount?: number
  readonly sheetRenames?: readonly string[]
  /** the values formulas settled to after the batch (live edits only) */
  readonly formulaValues?: readonly unknown[]
}

/** what a daemon command resolves to; `undefined` means the name is not ours */
export type DaemonCommandResult =
  SheetInfoResult | SheetReadResult | SheetReadFormatsResult | SheetSearchResult | SheetApplyOpsResult
