import { z } from 'zod'
import {
  parseAddress,
  parseRange,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  describeOperationErrors,
  workbookOperationSchema,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { normalizeSheetRefs, primaryCellOf, primarySheetId } from '../mcp-sheet-refs'
import type { McpSheetHandlers } from '../mcp-bridge'
import { FORMAT_CATEGORIES, type FormatCategory } from './format-state'
import { SHEETLESS_OPS } from './sheetless-ops'
import {
  type DaemonCommandResult,
  type RangeFormatsResult,
  type RangeReadResult,
  type SheetApplyOpsPayload,
  type SheetApplyOpsResult,
  type SheetInfoResult,
  type SheetReadFormatsPayload,
  type SheetReadFormatsResult,
  type SheetReadPayload,
  type SheetReadResult,
  type SheetSearchPayload,
  type SheetSearchResult,
} from './types'

/**
 * The daemon's own `office_sheet_*` commands, dispatched by the renderer bridge.
 *
 * The built-in MCP session tools live in `mcp-bridge.ts`; the daemon's `office_*`
 * family (reached over the control socket, never through the app's built-in MCP
 * server) is implemented here so the bridge stays a transport. The commands ride
 * the *same* channel and reuse the same readers/executors (`McpSheetHandlers`),
 * so a daemon edit lands in the edit journal, undo history and save pipeline
 * exactly like an in-app or MCP one.
 *
 * `dispatchDaemonCommand` owns only its own command names; `mcp-bridge`'s switch
 * handles the rest. The two small helpers it shares with that switch
 * (`addressInBounds`, `describeOpError`) live here and are imported back, so the
 * logic stays in one place.
 */

/** office_sheet_search match-count bounds (default/maximum hits per call) */
const SEARCH_RESULTS_DEFAULT = 100
const SEARCH_MAX_RESULTS = 500

/** fields an op uses to name its OWN sheet (a source/target ref does not count) */
const OWN_SHEET_FIELDS = ['sheet', 'sheetId'] as const

/**
 * Run one daemon command against the live workbook. A name this family does not
 * own is not an error here — the caller's own switch handles it — so it returns
 * `undefined` and the caller can tell the command was unowned. For a command it
 * does own, a thrown error is reported to the caller as that command's failure.
 */
export async function dispatchDaemonCommand(
  command: string,
  payload: Record<string, unknown>,
  handlers: McpSheetHandlers,
): Promise<DaemonCommandResult | undefined> {
  // A name this family does not own matches no case and falls through to
  // `undefined` below. The payload crosses the ipc boundary as an untyped JSON
  // object; each case reads it through its own payload type (the shape the
  // tool advertised).
  switch (command) {
    case 'office_sheet_info':
      return sheetInfo(handlers)
    case 'office_sheet_read':
      return sheetRead(payload as unknown as SheetReadPayload, handlers)
    case 'office_sheet_read_formats':
      return sheetReadFormats(payload as unknown as SheetReadFormatsPayload, handlers)
    case 'office_sheet_search':
      return sheetSearch(payload as unknown as SheetSearchPayload, handlers)
    case 'office_sheet_apply_ops':
      return sheetApplyOps(payload as unknown as SheetApplyOpsPayload, handlers)
  }
}

/**
 * The worksheets in the file and which one is active, each with its data extent
 * (rows/columns in use). The renderer's own context is far wider (selection,
 * known addresses, merges, charts, …); an agent needs the sheet names to address
 * and a sense of size, so narrow it here rather than shipping the whole context.
 * `current` is the active sheet's NAME (the sheet tools address sheets by name);
 * each summary also carries the session `id`.
 */
function sheetInfo(handlers: McpSheetHandlers): SheetInfoResult {
  const context = handlers.context() as {
    sheetId: string
    sheets: readonly { id: string; name: string; rows?: number; columns?: number }[]
  }
  return {
    current: context.sheets.find((sheet) => sheet.id === context.sheetId)?.name ?? '',
    sheets: context.sheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      rows: sheet.rows ?? 0,
      columns: sheet.columns ?? 0,
    })),
  }
}

/** does an A1 address fall inside an inclusive range? (malformed address → false) */
export function addressInBounds(address: string, bounds: RangeBounds): boolean {
  let coords
  try {
    coords = parseAddress(address)
  } catch {
    return false
  }
  return (
    coords.row >= bounds.startRow &&
    coords.row <= bounds.endRow &&
    coords.column >= bounds.startColumn &&
    coords.column <= bounds.endColumn
  )
}

/** First schema issue, with the offending op named so the client can fix it. */
export function describeOpError(ops: unknown[], error: z.ZodError): string {
  const text = describeOperationErrors(ops, error)
  return error.issues.some((issue) => issue.path.includes('sheetId'))
    ? `${text}\nUse a sheetId from the workbook overview (office_sheet_info / read_sheet).`
    : text
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() !== '' ? value.trim() : undefined
}

/** The worksheet the tab's view is on — what an empty `sheet` reference means. */
function activeSheetId(handlers: McpSheetHandlers): string | undefined {
  const context = handlers.context() as { sheetId?: string }
  return typeof context.sheetId === 'string' ? context.sheetId : undefined
}

/**
 * Resolve a sheet reference (`sheet`, a NAME — never an id) against the open
 * workbook, to the id the DSL addresses. The name must match exactly (case- and
 * space-sensitive). An empty reference means the sheet the view is on; a name
 * that matches no sheet fails loudly instead of addressing the wrong one.
 */
function resolveSheetRef(
  sheet: string | undefined,
  handlers: McpSheetHandlers,
): string | undefined {
  if (sheet === undefined || sheet === '') return activeSheetId(handlers)
  const sheets = handlers.sheets()
  const match = sheets.find((candidate) => candidate.name === sheet)
  if (match) return match.id
  const known = sheets.length === 0 ? 'none' : sheets.map((candidate) => candidate.name).join(', ')
  throw new Error(
    `no worksheet named "${sheet}" in this workbook (sheets: ${known}); call office_sheet_info for the current sheets`,
  )
}

/**
 * The target worksheet's descriptor — id, name, data extent and frozen panes.
 * With no `sheet` the sheet the view is on is read. When `range` is given, that
 * area's value grid (with formulas and the merged ranges touching it) is merged
 * in, paged by `maxRows`.
 */
async function sheetRead(
  payload: SheetReadPayload,
  handlers: McpSheetHandlers,
): Promise<SheetReadResult> {
  // no sheet ⇒ the active sheet; the name resolves to the id it addresses
  const sheetId = resolveSheetRef(payload.sheet, handlers)
  const meta = handlers.readSheetMeta(sheetId)
  if (!meta) throw new Error(`no worksheet "${payload.sheet ?? '(active)'}" is open in this tab`)
  const range = asString(payload.range)
  if (range === undefined) return meta
  const maxRows = typeof payload.maxRows === 'number' ? payload.maxRows : undefined
  const grid = (await handlers.readRange(
    range,
    sheetId,
    maxRows === undefined ? {} : { maxRows },
  )) as RangeReadResult
  return { ...meta, range: grid }
}

/**
 * The target worksheet's descriptor, with the explicit cell formatting over an
 * A1 area (only cells with a set format appear) nested under `range` — the same
 * shape as `office_sheet_read`, so a caller reads values and formats alike. The
 * area is paged by rows exactly like the value read: the renderer cuts it to what
 * one call can carry and names the range to continue from.
 */
async function sheetReadFormats(
  payload: SheetReadFormatsPayload,
  handlers: McpSheetHandlers,
): Promise<SheetReadFormatsResult> {
  const sheetId = resolveSheetRef(payload.sheet, handlers)
  const meta = handlers.readSheetMeta(sheetId)
  if (!meta) throw new Error(`no worksheet "${payload.sheet ?? '(active)'}" is open in this tab`)
  const range = asString(payload.range)
  if (range === undefined) throw new Error('office_sheet_read_formats needs a "range"')
  const maxRows = typeof payload.maxRows === 'number' ? payload.maxRows : undefined
  const category = Array.isArray(payload.category)
    ? payload.category.filter((name): name is FormatCategory =>
        FORMAT_CATEGORIES.includes(name as FormatCategory),
      )
    : []
  const formats = (await handlers.readRangeFormats(range, sheetId, {
    ...(maxRows === undefined ? {} : { maxRows }),
    ...(category.length === 0 ? {} : { category }),
  })) as RangeFormatsResult
  return { ...meta, range: formats }
}

/** Value/formula search over a sheet (or the whole workbook), addressed like the reads. */
async function sheetSearch(
  payload: SheetSearchPayload,
  handlers: McpSheetHandlers,
): Promise<SheetSearchResult> {
  const query = typeof payload.query === 'string' ? payload.query : ''
  if (!query.trim()) throw new Error('office_sheet_search needs a non-empty "query"')
  // unlike the single-sheet reads, an unnamed sheet here searches every worksheet,
  // so only a named sheet is resolved (to its id) — empty stays unfiltered
  const sheetId = payload.sheet ? resolveSheetRef(payload.sheet, handlers) : undefined
  const range = asString(payload.range)
  let bounds: RangeBounds | undefined
  if (range !== undefined) {
    try {
      bounds = parseRange(range)
    } catch {
      throw new Error(`invalid range "${range}"`)
    }
  }
  const lookInRaw = payload.lookIn
  const maxRaw = Number(payload.maxResults)
  const outcome = await handlers.findCells({
    query,
    regex: payload.regex === true,
    lookIn: lookInRaw === 'values' || lookInRaw === 'formulas' ? lookInRaw : 'both',
    errorsOnly: false,
    matchCase: payload.matchCase === true,
    sheetId,
    maxResults:
      Number.isFinite(maxRaw) && maxRaw >= 1
        ? Math.min(Math.trunc(maxRaw), SEARCH_MAX_RESULTS)
        : SEARCH_RESULTS_DEFAULT,
  })
  if (outcome.error) throw new Error(outcome.error)
  const kept = bounds
    ? outcome.matches.filter((match) => addressInBounds(match.address, bounds))
    : outcome.matches
  return {
    total: kept.length,
    matches: kept.map((match) => ({
      sheet: match.sheetName,
      address: match.address,
      value: match.value,
      ...(match.formula === undefined ? {} : { formula: match.formula }),
    })),
    truncated: outcome.truncated,
    ...(outcome.incompleteSheets.length ? { incompleteSheets: [...outcome.incompleteSheets] } : {}),
  }
}

/**
 * Seed a default sheet into the ops that are scoped to a worksheet but name none
 * of their own, so `office_sheet_apply_ops(tabId, sheet, ops)` need not repeat the
 * sheet on every op. Whether an op takes a sheet is decided by its NAME against
 * the DSL schema (`SHEETLESS_OPS`), not by whether it happens to carry a cell
 * address — a sheet-scoped op like set_freeze is seeded, a workbook-level op like
 * add_sheet is left be.
 *
 * An op names its own sheet only with a non-empty `sheet`/`sheetId`: an empty or
 * blank ref counts as absent (mcp-sheet-refs strips it the same way), so the op
 * still gets the default rather than falling through to a schema failure.
 */
function seedSheetId(op: unknown, sheetId: string): unknown {
  if (typeof op !== 'object' || op === null || Array.isArray(op)) return op
  const record = { ...(op as Record<string, unknown>) }
  const name = typeof record.op === 'string' ? record.op : undefined
  if (name === undefined || SHEETLESS_OPS.has(name)) return record
  const named = OWN_SHEET_FIELDS.some((field) => {
    const value = record[field]
    return typeof value === 'string' && value.trim() !== ''
  })
  if (!named) record.sheetId = sheetId
  return record
}

/** Apply a workbook DSL batch to the visible grid, addressed by the named sheet. */
async function sheetApplyOps(
  payload: SheetApplyOpsPayload,
  handlers: McpSheetHandlers,
): Promise<SheetApplyOpsResult> {
  const ops = Array.isArray(payload.ops) ? payload.ops : []
  if (ops.length === 0) throw new Error('ops must be a non-empty array')
  // the seeded sheet follows `sheet`'s semantics — the named one, or the tab's
  // active sheet when it is omitted — resolved here to the id the ops take
  const defaultSheetId = resolveSheetRef(payload.sheet, handlers)
  const seeded =
    defaultSheetId === undefined ? ops : ops.map((op) => seedSheetId(op, defaultSheetId))
  const named = normalizeSheetRefs(seeded, handlers.sheets())
  if (!named.ok) throw new Error(named.error)
  const parsed = z.array(workbookOperationSchema).safeParse(named.ops)
  if (!parsed.success) throw new Error(describeOpError(named.ops, parsed.error))
  const sheetId = primarySheetId(parsed.data)
  const dryRun = payload.dryRun === true
  if (sheetId !== undefined && !dryRun) handlers.focusSheet(sheetId, primaryCellOf(parsed.data))
  const result = await handlers.applyOps(parsed.data, dryRun)
  const outcome =
    result && typeof result === 'object' ? (result as { ok?: boolean; reason?: string }) : undefined
  if (outcome?.ok === false) throw new Error(outcome.reason ?? 'the batch could not be applied')
  return result as SheetApplyOpsResult
}
