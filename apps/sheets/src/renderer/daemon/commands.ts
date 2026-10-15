import { z } from 'zod'
import { parseAddress, parseRange, type RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  describeOperationErrors,
  workbookOperationSchema,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { normalizeSheetRefs, primaryCellOf, primarySheetId } from '../mcp-sheet-refs'
import type { McpSheetHandlers } from '../mcp-bridge'
import {
  SHEETS_DAEMON_COMMANDS,
  type DaemonCommandName,
  type DaemonCommandResult,
  type RangeFormatsResult,
  type RangeReadResult,
  type SheetEditPayload,
  type SheetEditResult,
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

const PAYLOAD_FIELDS = [
  'sheet',
  'sheetId',
  'sourceSheet',
  'sourceSheetId',
  'targetSheet',
  'targetSheetId',
] as const
/** fields that mark an op as targeting a cell/area (as opposed to a whole sheet) */
const AREA_FIELDS = ['address', 'range', 'target', 'source'] as const

/** the daemon command names this module serves (single-sourced in `shared/`) */
const DAEMON_COMMANDS: ReadonlySet<DaemonCommandName> = new Set(SHEETS_DAEMON_COMMANDS)

/** does `command` belong to the daemon's `office_sheet_*` family? */
export function isDaemonCommand(command: string): command is DaemonCommandName {
  return DAEMON_COMMANDS.has(command as DaemonCommandName)
}

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
  if (!isDaemonCommand(command)) return undefined
  // the payload crosses the ipc boundary as an untyped JSON object; each case
  // reads it through its own payload type (the shape the tool advertised)
  switch (command) {
    case 'office_sheet_info':
      return sheetInfo(handlers)
    case 'office_sheet_read':
      return sheetRead(payload as unknown as SheetReadPayload, handlers)
    case 'office_sheet_read_formats':
      return sheetReadFormats(payload as unknown as SheetReadFormatsPayload, handlers)
    case 'office_sheet_search':
      return sheetSearch(payload as unknown as SheetSearchPayload, handlers)
    case 'office_sheet_edit':
      return sheetEdit(payload as unknown as SheetEditPayload, handlers)
  }
}

/**
 * The worksheets in the file and which one is active, each with its data extent
 * (rows/columns in use). The renderer's own context is far wider (selection,
 * known addresses, merges, charts, …); an agent only needs the sheet ids/names
 * to address and a sense of size, so narrow it here rather than shipping the
 * whole context.
 */
function sheetInfo(handlers: McpSheetHandlers): SheetInfoResult {
  const context = handlers.context() as {
    sheetId: string
    sheets: readonly { id: string; name: string; rows?: number; columns?: number }[]
  }
  return {
    currentSheetId: context.sheetId,
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

/**
 * Resolve a sheet reference (a name or an id) against the open workbook. A name
 * that no longer matches fails loudly instead of addressing the wrong sheet;
 * `undefined` means "the active sheet".
 */
function resolveSheetRef(value: unknown, handlers: McpSheetHandlers): string | undefined {
  const ref = asString(value)
  if (ref === undefined) return undefined
  const sheets = handlers.sheets()
  if (sheets.some((sheet) => sheet.id === ref)) return ref
  const match = sheets.find((sheet) => sheet.name.trim().toLowerCase() === ref.toLowerCase())
  if (match) return match.id
  const known = sheets.length === 0 ? 'none' : sheets.map((sheet) => sheet.name).join(', ')
  throw new Error(
    `no worksheet named "${ref}" in this workbook (sheets: ${known}); call office_sheet_info for the current sheets`,
  )
}

/**
 * The target worksheet's descriptor — id, name, data extent and frozen panes.
 * With no `sheetId` the sheet the view is on is read. When `range` is given, that
 * area's value grid (with formulas and the merged ranges touching it) is merged
 * in, paged by `maxRows`.
 */
async function sheetRead(
  payload: SheetReadPayload,
  handlers: McpSheetHandlers,
): Promise<SheetReadResult> {
  // no sheetId ⇒ the active sheet; a name resolves to the id it addresses
  const sheetId = resolveSheetRef(payload.sheetId, handlers)
  const meta = handlers.readSheetMeta(sheetId)
  if (!meta) throw new Error(`no worksheet "${payload.sheetId ?? '(active)'}" is open in this tab`)
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
  const sheetId = resolveSheetRef(payload.sheetId, handlers)
  const meta = handlers.readSheetMeta(sheetId)
  if (!meta) throw new Error(`no worksheet "${payload.sheetId ?? '(active)'}" is open in this tab`)
  const range = asString(payload.range)
  if (range === undefined) throw new Error('office_sheet_read_formats needs a "range"')
  const maxRows = typeof payload.maxRows === 'number' ? payload.maxRows : undefined
  const formats = (await handlers.readRangeFormats(
    range,
    sheetId,
    maxRows === undefined ? {} : { maxRows },
  )) as RangeFormatsResult
  return { ...meta, range: formats }
}

/** Value/formula search over a sheet (or the whole workbook), addressed like the reads. */
async function sheetSearch(
  payload: SheetSearchPayload,
  handlers: McpSheetHandlers,
): Promise<SheetSearchResult> {
  const query = typeof payload.query === 'string' ? payload.query : ''
  if (!query.trim()) throw new Error('office_sheet_search needs a non-empty "query"')
  const sheetId = resolveSheetRef(payload.sheetId, handlers)
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
 * When the batch names a default sheet, seed it into the ops that target an area
 * but carry no sheet reference, so `office_sheet_edit(tabId, sheetId, ops)` need
 * not repeat `sheetId` on every op. A sheet-level op (add_sheet, …) is left be —
 * it has no area field to key on.
 */
function seedSheetId(op: unknown, sheetId: string): unknown {
  if (typeof op !== 'object' || op === null || Array.isArray(op)) return op
  const record = { ...(op as Record<string, unknown>) }
  const named = PAYLOAD_FIELDS.some((field) => record[field] !== undefined)
  const targetsArea = AREA_FIELDS.some((field) => typeof record[field] === 'string')
  if (!named && targetsArea) record.sheetId = sheetId
  return record
}

/** Apply a workbook DSL batch to the visible grid, addressed by the named sheet. */
async function sheetEdit(
  payload: SheetEditPayload,
  handlers: McpSheetHandlers,
): Promise<SheetEditResult> {
  const ops = Array.isArray(payload.ops) ? payload.ops : []
  if (ops.length === 0) throw new Error('ops must be a non-empty array')
  const defaultSheetId = resolveSheetRef(payload.sheetId, handlers)
  const seeded = defaultSheetId === undefined ? ops : ops.map((op) => seedSheetId(op, defaultSheetId))
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
  return result as SheetEditResult
}
