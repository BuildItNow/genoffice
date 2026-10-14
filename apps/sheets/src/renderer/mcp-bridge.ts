import { z } from 'zod'
import type { McpCommandMessage } from '../shared/desktop-api'
import {
  describeOperationErrors,
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  parseAddress,
  parseRange,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { normalizeSheetRefs, primaryCellOf, primarySheetId, type SheetRef } from './mcp-sheet-refs'
import type { FindCellsOptions, FindCellsOutcome } from './ai/tools'

/**
 * Renderer half of the MCP → sheets-grid bridge.
 *
 * The shell main process pushes one command at a time (`sheets:mcp-command`);
 * each command runs against the live Univer workbook through the same executors
 * the built-in AI uses (readers from ai/workbook-readers.ts, the op planner and
 * apply path from op-executor.ts, the save pipeline from save-actions.ts), so
 * external edits land in the edit journal and undo history exactly like
 * in-app/AI ones. Results are reported back correlated by requestId
 * (`sheets:mcp-result`), and once the workbook is mounted the bridge announces
 * `sheets:mcp-ready` — the shell waits for it before accepting a session.
 *
 * No-op when the preload lacks the bridge channels (standalone/dev builds
 * without MCP wiring).
 */

const READY_POLL_MS = 50
/** after the fast window the poll backs off instead of stopping (see below) */
const READY_SLOW_POLL_MS = 1_000
const READY_FAST_WINDOW_MS = 20_000

/** search_sheet match-count bounds, mirroring the MCP search tool's own */
const SEARCH_RESULTS_DEFAULT = 100
const SEARCH_MAX_RESULTS = 500

/** does an A1 address fall inside an inclusive range? (malformed address → false) */
function addressInBounds(address: string, bounds: RangeBounds): boolean {
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

export interface McpSheetHandlers {
  /** workbook mounted and editable? (drives the ready announce) */
  hasWorkbook: () => boolean
  /** workbook overview: sheets with ids/names/extents, active sheet */
  context: () => unknown
  /** values + formulas for the requested A1 addresses */
  readCells: (addresses: string[], sheetId?: string) => unknown
  /** read a rectangular A1 range as a value grid + formulas keyed by address, paged by rows (the CLI's `sheet read --range`) */
  readRange: (
    range: string,
    sheetId: string | undefined,
    opts: { maxRows?: number },
  ) => unknown | Promise<unknown>
  /** find cells across the workbook (the in-app find_cells engine); hits carry {sheetName,address,value,formula?} */
  findCells: (options: FindCellsOptions) => FindCellsOutcome | Promise<FindCellsOutcome>
  /** the workbook's worksheets, for resolving name-addressed ops */
  sheets: () => readonly SheetRef[]
  /** apply a workbook DSL batch (planFromOps + applyChangePlan) */
  applyOps: (ops: WorkbookOperation[], dryRun: boolean) => Promise<unknown>
  /** bring a sheet (and optionally a cell) into view before/after an edit */
  focusSheet: (sheetId: string, address?: string) => void
  /** save to an explicit absolute path (dialog-free) */
  saveTo: (
    path: string,
    overwrite: boolean,
  ) => Promise<{ ok: boolean; path?: string; error?: string }>
}

export function installSheetsMcpBridge(handlers: McpSheetHandlers): () => void {
  const api = window.desktopApi
  if (typeof api?.onMcpCommand !== 'function') return () => {}

  /** one command at a time, in arrival order — applies and saves must not interleave */
  let queue: Promise<void> = Promise.resolve()
  let disposed = false

  // Announce readiness as soon as the workbook is mounted, then stop polling.
  //
  // The poll must never give up permanently: a cold dev start (vite serving the
  // Univer bundle, the sidecar opening the workbook) can take longer than the
  // fast window, and once the timer has stopped the tab stays unaddressable
  // over MCP for the rest of its life — the grid looks perfectly normal in the
  // UI while every sheet command times out in the shell. Back off instead.
  let readyTimer: ReturnType<typeof setTimeout> | null = null
  const mountedAt = performance.now()
  const announceWhenMounted = (): void => {
    readyTimer = null
    if (disposed) return
    if (handlers.hasWorkbook()) {
      api.signalMcpReady()
      return
    }
    const slow = performance.now() - mountedAt > READY_FAST_WINDOW_MS
    readyTimer = setTimeout(announceWhenMounted, slow ? READY_SLOW_POLL_MS : READY_POLL_MS)
  }
  readyTimer = setTimeout(announceWhenMounted, READY_POLL_MS)

  const off = api.onMcpCommand((message: McpCommandMessage) => {
    queue = queue.then(() => runCommand(message)).catch(() => undefined)
  })

  return () => {
    disposed = true
    if (readyTimer !== null) clearTimeout(readyTimer)
    off()
  }

  /** First schema issue, with the offending op named so the client can fix it. */
  function describeOpError(ops: unknown[], error: z.ZodError): string {
    const text = describeOperationErrors(ops, error)
    return error.issues.some((issue) => issue.path.includes('sheetId'))
      ? `${text}\nCall read_sheet first and use a sheetId from its output.`
      : text
  }

  /**
   * A read's target sheet: the name the caller gave (`sheet`, the CLI's
   * vocabulary) resolved against the open workbook, else the id it passed.
   * Ids are session artifacts, so a name that no longer matches fails loudly
   * instead of reading the wrong sheet.
   */
  function resolveReadSheet(
    name: string | undefined,
    id: string | undefined,
  ): { ok: true; sheetId?: string } | { ok: false; error: string } {
    if (name === undefined) return { ok: true, ...(id === undefined ? {} : { sheetId: id }) }
    const sheets = handlers.sheets()
    if (sheets.some((sheet) => sheet.id === name)) return { ok: true, sheetId: name }
    const match = sheets.find(
      (sheet) => sheet.name.trim().toLowerCase() === name.trim().toLowerCase(),
    )
    if (match) return { ok: true, sheetId: match.id }
    const known = sheets.length === 0 ? 'none' : sheets.map((sheet) => sheet.name).join(', ')
    return {
      ok: false,
      error: `no worksheet named "${name}" in this workbook (sheets: ${known}); call read_sheet for the current sheets`,
    }
  }

  async function runCommand(message: McpCommandMessage): Promise<void> {
    const reply = (ok: boolean, result?: unknown, error?: string): void => {
      api.reportMcpResult({
        requestId: message.requestId,
        ok,
        ...(result !== undefined ? { result } : {}),
        ...(error !== undefined ? { error } : {}),
      })
    }
    try {
      const payload = (message.payload ?? {}) as Record<string, unknown>
      switch (message.command) {
        case 'read_sheet': {
          const addresses = Array.isArray(payload.addresses)
            ? payload.addresses.filter((a): a is string => typeof a === 'string')
            : []
          const resolved = resolveReadSheet(
            typeof payload.sheet === 'string' ? payload.sheet : undefined,
            typeof payload.sheetId === 'string' ? payload.sheetId : undefined,
          )
          if (!resolved.ok) {
            reply(false, undefined, resolved.error)
            return
          }
          // A named A1 area takes the range read (the CLI's `sheet read --range`),
          // which pages a tall grid; the flat address list and the overview follow.
          const range = typeof payload.range === 'string' ? payload.range.trim() : ''
          if (range) {
            reply(
              true,
              await handlers.readRange(range, resolved.sheetId, {
                ...(typeof payload.maxRows === 'number' ? { maxRows: payload.maxRows } : {}),
              }),
            )
            return
          }
          if (addresses.length > 0) {
            reply(true, { cells: handlers.readCells(addresses, resolved.sheetId) })
          } else {
            reply(true, { context: handlers.context() })
          }
          return
        }
        case 'search_sheet': {
          const query = typeof payload.query === 'string' ? payload.query : ''
          if (!query.trim()) {
            reply(false, undefined, 'search_sheet needs a non-empty "query"')
            return
          }
          const resolved = resolveReadSheet(
            typeof payload.sheet === 'string' ? payload.sheet : undefined,
            typeof payload.sheetId === 'string' ? payload.sheetId : undefined,
          )
          if (!resolved.ok) {
            reply(false, undefined, resolved.error)
            return
          }
          // A range narrows the hits to one A1 area; the engine scans a sheet
          // (or the whole workbook), so the area filter is applied afterwards.
          const range = typeof payload.range === 'string' ? payload.range.trim() : ''
          let bounds: RangeBounds | undefined
          if (range) {
            try {
              bounds = parseRange(range)
            } catch (error) {
              reply(
                false,
                undefined,
                error instanceof Error ? error.message : `invalid range "${range}"`,
              )
              return
            }
          }
          const lookInRaw = payload.lookIn
          const maxRaw = Number(payload.maxResults)
          const outcome = await handlers.findCells({
            query,
            regex: payload.regex === true,
            lookIn: lookInRaw === 'values' || lookInRaw === 'formulas' ? lookInRaw : 'both',
            errorsOnly: false,
            sheetId: resolved.sheetId,
            maxResults:
              Number.isFinite(maxRaw) && maxRaw >= 1
                ? Math.min(Math.trunc(maxRaw), SEARCH_MAX_RESULTS)
                : SEARCH_RESULTS_DEFAULT,
          })
          if (outcome.error) {
            reply(false, undefined, outcome.error)
            return
          }
          const kept = bounds
            ? outcome.matches.filter((match) => addressInBounds(match.address, bounds))
            : outcome.matches
          reply(true, {
            total: kept.length,
            matches: kept.map((match) => ({
              sheet: match.sheetName,
              address: match.address,
              value: match.value,
              ...(match.formula === undefined ? {} : { formula: match.formula }),
            })),
            truncated: outcome.truncated,
            ...(outcome.incompleteSheets.length
              ? { incompleteSheets: outcome.incompleteSheets }
              : {}),
          })
          return
        }
        case 'apply_ops': {
          const ops = Array.isArray(payload.ops) ? payload.ops : []
          if (ops.length === 0) {
            reply(false, undefined, 'ops must be a non-empty array')
            return
          }
          // Worksheet names are the stable reference (the CLI's vocabulary);
          // ids remain accepted. Resolved here, in the renderer, because only
          // this process knows the open workbook's sheets.
          const named = normalizeSheetRefs(ops, handlers.sheets())
          if (!named.ok) {
            reply(false, undefined, named.error)
            return
          }
          // Same validation the built-in AI path applies (ai/tools.ts): without
          // it a missing sheetId surfaces later as a cryptic
          // "Unknown sheet: undefined" from the planner instead of naming the op.
          const parsed = z.array(workbookOperationSchema).safeParse(named.ops)
          if (!parsed.success) {
            reply(false, undefined, describeOpError(named.ops, parsed.error))
            return
          }
          const sheetId = primarySheetId(parsed.data)
          const dryRun = payload.dryRun === true
          if (sheetId !== undefined && !dryRun)
            handlers.focusSheet(sheetId, primaryCellOf(parsed.data))
          reply(true, await handlers.applyOps(parsed.data, dryRun))
          return
        }
        case 'save_sheet': {
          const path = typeof payload.path === 'string' ? payload.path : ''
          const overwrite = payload.overwrite === true
          if (!path) {
            reply(false, undefined, 'save_sheet needs an absolute path')
            return
          }
          const saved = await handlers.saveTo(path, overwrite)
          // a refused save must surface as an error so the agent keeps its session and can retry
          if (!saved.ok) {
            reply(false, undefined, saved.error ?? 'the spreadsheet could not be saved')
            return
          }
          reply(true, saved)
          return
        }
        default:
          reply(false, undefined, `unknown command: ${String(message.command)}`)
      }
    } catch (error: unknown) {
      reply(false, undefined, error instanceof Error ? error.message : String(error))
    }
  }
}
