import { z } from 'zod'
import { createFileViaCli } from '../headless-cli'
import { resolveOutputPath } from './document-tools'
import { generateExtension } from './formats'
import type { FamilyDriver, SessionHost, TargetResolver } from './session-tools'
import { contentTarget } from './session-tools'
import type { McpToolDefinition } from '../mcp-server'
import type { CliRunner } from '../cli-runner'

/**
 * Sheets (xlsx) tool surface for the MCP server.
 *
 * Two paths, mirroring the docx/slides tools:
 * - headless `create_xlsx`: a row matrix handed to the bundled `genoffice` CLI
 *   (`create --type xlsx --from`), gated behind the "background generation"
 *   setting. The CLI writes through the app's xlsx gateway and evaluates
 *   formulas with the sidecar, so strings starting with "=" become real formula
 *   cells with cached values — more than the values-only writer this used to
 *   carry.
 * - a visible grid session: the tools drive a real sheets tab the user
 *   watches. The workbook lives in the renderer (Univer), so this needs a
 *   request/response bridge into the renderer (`sheets-bridge.ts` here, the
 *   renderer half at `apps/sheets/src/renderer/mcp-bridge.ts`) — the same
 *   pattern as the docs session.
 */

export interface SheetsToolDeps {
  /** directory generated files land in when the caller gives no path */
  defaultSaveDir: () => string
  /** expose the headless create_xlsx tool; default true — the shell passes the user's setting */
  background?: boolean
  /** visible-grid control; absent in headless/unit runs, which drops the session tools */
  sheets?: SheetsControl
  /** the bundled genoffice CLI, used by the headless tool */
  cli?: CliRunner
  /**
   * Resolve a caller-supplied document reference (tab id or path) to the
   * webContents of that open tab, so the grid tools can edit a workbook the
   * *user* has open rather than only the session's own blank one. Absent in
   * headless/unit runs, where a `document` argument is refused.
   */
  resolveTarget?: TargetResolver
}

/**
 * Visible-grid control implemented in the shell main process
 * (`apps/shell/src/main/mcp/sheets-bridge.ts`): opens a blank sheets tab and
 * forwards commands to the renderer that owns the Univer workbook.
 */
export interface SheetsControl {
  /** open a fresh blank sheets tab; resolves to its webContents id once the renderer is ready */
  openBlankTab: () => Promise<number>
  /** run one command in that tab and resolve its result */
  runCommand: (
    wcId: number,
    command: 'apply_ops' | 'read_sheet' | 'search_sheet' | 'save_sheet',
    payload: unknown,
  ) => Promise<unknown>
}

const XLSX_EXT = `.${generateExtension('xlsx')}`
// same ceiling as the built-in AI's read_cells
const MAX_READ_ADDRESSES = 100

/** the headless tool: a row matrix -> xlsx through the bundled CLI */
function createHeadlessXlsxTool(deps: SheetsToolDeps): McpToolDefinition {
  return {
    name: 'create_xlsx',
    description:
      'Create an Excel .xlsx file and save it to disk without opening the app UI. `data` is a 2D ' +
      'array of rows; numbers become numeric cells, and a string starting with "=" becomes a ' +
      'formula (the genoffice CLI evaluates it and stores the cached value). Returns the absolute ' +
      'path of the written file.',
    inputSchema: {
      title: z.string().describe('workbook title, used as the file name'),
      data: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe('rows of cell values; row 0 becomes spreadsheet row 1'),
      sheetName: z.string().optional().describe('name of the single sheet; default Sheet1'),
      path: z
        .string()
        .optional()
        .describe('absolute output path; default is a new file in the default save folder'),
      overwrite: z
        .boolean()
        .optional()
        .describe('allow replacing an existing file at `path`; default false'),
    },
    handler: async (args) => {
      const title = String(args.title ?? '').trim()
      if (!title) throw new Error('title must not be empty')
      if (!Array.isArray(args.data)) throw new Error('data must be a 2D array of rows')
      if (args.data.some((row) => !Array.isArray(row))) {
        throw new Error('data must be a 2D array of rows')
      }
      if (!deps.cli) throw new Error('headless generation is not available in this build')
      const sheetName =
        typeof args.sheetName === 'string' && args.sheetName.trim()
          ? args.sheetName.trim()
          : 'Sheet1'

      const targetPath = resolveOutputPath({
        defaultSaveDir: deps.defaultSaveDir,
        ext: XLSX_EXT,
        title,
        requestedPath: typeof args.path === 'string' ? args.path : undefined,
        overwrite: args.overwrite === true,
      })

      const rows = (args.data as unknown[][]).map((row) =>
        row.map((cell) => {
          if (cell === null || cell === undefined) return ''
          if (typeof cell === 'number') return Number.isFinite(cell) ? cell : ''
          return String(cell)
        }),
      )
      const { summary, outputPath } = await createFileViaCli(deps.cli, {
        type: 'xlsx',
        input: {
          name: 'table.json',
          content: JSON.stringify({ sheets: [{ name: sheetName, rows }] }),
        },
        out: targetPath,
        overwrite: args.overwrite === true,
      })
      return { path: outputPath, summary }
    },
  }
}

export function createSheetsTools(deps: SheetsToolDeps, host: SessionHost): McpToolDefinition[] {
  return [
    // headless generation is opt-in, same rule as create_docx/create_pptx
    ...(deps.background === false ? [] : [createHeadlessXlsxTool(deps)]),
    ...createGridContentTools(deps, host),
  ]
}

/**
 * The xlsx session lifecycle as seen by the shared create_session / save_session
 * tools. The content tools below address the tab this driver opened.
 */
export function sheetsDriver(sheets: SheetsControl): FamilyDriver {
  return {
    family: 'xlsx',
    openBlankTab: () => sheets.openBlankTab(),
    save: (wcId, path, overwrite) => sheets.runCommand(wcId, 'save_sheet', { path, overwrite }),
  }
}

/**
 * Visible-grid content tools: fill the spreadsheet the shared session opened, so
 * the user watches the grid take shape. The ops are the same zod-validated
 * workbook DSL the built-in AI uses (planFromOps + applyChangePlan), executed by
 * the tab's renderer.
 *
 * Only registered when the shell wired a SheetsControl — headless/unit runs
 * keep the file-only surface.
 */
function createGridContentTools(deps: SheetsToolDeps, host: SessionHost): McpToolDefinition[] {
  const sheets = deps.sheets
  if (!sheets) return []

  const target = contentTarget(host, deps.resolveTarget, 'xlsx')
  const documentField = z
    .string()
    .optional()
    .describe(
      "tab id or path of an open spreadsheet to edit instead of the session's own tab " +
        '(from open_documents {action:"list"}); the edit lands in that live document, and a ' +
        'mutating tool switches the UI to that tab first so the user sees it happen',
    )

  return [
    {
      name: 'read_sheet',
      description:
        'Read the visible workbook. By default an overview: sheet names with data extents. `range` ' +
        'reads a rectangular A1 area (e.g. "A1:D20") as a value grid with formulas keyed by address — ' +
        'the same read as the genoffice CLI `sheet read --range`; a grid too large for one call is ' +
        'paged by rows and the result names the `range` to continue from, so paging the pages reads the ' +
        'whole area. `addresses` reads up to 100 scattered cells instead. Address a worksheet by NAME ' +
        '(the CLI convention, stable across sessions) or by the sheetId the overview reports.',
      inputSchema: {
        range: z
          .string()
          .optional()
          .describe('A1 range to read, e.g. "A1:D20" or "B2"; returns a value grid with formulas'),
        maxRows: z
          .number()
          .int()
          .positive()
          .optional()
          .describe(
            'page the range: stop after this many rows; the result names the range to continue from',
          ),
        addresses: z
          .array(z.string())
          .optional()
          .describe(
            'A1 addresses to read (values + formulas, at most 100 per call); use `range` for an area',
          ),
        sheet: z
          .string()
          .optional()
          .describe('worksheet name to read from; default is the active sheet'),
        sheetId: z
          .string()
          .optional()
          .describe(
            'worksheet id (from a previous overview) as an alternative to `sheet`; a name is preferred',
          ),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document)
        const range = typeof args.range === 'string' ? args.range.trim() : ''
        if (range && Array.isArray(args.addresses) && args.addresses.length > 0) {
          throw new Error('give either `range` or `addresses`, not both')
        }
        if (Array.isArray(args.addresses) && args.addresses.length > MAX_READ_ADDRESSES) {
          throw new Error(`read at most ${MAX_READ_ADDRESSES} addresses per call`)
        }
        const payload = {
          ...(range ? { range } : {}),
          ...(typeof args.maxRows === 'number' ? { maxRows: args.maxRows } : {}),
          ...(Array.isArray(args.addresses) ? { addresses: args.addresses.map(String) } : {}),
          ...(typeof args.sheet === 'string' ? { sheet: args.sheet } : {}),
          ...(typeof args.sheetId === 'string' ? { sheetId: args.sheetId } : {}),
        }
        return sheets.runCommand(wc, 'read_sheet', payload)
      },
    },
    {
      name: 'search_sheet',
      description:
        'Search inside the visible workbook and return every hit, addressed the way read_sheet ' +
        'addresses cells — {sheet, address, value, formula?} — so a hit feeds straight into ' +
        'read_sheet { addresses: [address] }. `range` scopes the search to an A1 area ("A1:D20"); ' +
        'sheet/sheetId pick one worksheet. regex treats `query` as a regular expression; lookIn ' +
        'chooses values, formulas or both (default both). Matching is case-insensitive. maxResults ' +
        'caps the returned hits (default 100, max 500).',
      inputSchema: {
        document: documentField,
        query: z.string().describe('the text to find'),
        range: z
          .string()
          .optional()
          .describe('scope the search to an A1 area, e.g. "A1:D20"; default is the whole sheet'),
        sheet: z
          .string()
          .optional()
          .describe('worksheet name to search; default is every worksheet in the workbook'),
        sheetId: z
          .string()
          .optional()
          .describe('worksheet id (from read_sheet) as an alternative to `sheet`'),
        regex: z
          .boolean()
          .optional()
          .describe('treat `query` as a regular expression instead of literal text'),
        lookIn: z
          .enum(['values', 'formulas', 'both'])
          .optional()
          .describe('match values, formulas, or both (default both)'),
        maxResults: z
          .number()
          .int()
          .positive()
          .optional()
          .describe('cap on returned hits (default 100, max 500)'),
      },
      handler: async (args) => {
        const query = String(args.query ?? '')
        if (!query) throw new Error('query must not be empty')
        const wc = await target(args.document)
        return sheets.runCommand(wc, 'search_sheet', {
          query,
          ...(typeof args.range === 'string' && args.range.trim()
            ? { range: args.range.trim() }
            : {}),
          ...(typeof args.sheet === 'string' ? { sheet: args.sheet } : {}),
          ...(typeof args.sheetId === 'string' ? { sheetId: args.sheetId } : {}),
          regex: args.regex === true,
          ...(args.lookIn === 'values' || args.lookIn === 'formulas'
            ? { lookIn: args.lookIn }
            : {}),
          maxResults: typeof args.maxResults === 'number' ? args.maxResults : undefined,
        })
      },
    },
    {
      name: 'apply_sheet_ops',
      description:
        'Apply workbook DSL operations to the visible spreadsheet (the same vocabulary the CLI ' +
        'and the built-in AI use). Common ops: set_cell, set_formula, clear_cell, set_range, ' +
        'clear_range, fill_range, copy_range, convert_to_values, insert_rows, delete_rows, ' +
        'insert_cols, delete_cols, add_sheet, delete_sheet, add_chart, add_table, set_filter, ' +
        'set_hyperlink, add_conditional_format, set_data_validation, set_note. Address a worksheet ' +
        "by NAME with `sheet` (also `sourceSheet`/`targetSheet`) — the CLI's stable vocabulary — " +
        'or with `sheetId` from read_sheet; a name that is not in the workbook is rejected before ' +
        'anything applies. The target sheet is brought into view so the user sees the edit. ' +
        'Addresses are A1. One batch applies as one undo step; a failed batch changes nothing.',
      inputSchema: {
        ops: z
          .array(z.any())
          .describe(
            'array of workbook DSL op objects; each op addresses its worksheet with sheet (a name) or sheetId',
          ),
        dryRun: z
          .boolean()
          .optional()
          .describe('plan the batch and report what would change, without modifying the grid'),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document, { focus: true })
        if (!Array.isArray(args.ops)) throw new Error('ops must be an array')
        const result = (await sheets.runCommand(wc, 'apply_ops', {
          ops: args.ops,
          dryRun: args.dryRun === true,
        })) as { ok?: boolean; reason?: string }
        // a rejected batch is a tool-level error so the caller reacts to it
        if (result?.ok === false) {
          throw new Error(result.reason ?? 'the batch could not be applied')
        }
        return result
      },
    },
  ]
}
