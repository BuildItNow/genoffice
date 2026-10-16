import { sheetHelp } from '@genoffice/cli/sheet-catalog'
import { z, type ZodRawShape } from 'zod'
import { OFFICE_PREFIX } from '../names'

/**
 * The daemon's own `office_*` family: the tools that drive a live GenOffice
 * window, but over the control socket rather than the app's built-in MCP server
 * (see `../upstream.ts`). Keeping them off the built-in server is the point —
 * that surface stays exactly as the app ships it.
 *
 * The table is static, so the tools are listed even while the app is down; a
 * call starts the app on demand. Every name is the bare form (`office_list`
 * registers `list`); the exposed name is `OFFICE_PREFIX + name`. The shell handles
 * `name` in its daemon channel (apps/shell/src/main/daemon/daemon-control.ts),
 * which routes a per-family command to the owning tab's renderer (the sheet_*
 * ones land in apps/sheets/src/renderer/daemon/commands.ts).
 *
 * list / create / save / close are family-agnostic: they act on a tab whatever
 * it holds. The per-family tools are grouped beside them (the sheet_* set and
 * the pdf_info / pdf_read / pdf_search reads today), and a new family adds its
 * own `*_info` / `*_read` / `*_edit` set.
 */

export interface OfficeToolSpec {
  /** the bare tool name; the exposed name is `OFFICE_PREFIX + name` */
  name: string
  description: string
  shape: ZodRawShape
  readOnly?: boolean
  /**
   * Answer the call in the daemon instead of forwarding it over the control
   * socket. For tools that are pure static metadata and need no app or tab
   * (e.g. the op reference), so they keep working while the app is closed.
   */
  local?: (args: Record<string, unknown>) => unknown | Promise<unknown>
}

const tabIdField = z.string().describe('the tab to act on, as an id from office_list')
/** the sheet reference addressed by NAME (ids are session artifacts — see apps/sheets mcp-sheet-refs) */
const sheetField = z
  .string()
  .optional()
  .describe('worksheet name; default is the tab\'s active sheet')

export const OFFICE_TOOLS: OfficeToolSpec[] = [
  {
    name: 'list',
    description:
      'List the documents open in GenOffice right now, with the tabId each later call addresses. ' +
      'Returns one row per editor tab: { tabId, kind, title, path, active, dirty }. Use a tabId ' +
      'with the other office_* tools. The list is live — a tab the user opens or closes shows up ' +
      '(or drops out) on the next call.',
    shape: {},
    readOnly: true,
  },
  {
    name: 'create',
    description:
      'Create a new empty document on disk and open it in a visible GenOffice tab. `path` is the ' +
      'absolute folder to create it in (it must exist), `name` the file name without extension ' +
      '(the file must not already exist), and `type` the format to create: "xlsx" (a workbook), ' +
      '"docx" (a document), "pptx" (a presentation) or "pdf" (a blank page). Returns ' +
      '{ tabId, path } (path is the file written); use tabId with the other office_* tools to ' +
      'work on it.',
    shape: {
      path: z.string().describe('absolute folder to create the file in; must already exist'),
      name: z.string().describe('file name without extension, e.g. "Book1" (no path separators)'),
      type: z
        .enum(['xlsx', 'docx', 'pptx', 'pdf'])
        .describe('the file format to create: a workbook, document, presentation or blank PDF'),
    },
  },
  {
    name: 'open',
    description:
      'Open an existing file on disk in a visible GenOffice tab and bring its tab into view. ' +
      '`path` is the absolute path to a workbook (xlsx), document (docx), presentation (pptx) or ' +
      'PDF that already exists (it is opened, never created — use office_create for a new file); a ' +
      'file already open focuses its existing tab instead of a second copy. Returns { tabId, path }, ' +
      'plus `info` — the workbook\'s active sheet and sheet list, the same shape office_sheet_info ' +
      'returns — when the file is a workbook (xlsx), so its sheets need no separate call. Use tabId ' +
      'with the other office_* tools to work on it.',
    shape: {
      path: z.string().describe('absolute path to the file to open; it must already exist'),
    },
  },
  {
    name: 'close',
    description:
      'Close the tab a `tabId` names (see office_list). By default the document is saved on its ' +
      'way out — to its own file, or into the default save folder when it has never been saved — ' +
      'so no unsaved work is lost; pass discard:true to drop the unsaved changes instead. Returns ' +
      '{ tabId, closed: true }, plus savedPath when it saved.',
    shape: {
      tabId: tabIdField,
      discard: z
        .boolean()
        .optional()
        .describe('drop the unsaved changes instead of saving before closing (default false)'),
    },
  },
  {
    name: 'save',
    description:
      "Save an open tab's live content back to its own file (like the app's Save). A tab that has " +
      'never been saved lands in the default save folder. Returns { tabId, savedPath }. The tab ' +
      'stays open, so editing can continue afterwards.',
    shape: { tabId: tabIdField },
  },
  {
    name: 'sheet_info',
    description:
      'List the worksheets of the workbook in one tab and which one is active. Returns ' +
      '{ current, sheets: [{ id, name, rows, columns }], dirty } — `current` the active sheet ' +
      "name, `sheets` the sheet names the other sheet tools address with each one's data extent " +
      '(plus its session `id`), and `dirty`: true when the workbook has edits not yet saved to ' +
      'disk. Call this first to learn them.',
    shape: { tabId: tabIdField },
    readOnly: true,
  },
  {
    name: 'sheet_read',
    description:
      'Read one worksheet of a tab: its name, data extent (rows/columns) and frozen panes. ' +
      'Omit `sheet` to read the sheet the view is currently on. When `range` is given, that A1 ' +
      'area is also returned as a value grid (with formulas keyed by address and the merged ranges ' +
      'touching it); a tall range is paged by rows — pass `maxRows` to bound one call and continue ' +
      'from the range the result names.',
    shape: {
      tabId: tabIdField,
      sheet: sheetField,
      range: z
        .string()
        .optional()
        .describe('A1 area to read, e.g. "A1:D20"; omit for the sheet descriptor alone'),
      maxRows: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('page the range: stop after this many rows and report the range to continue from'),
    },
    readOnly: true,
  },
  {
    name: 'sheet_read_formats',
    description:
      'Read one worksheet of a tab exactly like sheet_read, but each grid cell carries the cell\'s ' +
      'explicit formatting instead of its value (a cell with no set format is null, mirroring an ' +
      'empty value):\n' +
      '- font — bold/italic/underline/strikethrough, family, size, color\n' +
      '- fill — fill color\n' +
      '- border — per edge (top, right, bottom, left), an OOXML line-style name with its color\n' +
      '- alignment — horizontal/vertical, wrap, text rotation, indent\n' +
      '- number — number format\n' +
      'Returns the sheet descriptor (name, data extent, frozen panes) plus `range`: the same ' +
      '{ range, rows, merges, rowsShown, rowsTotal, truncated, note? } shape sheet_read uses for ' +
      'values, minus the formulas key. The area is paged by rows the same way: a tall range is cut ' +
      'to one call and the range to continue from is named — pass `maxRows` to bound one call. ' +
      '`category` picks which of these groups to read (combine several, e.g. ["font","border"]); it ' +
      'defaults to the common ["font","fill","alignment","number"] set — pass ["all"] to include ' +
      'the per-edge borders too. Address the worksheet by name.',
    shape: {
      tabId: tabIdField,
      sheet: sheetField,
      range: z.string().describe('A1 area whose formatting to read, e.g. "A1:D20"'),
      maxRows: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('page the area: stop after this many rows and report the range to continue from'),
      category: z
        .array(z.enum(['font', 'fill', 'border', 'alignment', 'number', 'all']))
        .min(1)
        .default(['font', 'fill', 'alignment', 'number'])
        .describe('which Format Cells groups to read'),
    },
    readOnly: true,
  },
  {
    name: 'sheet_search',
    description:
      'Search inside the workbook of one tab and return every hit addressed the way sheet_read ' +
      'addresses cells — { sheet, address, value, formula? } — so a hit feeds straight into ' +
      'sheet_read. `range` scopes the search to an A1 area; `sheet` picks one worksheet (default: ' +
      'every worksheet). regex treats `query` as a regular expression; lookIn chooses values, ' +
      'formulas or both (default both); matchCase makes the match case-sensitive (default is ' +
      'case-insensitive); maxResults caps the hits (default 100, max 500).',
    shape: {
      tabId: tabIdField,
      sheet: sheetField,
      query: z.string().describe('the text to find'),
      range: z.string().optional().describe('scope the search to an A1 area, e.g. "A1:D20"'),
      regex: z.boolean().optional().describe('treat `query` as a regular expression'),
      lookIn: z
        .enum(['values', 'formulas', 'both'])
        .optional()
        .describe('match values, formulas, or both (default both)'),
      matchCase: z.boolean().optional().describe('match case (default is case-insensitive)'),
      maxResults: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('cap on returned hits (default 100, max 500)'),
    },
    readOnly: true,
  },
  {
    name: 'sheet_edit',
    description:
      'Apply workbook DSL operations to one worksheet of a tab. `ops` is an array of op objects ' +
      'addressed by A1. Do NOT give each op its own `sheetId`: this call\'s `sheet` (default: the ' +
      'active sheet) is filled into every op that omits a sheet reference, so you address the ' +
      'worksheet once, on the call — an op may still override it with its own `sheet`/`sheetId`. ' +
      'Common ops:\n' +
      '- set_cell {address, value} — write a value (a string starting "=" becomes a formula)\n' +
      '- set_formula {address, formula} — write a formula (must start "=")\n' +
      '- set_range {start|range, values} — write a rectangular block of values\n' +
      '- clear_range {range} — clear a range\n' +
      '- fill_range {source, target} — tile a source block across the target (formulas shift)\n' +
      '- copy_range {source, target} — copy a block once (formulas shift)\n' +
      '- format_range {range, format} — set cell formats: bold/italic/fill/font/number format/…\n' +
      '- merge_cells {range} — merge a range\n' +
      '- set_col_width {column, widthPx, count?} — column width in pixels\n' +
      '- set_row_height {row, heightPoints, count?} — row height in points\n' +
      '- set_freeze {rows, columns} — freeze panes (0/0 unfreezes)\n' +
      '- sort_range {range, byColumn, order, hasHeader?} — sort a range\n' +
      'For the other ops and each one\'s exact fields, call office_sheet_edit_help. ' +
      'The target sheet is brought into view so the user sees the edit. One batch applies as one ' +
      'undo step; a failed batch changes nothing — pass dryRun to plan and report without modifying ' +
      'the grid.',
    shape: {
      tabId: tabIdField,
      sheet: sheetField,
      ops: z.array(z.any()).describe('array of workbook DSL op objects (A1-addressed)'),
      dryRun: z
        .boolean()
        .optional()
        .describe('plan the batch and report what would change, without modifying the grid'),
    },
  },
  {
    name: 'sheet_edit_help',
    description:
      'The reference for office_sheet_edit ops. `list` names categories and returns each one\'s op ' +
      'names; `ops` names ops and returns each one\'s field signature and note. Categories:\n' +
      '- content — cell values, formulas, find/replace, sorting\n' +
      '- format — cell formats, conditional formats, data validation\n' +
      '- layout — merges, sizes, hidden rows/columns, panes, filters, page setup, protection\n' +
      '- data — hyperlinks, notes, defined names, Excel tables, pivot tables\n' +
      '- visuals — charts, images, shapes, sparklines\n' +
      '- structure — rows, columns and sheets\n' +
      '- editor-only — accepted by the app, refused headless\n' +
      'Give both to get both; with neither the result is empty. An unknown name comes back ' +
      'under `unknown` with a nearest-match suggestion, so a typo does not fail the call, and an op ' +
      'only the live editor can run (e.g. add_table_row) carries a requiresEditorSession caveat. ' +
      'Pure static reference — needs no open document, so call it before office_sheet_edit.',
    shape: {
      list: z
        .array(z.string())
        .optional()
        .describe('category names to list the ops of, e.g. ["format", "layout"]'),
      ops: z
        .array(z.string())
        .optional()
        .describe('op names to document, e.g. ["format_range", "add_chart"]'),
    },
    readOnly: true,
    local: (args) =>
      sheetHelp({
        ...(Array.isArray(args.list) ? { list: args.list as string[] } : {}),
        ...(Array.isArray(args.ops) ? { ops: args.ops as string[] } : {}),
      }),
  },
  {
    name: 'pdf_info',
    description:
      'Describe a PDF: its total page count and metadata (title/author), plus a preview of the ' +
      "opening pages' text, each clipped — one call tells you what the document is and how big it " +
      'is, addressed either by `tabId` (a PDF open in GenOffice right now, see office_list) or by ' +
      'an absolute `path` to a .pdf on disk that need not be open. A cheap first call before ' +
      'office_pdf_read (for pages whose text you want in full) or office_pdf_search.',
    shape: {
      tabId: z.string().optional().describe('an open PDF tab to describe, as an id from office_list'),
      path: z
        .string()
        .optional()
        .describe('absolute path to a .pdf file; used when no open tab is named'),
    },
    readOnly: true,
  },
  {
    name: 'pdf_read',
    description:
      'Read the text of specific pages of a PDF (whole pages: as many as fit, with `nextPages` ' +
      'naming the rest), addressed either by `tabId` (a PDF open in GenOffice right now, see ' +
      'office_list) or by an absolute `path` to a .pdf on disk that need not be open. `pages` is ' +
      'required (e.g. "3" or "1-5,8"); one call reads at most 20 pages. Text is extracted in ' +
      'content order, so complex multi-column layouts may scramble the reading order; scanned ' +
      'pages have no text layer and come back with empty text and hasTextLayer:false. For an ' +
      'overview of the document first, use office_pdf_info.',
    shape: {
      tabId: z.string().optional().describe('an open PDF tab to read, as an id from office_list'),
      path: z
        .string()
        .optional()
        .describe('absolute path to a .pdf file; used when no open tab is named'),
      pages: z
        .string()
        .describe(
          '1-based pages to read, e.g. "3" or "1-5,8" (at most 20 per call; pages past the end ' +
            'are ignored)',
        ),
    },
    readOnly: true,
  },
  {
    name: 'pdf_search',
    description:
      "Search a PDF's text layer and return every hit addressed the way office_pdf_read " +
      'addresses content — { page, snippet } with 1-based pages — so a hit feeds straight into ' +
      'office_pdf_read { tabId, pages: String(page) }. Address the PDF either by `tabId` (an open ' +
      'PDF, see office_list) or by an absolute `path` that need not be open. `pages` scopes the ' +
      'scan (e.g. "1-5,8"); matchCase makes the match case-sensitive (default is ' +
      'case-insensitive); maxResults caps the hits (default 100, max 500); context is the ' +
      'characters of context around each hit (default 30). Scanned pages have no text layer and ' +
      'never match.',
    shape: {
      tabId: z.string().optional().describe('an open PDF tab to search, as an id from office_list'),
      path: z
        .string()
        .optional()
        .describe('absolute path to a .pdf file; used when no open tab is named'),
      query: z.string().describe('the text to find'),
      pages: z
        .string()
        .optional()
        .describe('1-based pages to scan, e.g. "1-5,8"; default is every page'),
      matchCase: z.boolean().optional().describe('match case (default is case-insensitive)'),
      maxResults: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('cap on returned hits (default 100, max 500)'),
      context: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe('characters of context around each hit (default 30)'),
    },
    readOnly: true,
  },
]

/** exposed name (`office_…`) -> the bare tool name the shell routes on */
export const OFFICE_TOOL_NAMES: ReadonlyMap<string, string> = new Map(
  OFFICE_TOOLS.map((tool) => [`${OFFICE_PREFIX}${tool.name}`, tool.name]),
)
