import { z, type ZodRawShape } from 'zod'
import { resolveReadPdfPath } from './document-ref'
import type { AppUpstream } from './upstream'

/**
 * The GUI tool family: everything that can only run *inside* a live GenOffice
 * window — the visible editing sessions, the documents the user has open, and
 * opening a file in a tab. The daemon does not reimplement any of it: each tool
 * is a thin proxy to the app's own built-in MCP server (see upstream.ts), which
 * the daemon starts on demand.
 *
 * The table is deliberately static so the tools are always listed, even while
 * the app is down. It mirrors the app's *default* surface —
 * apps/shell/src/main/mcp/tools/*.ts — with the headless `create_*` generators
 * left out, because the daemon already serves those as `office_cli_create_*`.
 *
 * `read_pdf` stays, but reshaped: the app addresses it by file path, yet every
 * other read in this family addresses an open document. A PDF has no live editor
 * state (the pdf app is a viewer), so `office_gui_read_pdf` takes the same
 * `document` reference as its siblings and its `adapt` resolves that to the file
 * before the call goes upstream — the family stays uniform, and the read is
 * still just a file read.
 *
 * A `apps/shell/tests/mcp/daemon-parity.test.ts` guards the names against drift.
 */

const documentField = z
  .string()
  .optional()
  .describe(
    "tab id or path of an open document to act on instead of the active session's tab " +
      '(from office_gui_opened_documents {action:"list"}); the edit lands in that live document, ' +
      'and a mutating tool switches the UI to that tab first',
  )

/**
 * `office_gui_read_pdf`'s target. Unlike the other reads it is required: a PDF
 * has no session, so there is no "active tab" to fall back to.
 */
const pdfDocumentField = z
  .string()
  .describe(
    'tab id or path of an open PDF document ' +
      '(from office_gui_opened_documents {action:"list"})',
  )

export interface GuiToolSpec {
  /** the name registered on the app's own MCP server */
  upstream: string
  /** the name this daemon exposes (`office_gui_…`) */
  name: string
  description: string
  shape: ZodRawShape
  readOnly?: boolean
  /**
   * Rewrite the daemon's arguments into the upstream call, when the exposed
   * shape differs from the app's. Default: the arguments are forwarded as-is.
   * Only `office_gui_read_pdf` needs it — its `document` names an open PDF that
   * must be resolved to a file path before the app's `read_pdf` is called.
   */
  adapt?: (args: Record<string, unknown>, upstream: AppUpstream) => Promise<Record<string, unknown>>
}

export const GUI_TOOLS: GuiToolSpec[] = [
  {
    name: 'office_gui_create_session',
    upstream: 'create_session',
    description:
      'Open a new empty file in a visible GenOffice tab and start an editing session; family picks ' +
      "the editor: docx, xlsx or pptx. Then add content with that family's tools (insert_content / " +
      'replace_blocks / apply_ops; apply_slide_ops; apply_sheet_ops) and finish with save_session. ' +
      'One session is active at a time — a new create_session replaces it. The user watches.',
    shape: {
      family: z.enum(['docx', 'xlsx', 'pptx']).describe('which editor to open'),
    },
  },
  {
    name: 'office_gui_save_session',
    upstream: 'save_session',
    description:
      'Save the active session to an absolute path and end it. A path without an extension gets the ' +
      "active family's format appended; a different extension is refused. Refuses to replace an " +
      'existing file unless overwrite is true.',
    shape: {
      path: z.string().describe('absolute output path for the file'),
      overwrite: z.boolean().optional().describe('allow replacing an existing file; default false'),
    },
  },
  {
    name: 'office_gui_opened_documents',
    upstream: 'open_documents',
    description:
      'Work with the documents the user currently has open in GenOffice tabs. Actions: "list" ' +
      '(no target) returns every open document with id, type, path, title, saved and dirty; ' +
      '"read" returns one document\'s live content including unsaved edits; "save" writes that ' +
      'live content back to the file and leaves the tab open; "close" closes one, saving first by ' +
      'default (pass unsaved:"discard" to drop the changes). Identify a document by path or by the ' +
      'id from "list". To edit one, pass its id/path as the `document` argument of a content tool.',
    shape: {
      action: z.enum(['list', 'read', 'save', 'close']).describe('what to do'),
      target: z
        .string()
        .optional()
        .describe('path or tab id; required for "read", "save" and "close"'),
      unsaved: z
        .enum(['save', 'discard'])
        .optional()
        .describe('"close" only: save before closing (default), or discard the changes'),
    },
  },
  {
    name: 'office_gui_open',
    upstream: 'open_in_genoffice',
    description:
      'Open an existing file in the running GenOffice app, focusing its tab. Only when the user ' +
      'asks to see the file: an open tab makes later writes on that path refuse without force.',
    shape: {
      path: z.string().describe('absolute path to the file to open'),
    },
  },
  {
    name: 'office_gui_get_app_info',
    upstream: 'get_app_info',
    readOnly: true,
    description:
      'Report GenOffice version, the default save folder, the document formats the app can ' +
      "generate, and the editor's open/save/export format matrix per family.",
    shape: {},
  },
  {
    name: 'office_gui_read_document',
    upstream: 'read_document',
    readOnly: true,
    description:
      'Read the visible document. By default an overview: a block list (one "index|type|text" line per ' +
      'block, where the index is what apply_ops and insert_content address) plus text stats and the ' +
      'current selection; long blocks are clipped. Pass `range` (0-based block indexes, e.g. "3" or ' +
      '"3-10") and/or `full` to read the full text as restricted HTML — a long range is paged, and the ' +
      'result names the `offset` to continue from. `maxChars` moves the page size.',
    shape: {
      document: documentField,
      range: z
        .string()
        .optional()
        .describe('0-based block index range to read, e.g. "3" or "3-10"; default is every block'),
      full: z
        .boolean()
        .optional()
        .describe('read the whole range in one page instead of paging it'),
      maxChars: z
        .number()
        .int()
        .optional()
        .describe('page size in characters (default 24000); ignored when full is true'),
      offset: z
        .number()
        .int()
        .optional()
        .describe('character offset to continue a paged read (default 0)'),
    },
  },
  {
    name: 'office_gui_insert_content',
    upstream: 'insert_content',
    description:
      'Insert content into the visible document as HTML (headings, paragraphs, bold/italic, lists, ' +
      'tables, links). Appends at the end unless afterBlockIndex is given.',
    shape: {
      html: z.string().describe('restricted HTML fragment to insert'),
      afterBlockIndex: z
        .number()
        .int()
        .optional()
        .describe('insert after this block index; default appends at the end'),
      document: documentField,
    },
  },
  {
    name: 'office_gui_replace_blocks',
    upstream: 'replace_blocks',
    description:
      'Replace a range of blocks in the visible document with new HTML content. Use read_document ' +
      'to learn block indexes.',
    shape: {
      startBlockIndex: z.number().int().describe('first block index to replace (inclusive)'),
      endBlockIndex: z.number().int().describe('last block index to replace (inclusive)'),
      html: z.string().describe('restricted HTML fragment the range is replaced with'),
      document: documentField,
    },
  },
  {
    name: 'office_gui_apply_ops',
    upstream: 'apply_ops',
    description:
      'Apply formatting commands to the visible document (font, paragraph format, heading level, ' +
      'find/replace, list/indent). `ops` is the same batch the built-in AI editor accepts; the ' +
      'whole batch is atomic. Use read_document for block indexes.',
    shape: {
      ops: z.array(z.any()).describe('array of op objects'),
      dryRun: z.boolean().optional().describe('plan the batch without changing the document'),
      document: documentField,
    },
  },
  {
    name: 'office_gui_read_sheet',
    upstream: 'read_sheet',
    readOnly: true,
    description:
      'Read the visible workbook. By default an overview: sheet names with data extents. `range` reads ' +
      'a rectangular A1 area (e.g. "A1:D20") as a value grid with formulas keyed by address — the same ' +
      'read as the CLI `sheet read --range`; a tall grid is paged by rows and the result names the ' +
      'range to continue from. `addresses` reads up to 100 scattered cells instead.',
    shape: {
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
        .describe('A1 addresses to read (values + formulas, at most 100); use `range` for an area'),
      sheet: z.string().optional().describe('worksheet name; default is the active sheet'),
      sheetId: z.string().optional().describe('worksheet id as an alternative to `sheet`'),
      document: documentField,
    },
  },
  {
    name: 'office_gui_apply_sheet_ops',
    upstream: 'apply_sheet_ops',
    description:
      'Apply workbook DSL operations to the visible spreadsheet (set_cell, set_formula, set_range, ' +
      'insert_rows, add_sheet, add_chart, add_table, set_filter, set_data_validation, …). One batch ' +
      'applies as one undo step; a failed batch changes nothing.',
    shape: {
      ops: z.array(z.any()).describe('array of workbook DSL op objects'),
      dryRun: z.boolean().optional().describe('plan the batch without modifying the grid'),
      document: documentField,
    },
  },
  {
    name: 'office_gui_read_deck',
    upstream: 'read_deck',
    readOnly: true,
    description:
      'Read the visible presentation: every slide with its elements (ids, types, text, geometry). ' +
      'Ids and slide indexes are what apply_slide_ops addresses; geometry is document-space EMU. ' +
      'Text and table rows are previewed (clipped) by default; pass full:true for whole text and ' +
      'every table row, maxChars to move the clip, slide to read one slide (0-based).',
    shape: {
      document: documentField,
      slide: z
        .number()
        .int()
        .optional()
        .describe('0-based slide index to read; default every slide'),
      full: z
        .boolean()
        .optional()
        .describe('return whole text and every table row (no preview clipping)'),
      maxChars: z
        .number()
        .int()
        .optional()
        .describe('clip text to this many characters; default 300; ignored when full is true'),
    },
  },
  {
    name: 'office_gui_apply_slide_ops',
    upstream: 'apply_slide_ops',
    description:
      'Apply slide editing ops to the visible presentation as one transaction (atomic by default). ' +
      'Ops address {slide, el} in document-space EMU (setText, addElement, addTable, deleteSlide, ' +
      'setTransform, applyTheme, setNotes, …). Use read_deck for ids and geometry first.',
    shape: {
      ops: z.array(z.any()).describe('array of op objects (at most 50 per transaction)').max(50),
      isolation: z
        .enum(['atomic', 'per_op'])
        .optional()
        .describe('atomic (default) applies all-or-nothing; per_op applies independently'),
      dryRun: z.boolean().optional().describe('plan the batch without changing the deck'),
      document: documentField,
    },
  },
  {
    name: 'office_gui_read_pdf',
    upstream: 'read_pdf',
    readOnly: true,
    description:
      'Read an open PDF — the same shape as the other visible-document reads, but a PDF has no ' +
      'live editor state (the pdf app is a viewer), so this reads its file. Pass the PDF as ' +
      '`document` (its tab id or path). By default a preview: page count, metadata and the first ' +
      'few pages\' text, each clipped with a `…(+N chars)` marker. Pass `pages` (e.g. "1-5,8") to ' +
      'read those pages in full — one call reads at most 20 pages and the result names the ' +
      '`nextPages` to continue from (pages past the end are ignored); scanned pages come back with ' +
      'empty text and hasTextLayer:false.',
    shape: {
      document: pdfDocumentField,
      pages: z
        .string()
        .optional()
        .describe(
          '1-based pages to read, e.g. "3" or "1-5,8" (at most 20 per call; pages past the end ' +
            'are ignored); omit for the preview (first pages, clipped)',
        ),
    },
    adapt: async (args, upstream) => ({
      path: await resolveReadPdfPath(upstream, args.document),
      pages: args.pages,
    }),
  },
  {
    name: 'office_gui_search_document',
    upstream: 'search_document',
    readOnly: true,
    description:
      'Search inside a Word document for text and return every hit, addressed the way ' +
      'office_gui_read_document addresses content — {block, snippet} — ' +
      'so a hit feeds straight into office_gui_read_document { range }. Pass `document` to name an ' +
      'open document (from office_gui_opened_documents {action:"list"}), or omit it to search the ' +
      "active session's document. `range` scopes the search to a 0-based block index range; " +
      'matchCase, wholeWord and regex tune the match; maxResults caps the hits (default 100).',
    shape: {
      document: documentField,
      query: z.string().describe('the text to find'),
      range: z
        .string()
        .optional()
        .describe('scope: a 0-based block index range, e.g. "3" or "3-10"'),
      matchCase: z.boolean().optional().describe('match case (default is case-insensitive)'),
      wholeWord: z.boolean().optional().describe('match whole words only'),
      regex: z
        .boolean()
        .optional()
        .describe('treat `query` as a regular expression instead of literal text'),
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
  },
  {
    name: 'office_gui_search_sheet',
    upstream: 'search_sheet',
    readOnly: true,
    description:
      'Search inside an Excel workbook for text and return every hit, addressed the way ' +
      'office_gui_read_sheet addresses cells — {sheet, address, value, formula?} — so a hit feeds ' +
      'straight into office_gui_read_sheet { addresses: [address] }. Pass `document` or omit it to ' +
      "search the active session's workbook. `range` scopes the search to an A1 area; sheet/sheetId " +
      'pick one worksheet; regex and lookIn tune the match (matching is case-insensitive); ' +
      'maxResults caps the hits (default 100).',
    shape: {
      document: documentField,
      query: z.string().describe('the text to find'),
      range: z.string().optional().describe('scope: an A1 area, e.g. "A1:D20"'),
      sheet: z
        .string()
        .optional()
        .describe('worksheet name to search; default is every worksheet in the workbook'),
      sheetId: z.string().optional().describe('worksheet id as an alternative to `sheet`'),
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
  },
  {
    name: 'office_gui_search_deck',
    upstream: 'search_deck',
    readOnly: true,
    description:
      'Search inside a PowerPoint presentation for text and return every hit, addressed the way ' +
      'office_gui_read_deck / office_gui_apply_slide_ops address content — {slide, el, type, ' +
      'snippet}, a 0-based slide plus element id — so a hit feeds straight into ' +
      'office_gui_read_deck { slide } or office_gui_apply_slide_ops { slide, el }. Pass `document` ' +
      "or omit it to search the active session's deck. `slide` scopes the search to one slide; " +
      'matchCase tunes the match; maxResults caps the hits (default 100).',
    shape: {
      document: documentField,
      query: z.string().describe('the text to find'),
      slide: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe('0-based slide to search; default is every slide'),
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
  },
  {
    name: 'office_gui_search_pdf',
    upstream: 'search_pdf',
    readOnly: true,
    description:
      'Search an open PDF for text and return every hit, addressed the way office_gui_read_pdf ' +
      'addresses content — {page, snippet} with 1-based pages — so a hit feeds straight ' +
      'into office_gui_read_pdf { document, pages: String(page) }. Pass the PDF as `document` (its ' +
      'tab id or path, or an absolute .pdf path for a file that is not open). `pages` scopes the ' +
      'scan; matchCase tunes the match; maxResults caps the hits (default 100).',
    shape: {
      document: pdfDocumentField,
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
    adapt: async (args, upstream) => ({
      path: await resolveReadPdfPath(upstream, args.document),
      query: args.query,
      pages: args.pages,
      matchCase: args.matchCase,
      maxResults: args.maxResults,
      context: args.context,
    }),
  },
]

/** upstream name -> exposed name, for the proxy handler. */
export const GUI_TOOL_NAMES: ReadonlyMap<string, string> = new Map(
  GUI_TOOLS.map((t) => [t.name, t.upstream]),
)
