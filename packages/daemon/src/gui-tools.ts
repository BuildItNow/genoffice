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
    'tab id or path of an open PDF document to read ' +
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
      target: z.string().optional().describe('path or tab id; required for "read", "save" and "close"'),
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
      'Read the visible document as a block list: one "index|type|text" line per block, where the ' +
      'index is what apply_ops and insert_content address. An overview, not a lossless copy.',
    shape: { document: documentField },
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
      'Read the visible workbook: sheet names with data extents, and optionally the values/formulas ' +
      'of specific cells. Address a worksheet by NAME or by the sheetId the overview reports.',
    shape: {
      addresses: z
        .array(z.string())
        .optional()
        .describe('A1 addresses to read (values + formulas, at most 100); omit for the overview'),
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
      'Ids and slide indexes are what apply_slide_ops addresses; geometry is document-space EMU.',
    shape: { document: documentField },
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
      '`document` (its tab id or path). Returns page count, metadata, and per-page text with ' +
      'page size; scanned pages come back with empty text and hasTextLayer:false. Use `pages` ' +
      'for a bounded read.',
    shape: {
      document: pdfDocumentField,
      pages: z
        .string()
        .optional()
        .describe('1-based pages to read, e.g. "3" or "1-5,8"; default is the whole document'),
    },
    adapt: async (args, upstream) => ({
      path: await resolveReadPdfPath(upstream, args.document),
      pages: args.pages,
    }),
  },
]

/** upstream name -> exposed name, for the proxy handler. */
export const GUI_TOOL_NAMES: ReadonlyMap<string, string> = new Map(
  GUI_TOOLS.map((t) => [t.name, t.upstream]),
)
