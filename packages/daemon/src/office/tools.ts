import { docHelp, type DocOpHelp } from '@genoffice/cli/doc-catalog'
import { sheetHelp } from '@genoffice/cli/sheet-catalog'
import { z, type ZodRawShape } from 'zod'
import { OFFICE_PREFIX } from '../names'
import { createDocFuncs, nearestName, type DocFuncCategory } from './funcs-util'

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
 * it holds. The per-family tools are grouped beside them (the sheet_* set, the
 * doc_* set and the pdf_info / pdf_read / pdf_search reads today), and a new
 * family adds its own `*_info` / `*_read` / `*_edit` set.
 */

export interface OfficeToolSpec {
  /** the bare tool name; the exposed name is `OFFICE_PREFIX + name` */
  name: string
  description: string
  shape: ZodRawShape
  readOnly?: boolean
  /**
   * Validate (and normalize) the call in the daemon before it is forwarded or
   * handed to `local`. Throwing reports the call as an error without reaching
   * the app; the returned object is what is forwarded.
   */
  prepare?: (
    args: Record<string, unknown>,
  ) => Record<string, unknown> | Promise<Record<string, unknown>>
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
  .describe("worksheet name; default is the tab's active sheet")

/**
 * One document-operation spec, named after the editor tool it runs
 * (`AGENT_TOOLS` in apps/docs/src/renderer/ai/tools.ts), so the external surface
 * mirrors the in-app agent's tool set. Each carries the target document's
 * `tabId` plus that tool's own fields; the renderer runs it through the same
 * `executeTool` the built-in agent uses, so a daemon edit lands in the editor
 * exactly like an in-app one. Only apply_ops is registered as its own tool; the
 * rest are exposed as deferred funcs.
 */
function docOp(name: string, description: string, shape: ZodRawShape): OfficeToolSpec {
  return { name: `doc_${name}`, description, shape: { tabId: tabIdField, ...shape } }
}

/** a length: a number of points, or a string like "2.54cm" / "1in" / "72pt" / "96px" */
const docLength = z.union([z.number(), z.string()])

/** the target selector shared by the tracked-change accept/reject tools */
const revisionSelector: ZodRawShape = {
  all: z.boolean().optional().describe('true = every pending change'),
  ids: z
    .array(z.string())
    .optional()
    .describe('revision ids from the read_revisions func (r1, r2, …); re-read before each call'),
  author: z.string().optional().describe('only changes by this author (case-insensitive)'),
  type: z
    .enum(['insertion', 'deletion', 'formatting', 'move'])
    .optional()
    .describe('only changes of this type'),
  blockIndex: z.number().int().optional().describe('only changes inside this block'),
  blockRange: z
    .array(z.number().int())
    .length(2)
    .optional()
    .describe('[from, to] block indexes, inclusive'),
  before: z.string().optional().describe('ISO date; only changes recorded before it'),
}

/** every func-backed document operation, in a readable order (reads last) */
const DOC_OP_TOOLS: OfficeToolSpec[] = [
  docOp(
    'insert_content',
    'Insert new content (restricted HTML, may contain multiple blocks) into the document of one ' +
      'tab, addressed by `tabId`. For writing/continuing/generating new content; to rewrite existing ' +
      'content use the replace_blocks func (office_doc_apply_func). Appended at the end when ' +
      '`afterBlockIndex` is omitted.',
    {
      html: z.string().describe('restricted HTML fragment to insert'),
      afterBlockIndex: z
        .number()
        .int()
        .optional()
        .describe(
          'insert after this block index; -1 = start of document; omitted = end of document',
        ),
    },
  ),
  docOp(
    'replace_blocks',
    'Replace a block range with new content (restricted HTML). For rewriting/translating/condensing ' +
      'existing content; the new block count may differ from the old. New blocks inherit the replaced ' +
      "blocks' formatting automatically. Block indexes come from office_doc_read.",
    {
      startBlockIndex: z.number().int().describe('start block index (0-based, inclusive)'),
      endBlockIndex: z.number().int().describe('end block index (inclusive)'),
      html: z.string().describe('replacement restricted HTML fragment'),
    },
  ),
  docOp(
    'insert_image',
    'Download a direct image link and insert it into the document of one tab (appended at the end). ' +
      'For a sized / floating picture at a chosen block, use the insert_picture func ' +
      '(office_doc_apply_func).',
    {
      url: z.string().describe('direct image link (png/jpg/gif)'),
      maxWidthPx: z.number().int().optional().describe('maximum width in px (default 480)'),
    },
  ),
  docOp(
    'insert_picture',
    'Insert a picture from a direct image link at a block position, at a given size, inline or ' +
      'floating (positioned on the page, with text wrapping around/behind/in front).',
    {
      url: z.string().describe('direct image link (png/jpg/gif)'),
      afterBlockIndex: z
        .number()
        .int()
        .optional()
        .describe(
          'insert after this block index; -1 = start of document; omitted = end of document',
        ),
      width: docLength.optional().describe('display width; default natural size capped at 480px'),
      height: docLength.optional().describe('display height; omit to keep the aspect ratio'),
      float: z
        .any()
        .optional()
        .describe(
          'floating placement: { anchor: "paragraph"|"page"|"margin", wrap: "square"|"tight"|"topAndBottom"|"behind"|"inFront", x?, y? }',
        ),
      altText: z.string().optional().describe('alternative text (accessibility)'),
    },
  ),
  docOp(
    'insert_text_box',
    'Insert a floating text box (a positioned rectangle holding text, like a callout or sidebar) ' +
      'anchored to a block.',
    {
      afterBlockIndex: z
        .number()
        .int()
        .optional()
        .describe(
          'anchor after this block index; -1 = start of document; omitted = end of document',
        ),
      text: z.string().describe('box text; \\n separates paragraphs'),
      width: docLength.optional().describe('box width'),
      height: docLength.optional().describe('box height'),
      x: docLength.optional().describe('horizontal offset'),
      y: docLength.optional().describe('vertical offset'),
      anchor: z
        .enum(['paragraph', 'page'])
        .optional()
        .describe('offsets measure from the anchor paragraph (default) or the page corner'),
      wrap: z
        .enum(['topAndBottom', 'none'])
        .optional()
        .describe('body text resumes below the box, or ignores it'),
      fill: z
        .string()
        .nullable()
        .optional()
        .describe('"#RRGGBB" background (default white); null = transparent'),
    },
  ),
  docOp(
    'insert_chart',
    'Insert a chart (saved as a native Word chart) into the document of one tab. Data must be real.',
    {
      kind: z.enum(['bar', 'line', 'pie']).describe('chart type'),
      title: z.string().optional().describe('chart title'),
      categories: z.array(z.string()).describe('category (x axis / sector) labels'),
      series: z
        .array(z.any())
        .describe(
          'data series: [{ name?, values: (number|null)[] }]; values match categories in length',
        ),
      afterBlockIndex: z
        .number()
        .int()
        .optional()
        .describe(
          'insert after this block index; -1 = start of document; omitted = end of document',
        ),
    },
  ),
  docOp(
    'edit_chart',
    'Edit the data of an existing chart block (title / category labels / series names / values). The ' +
      'category count and the number of values per series must match the original chart.',
    {
      blockIndex: z.number().int().describe('block index of the chart'),
      title: z.string().optional().describe('new title (omit to keep)'),
      categories: z
        .array(z.string().nullable())
        .optional()
        .describe('new category labels, same length as the original; null keeps a position'),
      series: z
        .array(z.any())
        .optional()
        .describe('series to change: [{ index, name?, values? }]; null keeps a position'),
    },
  ),
  docOp(
    'set_header_footer',
    'Set the page header or footer text (plain text; \\n separates lines; {PAGE}/{NUMPAGES} become ' +
      'page-number fields; "" clears). view "first"/"even" writes the different-first-page / even-page ' +
      'variant (enabling that setting if needed).',
    {
      kind: z.enum(['header', 'footer']).describe('which band to write'),
      text: z
        .string()
        .describe('new text; \\n between lines; may contain {PAGE} / {NUMPAGES}; "" clears'),
      view: z
        .enum(['default', 'first', 'even'])
        .optional()
        .describe('which variant to write (default when omitted)'),
    },
  ),
  docOp(
    'set_page_setup',
    'Change paper size, orientation, margins, text columns, the different-first-page flag or page ' +
      'numbering of the document or of one section. Only the fields given change.',
    {
      section: z
        .number()
        .int()
        .optional()
        .describe('section index (0-based) to change; omit with blockIndex = every section'),
      blockIndex: z.number().int().optional().describe('change the section containing this block'),
      paper: z.string().optional().describe('named size (A3, A4, A5, B5, Letter, Legal, Tabloid)'),
      width: docLength.optional().describe('custom page width'),
      height: docLength.optional().describe('custom page height'),
      orientation: z.enum(['portrait', 'landscape']).optional(),
      margins: z
        .any()
        .optional()
        .describe('any of top/right/bottom/left/header/footer/gutter, each a length'),
      columns: z.any().optional().describe('{ count, spacing? }'),
      titlePg: z.boolean().optional().describe('different first page'),
      pageNumberStart: z.number().int().optional(),
      pageNumberFormat: z.string().optional(),
    },
  ),
  docOp(
    'insert_section_break',
    'Insert a section break after a block so the blocks after it form a new section with its own page ' +
      'setup, headers and footers (the new section starts as a copy of the current one).',
    {
      afterBlockIndex: z
        .number()
        .int()
        .describe('the break goes after this block index; -1 = before the first block'),
      type: z
        .enum(['nextPage', 'continuous', 'evenPage', 'oddPage'])
        .optional()
        .describe('default nextPage'),
    },
  ),
  docOp(
    'set_watermark',
    'Set, restyle or remove the page watermark: gray text or a washed-out picture, drawn behind the ' +
      'body on every page. Give text or image, not both; text null or image null removes it.',
    {
      text: z.string().nullable().optional().describe('watermark text; null removes the watermark'),
      fontFamily: z.string().optional(),
      color: z.string().optional().describe('"#RRGGBB" (default silver)'),
      opacity: z.number().optional().describe('0-1 (default 0.5)'),
      diagonal: z.boolean().optional().describe('diagonal (default) or horizontal layout'),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      image: z
        .string()
        .nullable()
        .optional()
        .describe('picture watermark URL (png/jpg/gif); null removes it'),
      scale: z.number().optional().describe('picture size as a percent of its natural size'),
      washout: z
        .boolean()
        .optional()
        .describe("fade the picture like Word's washout (default true)"),
    },
  ),
  docOp(
    'define_style',
    'Create a style or change the definition of an existing one (paragraph, character or table style ' +
      'in styles.xml). Only the fields given change; paragraphs using the style pick the change up ' +
      'automatically. Put a style on paragraphs with the applyStyle op (office_doc_apply_ops).',
    {
      styleId: z
        .string()
        .describe('style id (letters, digits, - and _; e.g. "Heading1", "MyCallout")'),
      name: z.string().optional().describe('display name (defaults to the id when creating)'),
      type: z
        .enum(['paragraph', 'character', 'table'])
        .optional()
        .describe('when creating (default paragraph)'),
      basedOn: z
        .string()
        .nullable()
        .optional()
        .describe('parent style id; null removes the inheritance'),
      next: z
        .string()
        .nullable()
        .optional()
        .describe('style of the paragraph created by Enter at the end'),
      quickFormat: z.boolean().optional(),
      paragraph: z.any().optional().describe('paragraph formatting (align, spacing, indent, …)'),
      run: z.any().optional().describe('character formatting (bold, italic, color, size, …)'),
      table: z.any().optional().describe('table formatting'),
    },
  ),
  docOp(
    'add_comment',
    'Start a new comment thread on a block or on an exact text span inside it, without changing the ' +
      'document: review notes, questions, suggestions the user should decide on.',
    {
      blockIndex: z.number().int().describe('block to annotate'),
      text: z
        .string()
        .optional()
        .describe('exact text inside the block to anchor to; omitted = whole block'),
      occurrence: z
        .number()
        .int()
        .optional()
        .describe('which match to anchor when text occurs more than once (1 = first)'),
      comment: z.string().describe('comment body; \\n starts a new paragraph'),
      author: z.string().optional().describe('author name; default: the AI author'),
      initials: z.string().optional().describe('author initials shown in the margin'),
    },
  ),
  docOp(
    'reply_comment',
    'Add a reply to a comment thread (ids from the read_comments func): after completing a ' +
      'requested change, summarize what changed; or answer/ask back when the comment is a question.',
    {
      parentId: z.string().describe('id of the comment thread to reply to'),
      text: z.string().describe('reply text'),
    },
  ),
  docOp(
    'resolve_comment',
    'Mark a comment thread as resolved (ids from the read_comments func). Only after the requested ' +
      'change was applied, or when the user explicitly asked to resolve.',
    { id: z.string().describe('id of the comment thread to resolve') },
  ),
  docOp(
    'delete_comment',
    'Delete a comment. A thread root with replies is refused unless withReplies is true; a reply id ' +
      'deletes just that reply.',
    {
      id: z.string().describe('id of the comment to delete'),
      withReplies: z
        .boolean()
        .optional()
        .describe('also delete the replies of a thread root (default false)'),
    },
  ),
  docOp(
    'insert_footnote',
    'Add a footnote: a superscript reference mark goes into the block (right after afterText, else at ' +
      'its end) and the note text prints at the bottom of that page.',
    {
      blockIndex: z.number().int().describe('block that gets the reference mark'),
      afterText: z
        .string()
        .optional()
        .describe('exact text the mark follows; omitted = end of the block'),
      text: z.string().describe('the note text (\\n separates paragraphs)'),
    },
  ),
  docOp(
    'insert_endnote',
    'Add an endnote (like the insert_footnote func, but the text collects at the end of the document).',
    {
      blockIndex: z.number().int().describe('block that gets the reference mark'),
      afterText: z
        .string()
        .optional()
        .describe('exact text the mark follows; omitted = end of the block'),
      text: z.string().describe('the note text (\\n separates paragraphs)'),
    },
  ),
  docOp(
    'edit_note',
    'Change the text of an existing footnote or endnote in place: the note keeps its id, reference mark ' +
      'and formatting. Use for a citation fix instead of delete + insert.',
    {
      kind: z.enum(['footnote', 'endnote']).optional().describe('omit when the id is unambiguous'),
      id: z.string().describe('note id from the read_notes func (office_doc_apply_func)'),
      find: z.string().describe('exact text inside the note to replace'),
      replace: z.string().describe('replacement text ("" deletes the match)'),
      matchCase: z.boolean().optional().describe('defaults to true'),
    },
  ),
  docOp(
    'delete_note',
    'Remove a footnote or endnote together with its reference mark (ids from the read_notes func).',
    {
      kind: z.enum(['footnote', 'endnote']),
      id: z.string().describe('note id from the read_notes func (office_doc_apply_func)'),
    },
  ),
  docOp(
    'accept_changes',
    'Accept pending tracked changes: insertions become plain text, deleted text disappears, new ' +
      'formatting stays. Select with all: true, ids from the read_revisions func, or a combination ' +
      'of author / type / blockIndex / blockRange / before.',
    revisionSelector,
  ),
  docOp(
    'reject_changes',
    'Reject pending tracked changes: inserted text disappears, deleted text is restored, formatting ' +
      'goes back. Same selector as the accept_changes func.',
    revisionSelector,
  ),
  docOp(
    'read_comments',
    'List all comment threads (including resolved ones) with ids, authors, anchored block indexes and ' +
      'anchor text.',
    {},
  ),
  docOp(
    'read_revisions',
    'List every pending tracked revision with its id, type, author, date, block index and the affected ' +
      'text. Ids are positional (r1 = first in document order) and renumber after every edit.',
    {},
  ),
  docOp(
    'read_notes',
    'List every footnote and endnote with its id, number, the block holding its reference mark and its ' +
      'text.',
    {},
  ),
  docOp(
    'list_styles',
    'List the styles defined in the document (id, name, type, parent, heading level) for the ' +
      'define_style func and the applyStyle / setFont / setMatchedFont ops.',
    {},
  ),
]

/** fields an external caller writes; the per-call tabId is not part of a func */
function funcFields(shape: ZodRawShape): ZodRawShape {
  return Object.fromEntries(Object.entries(shape).filter(([key]) => key !== 'tabId'))
}

/** illustrative args for the funcs whose shape is easiest to get wrong */
const DOC_FUNC_EXAMPLES: Record<string, Record<string, unknown>> = {
  insert_content: { html: '<p>New paragraph</p>', afterBlockIndex: 3 },
  replace_blocks: { startBlockIndex: 2, endBlockIndex: 3, html: '<p>Rewritten paragraph</p>' },
  add_comment: {
    blockIndex: 3,
    text: 'the deadline sentence',
    comment: 'Is this date final?',
  },
  insert_picture: {
    url: 'https://example.com/chart.png',
    width: '8cm',
    float: { anchor: 'paragraph', wrap: 'square' },
  },
  reply_comment: { parentId: 'c1', text: 'Done — updated in paragraph 4.' },
  set_page_setup: { paper: 'A4', orientation: 'landscape' },
}

/**
 * The func help categories, in the order the index lists them. A func's category
 * is named below, so `office_doc_func_help` can expand one group at a time
 * instead of dumping all 26.
 */
const DOC_FUNC_CATEGORIES: DocFuncCategory[] = [
  { name: 'content', summary: 'new content, block rewrites, pictures, text boxes and charts' },
  { name: 'layout', summary: 'header/footer, page setup, section breaks and watermark' },
  { name: 'styles', summary: 'style definitions and the style list' },
  { name: 'comments', summary: 'comment threads: read, add, reply, resolve, delete' },
  { name: 'notes', summary: 'footnotes and endnotes: read, insert, edit, delete' },
  { name: 'revisions', summary: 'tracked changes: read, accept, reject' },
]

/** func name -> category; every func must be named here exactly once */
const DOC_FUNC_CATEGORY: Record<string, string> = {
  insert_content: 'content',
  replace_blocks: 'content',
  insert_image: 'content',
  insert_picture: 'content',
  insert_text_box: 'content',
  insert_chart: 'content',
  edit_chart: 'content',
  set_header_footer: 'layout',
  set_page_setup: 'layout',
  insert_section_break: 'layout',
  set_watermark: 'layout',
  define_style: 'styles',
  list_styles: 'styles',
  read_comments: 'comments',
  add_comment: 'comments',
  reply_comment: 'comments',
  resolve_comment: 'comments',
  delete_comment: 'comments',
  read_notes: 'notes',
  insert_footnote: 'notes',
  insert_endnote: 'notes',
  edit_note: 'notes',
  delete_note: 'notes',
  read_revisions: 'revisions',
  accept_changes: 'revisions',
  reject_changes: 'revisions',
}

/**
 * Every func-backed operation, in the table's order: same spec, minus the
 * per-call tabId. `funcs-util` owns validation and help.
 */
const DOC_FUNCS = createDocFuncs(
  DOC_OP_TOOLS.map((tool) => {
    const name = tool.name.replace(/^doc_/, '')
    const category = DOC_FUNC_CATEGORY[name]
    if (!category) throw new Error(`doc func "${name}" has no help category`)
    const example = DOC_FUNC_EXAMPLES[name]
    return {
      name,
      description: tool.description,
      shape: funcFields(tool.shape),
      category,
      ...(example ? { example } : {}),
    }
  }),
  DOC_FUNC_CATEGORIES,
)

const FUNC_NAMES = DOC_FUNCS.names

/**
 * apply_ops' block ops, grouped for tiered help. The catalog's own `block` group
 * is one flat list; these give the AI a smaller first step. Every block op the
 * catalog carries must appear here exactly once (the tests enforce it), so the
 * grouping cannot drift from what apply_ops executes.
 */
const DOC_OP_CATEGORIES: { name: string; summary: string; ops: string[] }[] = [
  {
    name: 'text',
    summary: 'run-level character formatting and find/replace',
    ops: ['setFont', 'setMatchedFont', 'findReplace'],
  },
  {
    name: 'paragraph',
    summary: 'paragraph formatting, indent, outline, heading and paragraph style',
    ops: [
      'setParagraphFormat',
      'setParagraphAttrs',
      'setHeadingLevel',
      'setOutlineLevel',
      'stepIndent',
      'stepHangingIndent',
      'applyStyle',
    ],
  },
  { name: 'list', summary: 'list numbering', ops: ['setList', 'clearList'] },
  {
    name: 'structure',
    summary: 'delete or move blocks, insert a table of contents',
    ops: ['deleteBlocks', 'moveBlocks', 'insertToc'],
  },
  {
    name: 'table',
    summary: 'table rows and columns, merges, cell formatting and table style',
    ops: [
      'insertTableRow',
      'deleteTableRow',
      'insertTableColumn',
      'deleteTableColumn',
      'mergeTableCells',
      'splitTableCell',
      'setTableCellFormat',
      'setTableStyle',
    ],
  },
  {
    name: 'field',
    summary: 'inline fields, bookmarks and field refresh',
    ops: ['insertField', 'insertBookmark', 'updateFields'],
  },
  { name: 'picture', summary: 'image block properties', ops: ['setImageProperties'] },
]

interface DocOpsHelp {
  categories?: { name: string; summary: string; ops?: string[] }[]
  ops?: DocOpHelp[]
  unknown?: {
    categories: { name: string; suggestion?: string }[]
    ops: { op: string; suggestion?: string }[]
  }
}

/**
 * office_doc_ops_help: the catalog's `block` entries only, in the categories
 * above. The catalog also carries `content` entries (insert_content, add_comment,
 * …) that `genoffice docs apply` runs headless but the daemon exposes as funcs;
 * listing them here would advertise ops `office_doc_apply_ops` cannot execute.
 */
async function docOpsHelp(req: {
  list?: readonly string[]
  ops?: readonly string[]
}): Promise<DocOpsHelp> {
  // no names at all: the category index, so the first call stays tiny
  if (!req.list && !req.ops) {
    return {
      categories: DOC_OP_CATEGORIES.map(({ name, summary }) => ({ name, summary })),
    }
  }
  // the catalog remains the source of truth for op names and their fields
  const raw = await docHelp({ list: ['block'], ...(req.ops ? { ops: req.ops } : {}) })
  const block = raw.categories?.find((c) => c.name === 'block')
  const blockOps = new Set(block?.ops ?? [])
  const out: DocOpsHelp = {}
  const unknownCategories: { name: string; suggestion?: string }[] = []
  if (req.list) {
    out.categories = []
    for (const name of req.list) {
      const category = DOC_OP_CATEGORIES.find((c) => c.name === name)
      if (!category) {
        const suggestion = nearestName(
          name,
          DOC_OP_CATEGORIES.map((c) => c.name),
        )
        unknownCategories.push(suggestion ? { name, suggestion } : { name })
        continue
      }
      out.categories.push({
        name: category.name,
        summary: category.summary,
        ops: category.ops.filter((op) => blockOps.has(op)),
      })
    }
  }
  const unknownOps: { op: string; suggestion?: string }[] = []
  if (req.ops) {
    out.ops = []
    for (const entry of raw.ops ?? []) {
      if (entry.group === 'block') {
        out.ops.push(entry)
      } else {
        const suggestion = nearestName(entry.op, blockOps)
        unknownOps.push(suggestion ? { op: entry.op, suggestion } : { op: entry.op })
      }
    }
    for (const entry of raw.unknown?.ops ?? []) unknownOps.push(entry)
  }
  if (unknownCategories.length || unknownOps.length) {
    out.unknown = { categories: unknownCategories, ops: unknownOps }
  }
  return out
}

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
      "plus `info` — a workbook's sheets (the office_sheet_info shape), a Word document's block " +
      "outline and text stats (the office_doc_info shape), or a PDF's descriptor — so the first " +
      'call is enough and that info needs no separate one. Use tabId with the other office_* ' +
      'tools to work on it.',
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
        .describe(
          'page the range: stop after this many rows and report the range to continue from',
        ),
    },
    readOnly: true,
  },
  {
    name: 'sheet_read_formats',
    description:
      "Read one worksheet of a tab exactly like sheet_read, but each grid cell carries the cell's " +
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
    name: 'sheet_apply_ops',
    description:
      'Apply workbook DSL operations to one worksheet of a tab. `ops` is an array of op objects ' +
      "addressed by A1. Do NOT give each op its own `sheetId`: this call's `sheet` (default: the " +
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
      "For the other ops and each one's exact fields, call office_sheet_ops_help. " +
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
    name: 'sheet_ops_help',
    description:
      "The reference for office_sheet_apply_ops's workbook ops. `list` names categories and " +
      "returns each one's op names; `ops` names ops and returns each one's field signature and " +
      "note. Categories:\n" +
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
      'Pure static reference — needs no open document, so call it before office_sheet_apply_ops.',
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
    name: 'doc_info',
    description:
      'Describe the Word document in one tab: its block count, a headings outline and word / ' +
      "character counts, plus the tab's title, absolute path (null when never saved) and `dirty` " +
      '(edits not yet written to disk). Blocks are addressed by their 0-based index — the number ' +
      'office_doc_read and office_doc_apply_ops take — so the outline is a table of contents you can ' +
      'act on. A cheap first call before office_doc_read (the live overview or a range) or ' +
      'office_doc_search.',
    shape: { tabId: tabIdField },
    readOnly: true,
  },
  {
    name: 'doc_read',
    description:
      'Read the Word document in one tab, addressed by `tabId` (see office_list). By default it is ' +
      'an overview: a block list (one "index|type|text" line per block, where the index is what ' +
      'office_doc_apply_ops addresses), full-text stats and the current selection; long blocks are ' +
      'clipped. Pass `range` (0-based block indexes, e.g. "3" or "3-10") and/or `full` to read that ' +
      'range as restricted HTML — ready to rewrite and hand back to the replace_blocks func ' +
      '(office_doc_apply_func). A range longer ' +
      'than one page is paged: the result names the `offset` to continue from, so concatenating the ' +
      'pages yields the whole range; `maxChars` moves the page size. Mirrors `genoffice docs read`.',
    shape: {
      tabId: tabIdField,
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
        .positive()
        .optional()
        .describe('page size in characters (default 24000); ignored when full is true'),
      offset: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe('character offset to continue a paged read (default 0)'),
    },
    readOnly: true,
  },
  {
    name: 'doc_search',
    description:
      'Search inside the Word document of one tab and return every hit, addressed the way ' +
      'office_doc_read addresses content — { block, snippet } — so a hit feeds straight into ' +
      'office_doc_read { tabId, range: String(block) }. Uses the same engine as the in-app Find ' +
      'panel. `range` scopes the search to a 0-based block index range ("3" or "3-10"); matchCase, ' +
      'wholeWord and regex tune the match; maxResults caps the hits (default 100, max 500); context ' +
      'is the characters of context around each hit (default 30).',
    shape: {
      tabId: tabIdField,
      query: z.string().describe('the text to find'),
      range: z
        .string()
        .optional()
        .describe(
          'scope the search to a 0-based block index range, e.g. "3" or "3-10"; default every block',
        ),
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
    readOnly: true,
  },
  // the one doc op that keeps its own tool: its payload is an open op list, so
  // a per-func Zod shape cannot describe it
  {
    name: 'doc_apply_ops',
    description:
      'Apply block ops to the existing blocks of the Word document of one tab. `ops` is an array of ' +
      'flat { op, target?, ...fields } objects, executed in order; fields are patches (present = ' +
      'set, null = clear, absent = untouched). `target` selects the blocks an op applies to — at ' +
      'least one of `blockIndexes` (0-based, from office_doc_read), `nodeType`, `headingLevel`, ' +
      '`containsText` or `scope`: "selection". Common ops:\n' +
      '- setFont {target, bold?, italic?, underline?, color?, fontSize?, fontFamily?, …} — run-level ' +
      'character formatting\n' +
      '- setParagraphFormat {target, align?, lineSpacing?, indentLeft?, spaceBefore?, spaceAfter?, …} ' +
      '— paragraph formatting (lengths in twips)\n' +
      '- setHeadingLevel {target, level: 0-6} — heading level; 0 demotes to a body paragraph\n' +
      '- applyStyle {target, styleId} — put a style from styles.xml / define_style on the blocks\n' +
      '- setList {target, kind?: "bullet"|"number"} — paragraphs to list items (default bullet)\n' +
      '- findReplace {find, replace, matchCase?, target?} — plain-text find & replace inside runs\n' +
      '- deleteBlocks {target} — delete the targeted blocks\n' +
      '- moveBlocks {blockIndexes, afterBlockIndex} — move blocks; -1 = start of document\n' +
      '- insertTableRow {target, row, position?, count?} — insert rows into the matched table\n' +
      '- mergeTableCells {target, range: {row, col, rowEnd?, colEnd?}} — merge a grid rectangle\n' +
      '- setTableStyle {target, styleId?, look?} — table style and Word look options\n' +
      "For the other ops and each one's exact fields, call office_doc_ops_help. One batch applies " +
      'as one atomic transaction (a single undo step); a failed batch changes nothing — pass dryRun ' +
      'to validate and report the plan without modifying the document.',
    shape: {
      tabId: tabIdField,
      ops: z
        .array(z.any())
        .describe(
          'flat block-op objects ({ op, target?, ...fields }), executed in order, e.g. ' +
            '[{"op":"setFont","target":{"nodeType":"docHeading"},"color":"#FF0000"}]',
        ),
      dryRun: z
        .boolean()
        .optional()
        .describe('validate the batch and report the plan, without modifying the document'),
    },
  },
  {
    name: 'doc_apply_func',
    description:
      'Run one document function on the document of one tab, addressed by `tabId`. `func` picks ' +
      'the function; `args` carries its fields — do NOT repeat `tabId` inside `args`, the tab is ' +
      'addressed once, on the call. Common funcs:\n' +
      '- insert_content {html, afterBlockIndex?} — append new content (restricted HTML); -1 = start\n' +
      '- replace_blocks {startBlockIndex, endBlockIndex, html} — rewrite a block range in place\n' +
      '- insert_picture {url, afterBlockIndex?, width?, height?, float?} — sized inline or floating ' +
      'picture\n' +
      '- add_comment {blockIndex, comment, text?} — anchor a review comment to text or a whole block\n' +
      '- reply_comment {parentId, text} — answer a comment thread\n' +
      '- set_header_footer {kind, text, view?} — header/footer text; {PAGE}/{NUMPAGES} are fields\n' +
      "For the other funcs and each one's exact fields, call office_doc_func_help and expand " +
      'one category or a few funcs at a time. `args` is validated strictly — an unknown field is ' +
      'an error, not ignored. For formatting/structure edits of existing blocks use ' +
      'office_doc_apply_ops instead.',
    shape: {
      tabId: tabIdField,
      func: z.enum(FUNC_NAMES).describe('the function to run'),
      args: z
        .record(z.string(), z.unknown())
        .optional()
        .describe("that function's fields; office_doc_func_help lists them"),
    },
    prepare: (args) => DOC_FUNCS.prepare(args),
  },
  {
    name: 'doc_func_help',
    description:
      'The reference for office_doc_apply_func, expanded one tier at a time. With no arguments it ' +
      "returns the category index; `list` names categories and returns each one's func names; " +
      "`funcs` names funcs and returns each one's full field schema and example. Categories:\n" +
      DOC_FUNC_CATEGORIES.map((c) => `- ${c.name} — ${c.summary}`).join('\n') +
      '\nAn unknown name comes back under `unknown` with a nearest-match suggestion, so a typo ' +
      'does not fail the call. Pure static reference — needs no open document, so call it before ' +
      'office_doc_apply_func.',
    shape: {
      list: z
        .array(z.string())
        .optional()
        .describe('category names to list the funcs of, e.g. ["comments", "notes"]'),
      funcs: z
        .array(z.string())
        .optional()
        .describe('func names to document, e.g. ["insert_picture", "add_comment"]'),
    },
    readOnly: true,
    local: (args) =>
      DOC_FUNCS.help({
        ...(Array.isArray(args.list) ? { list: args.list as string[] } : {}),
        ...(Array.isArray(args.funcs) ? { funcs: args.funcs as string[] } : {}),
      }),
  },
  {
    name: 'doc_ops_help',
    description:
      "The reference for office_doc_apply_ops's block ops, expanded one tier at a time. With no " +
      "arguments it returns the category index; `list` names categories and returns each one's op " +
      "names; `ops` names ops and returns each one's field signature and note (and whether a " +
      '`target` is required and the keys it may carry). Categories:\n' +
      DOC_OP_CATEGORIES.map((c) => `- ${c.name} — ${c.summary}`).join('\n') +
      '\nThe content functions (pictures, charts, header/footer, comments, notes, tracked changes, ' +
      '…) are office_doc_apply_func calls — use office_doc_func_help for those. An unknown name ' +
      'comes back under `unknown` with a nearest-match suggestion, so a typo does not fail the ' +
      'call. Pure static reference — needs no open document, so call it before ' +
      'office_doc_apply_ops.',
    shape: {
      list: z
        .array(z.string())
        .optional()
        .describe('category names to list the ops of, e.g. ["text", "paragraph"]'),
      ops: z
        .array(z.string())
        .optional()
        .describe('op names to document, e.g. ["setFont", "setHeadingLevel"]'),
    },
    readOnly: true,
    local: (args) =>
      docOpsHelp({
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
      tabId: z
        .string()
        .optional()
        .describe('an open PDF tab to describe, as an id from office_list'),
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
