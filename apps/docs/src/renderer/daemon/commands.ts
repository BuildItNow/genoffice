import type { Editor } from '@tiptap/react'
import { BLANK_BULLET_NUM_ID, BLANK_ORDERED_NUM_ID } from '@genoffice/docx-engine'
import {
  executeTool,
  markDocSeen,
  type AiCommentsAccess,
  type AiDocExtras,
  type AiHeaderFooterAccess,
} from '../ai/tools'
import type { AiDocWriter } from '../ai/doc-writer'
import type { AiPageSetupAccess } from '../ai/page-setup'
import type { AiNotesAccess } from '../ai/note-ops'
import { findNumId, serializeRangeToHtml, type AiTrack, type NumIds } from '../ai/protocol'
import { isTrackedDeleted } from '../ai/doc-utils'
import {
  compileFind,
  findInDoc,
  matchContext,
  RESULTS_LIST_CAP,
  type FindOptions,
} from '../editor/find'
import { countWords, documentTextForWordCount } from '../word-count'
import type { FileActionContext } from '../file-actions'

/**
 * The daemon's own `office_doc_*` commands, dispatched by the renderer bridge
 * (`../mcp-bridge.ts`) ahead of the built-in MCP vocabulary.
 *
 * The surface is deliberately small: three reads (`office_doc_info` /
 * `office_doc_read` / `office_doc_search`), the block-op batch
 * `office_doc_apply_ops`, and `office_doc_apply_func` — one generic entry that
 * names the editor tool to run in `func` and passes its fields in `args` (the
 * daemon validates them against that tool's own Zod shape and expands them with
 * `office_doc_func_help`). Every content write, comment/note/revision/style read
 * and other edit runs as a func: `insert_content` / `replace_blocks` keep their
 * daemon-side bodies, the rest go through the built-in agent's own
 * `executeTool`, so a daemon edit lands in the edit journal, undo history and
 * save pipeline exactly like an in-app one. The three reads use the find engine
 * and `serializeRangeToHtml` directly.
 */

/**
 * The tool accessors the daemon's `office_doc_*` commands run with. They mirror
 * `createDocsSkill`'s getters (../ai/docs-skill.ts): the App owns the comments /
 * header-footer / page-setup / styles / notes stores, and reaches them through
 * these live getters. Omitted in a context without them (a tool that needs one
 * then fails with its own "not available here" message, exactly as in-app).
 */
export interface DocToolAccessors {
  getTrack?: () => AiTrack | undefined
  getComments?: () => AiCommentsAccess | undefined
  getHf?: () => AiHeaderFooterAccess | undefined
  getWriter?: () => AiDocWriter | undefined
  getPageSetup?: () => AiPageSetupAccess | undefined
  getExtras?: () => AiDocExtras | undefined
  getNotes?: () => AiNotesAccess | undefined
}

/** the daemon commands' deps: the live file-action context (refreshed every render by App) */
export interface DocEditorDeps extends DocToolAccessors {
  getCtx: () => FileActionContext
}

// ---- shared command inputs (the shapes the built-in tools already accept) ----

export interface InsertContentInput {
  html: string
  afterBlockIndex?: number
}
export interface ReplaceBlocksInput {
  startBlockIndex: number
  endBlockIndex: number
  html: string
}
export interface ApplyOpsInput {
  ops: unknown[]
  dryRun?: boolean
}
export interface ReadDocumentInput {
  /** 0-based block index range: "3" or "3-10"; omitted = every block */
  range?: string
  /** whole range in one page instead of paging it */
  full?: boolean
  /** page size in characters when not full */
  maxChars?: number
  /** character offset to continue a paged read */
  offset?: number
}
export interface SearchDocumentInput {
  query: string
  /** 0-based block index range to scope the search: "3" or "3-10"; omitted = whole body */
  range?: string
  matchCase?: boolean
  wholeWord?: boolean
  /** treat `query` as a regular expression instead of literal text */
  regex?: boolean
  maxResults?: number
  /** characters of context to include on either side of a hit */
  context?: number
}
/** one hit, addressed by top-level block index (what `read_document { range }` takes) */
export interface SearchBlockHit {
  block: number
  /** the matched text with `radius` characters of context on either side (ellipses when trimmed) */
  snippet: string
}

/** office_doc_info reply: the document's descriptor (the shell stamps title/path/dirty) */
export interface DocInfoResult {
  blocks: number
  headings: Array<{ index: number; level: number; text: string }>
  words: number
  characters: number
}

/**
 * The editor tools `office_doc_apply_func` may run. The name arrives in the
 * payload, so it must be allow-listed here: `executeTool` also knows the app's
 * chat tools (`web_search`, `image_search`, `create_document`, …), which an
 * external caller must not reach. Keep in sync with `DOC_FUNCS` in
 * packages/daemon/src/office/tools.ts.
 */
const DOC_FUNC_NAMES = new Set([
  'insert_content',
  'replace_blocks',
  'insert_image',
  'insert_picture',
  'insert_text_box',
  'insert_chart',
  'edit_chart',
  'set_header_footer',
  'set_page_setup',
  'insert_section_break',
  'set_watermark',
  'define_style',
  'add_comment',
  'reply_comment',
  'resolve_comment',
  'delete_comment',
  'insert_footnote',
  'insert_endnote',
  'edit_note',
  'delete_note',
  'accept_changes',
  'reject_changes',
  'read_comments',
  'read_revisions',
  'read_notes',
  'list_styles',
])

/** read_document paging: the in-app read_blocks page size, and the CLI's --max-chars cap. */
const READ_PAGE_CHARS = 24_000
const READ_MAX_CHARS = 1_000_000

/** search_document defaults, mirroring the MCP search tool's own bounds */
const SEARCH_RESULTS_DEFAULT = 100
const SEARCH_CONTEXT_DEFAULT = 30
const SEARCH_CONTEXT_MAX = 200

function numIdsFor(ctx: FileActionContext): NumIds {
  const blocks = ctx.doc?.parsed.blocks ?? []
  const isBlank = ctx.doc?.isBlank === true
  return {
    bullet: findNumId(blocks, 'bullet') ?? (isBlank ? BLANK_BULLET_NUM_ID : null),
    ordered: findNumId(blocks, 'ordered') ?? (isBlank ? BLANK_ORDERED_NUM_ID : null),
  }
}

/** mirror of AiPanel's clearAiHighlights: auto-accept the external edit, keeping undo history */
export function clearAiChangedFlags(editor: Editor): void {
  const view = editor.view
  let tr = view.state.tr
  let touched = false
  view.state.doc.forEach((node, offset) => {
    if (node.attrs.aiChanged) {
      tr = tr.setNodeMarkup(offset, undefined, { ...node.attrs, aiChanged: false })
      touched = true
    }
  })
  if (!touched) return
  tr = tr.setMeta('addToHistory', false)
  view.dispatch(tr)
  markDocSeen(editor)
}

/**
 * The executors are shared with the in-app agent, whose tool set includes
 * `get_document_context`. That name does not exist on the MCP surface, where the
 * same readout is `read_document` — so hinting at it would send an external
 * agent chasing a tool it cannot call. Rename the reference on the way out.
 */
export function mcpErrorText(output: string): string {
  return output.replaceAll('get_document_context', 'read_document')
}

/** a resolved editor + its context, or a clear "not ready" failure */
function requireEditor(deps: DocEditorDeps): { editor: Editor; ctx: FileActionContext } {
  const ctx = deps.getCtx()
  const editor = ctx.editor
  if (!editor || !ctx.doc) throw new Error('the document is not ready')
  return { editor, ctx }
}

/** clamp a caller's maxResults into 1..RESULTS_LIST_CAP (the find engine's list cap) */
function clampSearchResults(value: unknown): number {
  const n = Math.trunc(Number(value))
  if (!Number.isFinite(n) || n < 1) return Math.min(SEARCH_RESULTS_DEFAULT, RESULTS_LIST_CAP)
  return Math.min(n, RESULTS_LIST_CAP)
}
/** clamp a caller's context radius into 0..SEARCH_CONTEXT_MAX */
function clampSearchContext(value: unknown): number {
  const n = Math.trunc(Number(value))
  if (!Number.isFinite(n) || n < 0) return SEARCH_CONTEXT_DEFAULT
  return Math.min(n, SEARCH_CONTEXT_MAX)
}

/** parse a CLI-style block range ("3" | "3-10") into 0-based inclusive indexes; default = whole body */
function parseBlockRange(spec: string | undefined, count: number): { start: number; end: number } {
  if (spec === undefined || spec.trim() === '') return { start: 0, end: count - 1 }
  const m = /^(\d+)(?:-(\d+))?$/.exec(spec.trim())
  if (!m)
    throw new Error(`range must look like "3" or "3-10" (0-based block indexes), got "${spec}"`)
  const start = Number(m[1])
  const end = m[2] === undefined ? start : Number(m[2])
  if (start > end || end >= count) {
    throw new Error(
      `range "${spec}" is out of bounds (the document has ${count} blocks, 0-${count - 1})`,
    )
  }
  return { start, end }
}

/**
 * insert_content: append at the end when afterBlockIndex is omitted. An external
 * caller has no cursor, so the in-app default ("after the cursor") has no meaning
 * here — the CLI defaults to the document end the same way.
 */
export async function insertContent(deps: DocEditorDeps, payload: unknown): Promise<unknown> {
  const { editor, ctx } = requireEditor(deps)
  const input = (payload ?? {}) as InsertContentInput
  if (typeof input.html !== 'string') throw new Error('insert_content requires "html"')
  // The shared executor clamps an out-of-range afterBlockIndex to the last
  // block, so a mistyped index quietly inserts somewhere unintended while
  // replace_blocks rejects the same mistake. Reject it here too — the in-app
  // agent keeps the clamping behavior it was built around.
  const last = editor.state.doc.childCount - 1
  if (input.afterBlockIndex !== undefined) {
    if (!Number.isInteger(input.afterBlockIndex) || input.afterBlockIndex < -1) {
      throw new Error(`afterBlockIndex must be an integer >= -1 (got ${input.afterBlockIndex})`)
    }
    if (input.afterBlockIndex > last) {
      throw new Error(
        `afterBlockIndex ${input.afterBlockIndex} is out of range (valid: -1..${last}); ` +
          'call read_document for fresh block indexes',
      )
    }
  }
  const after = input.afterBlockIndex ?? last
  const outcome = await executeTool(
    editor,
    { id: 'mcp', name: 'insert_content', input: { ...input, afterBlockIndex: after } },
    numIdsFor(ctx),
  )
  if (outcome.isError) throw new Error(mcpErrorText(outcome.output))
  clearAiChangedFlags(editor)
  return { summary: outcome.summary, mutated: outcome.mutated }
}

/** replace_blocks: rewrite a block range with new HTML */
export async function replaceBlocks(deps: DocEditorDeps, payload: unknown): Promise<unknown> {
  const { editor, ctx } = requireEditor(deps)
  const input = (payload ?? {}) as ReplaceBlocksInput
  if (typeof input.html !== 'string') throw new Error('replace_blocks requires "html"')
  const outcome = await executeTool(
    editor,
    { id: 'mcp', name: 'replace_blocks', input: { ...input } },
    numIdsFor(ctx),
  )
  if (outcome.isError) throw new Error(mcpErrorText(outcome.output))
  clearAiChangedFlags(editor)
  return { summary: outcome.summary, mutated: outcome.mutated }
}

/** apply_ops: a formatting/structure batch the built-in agent accepts, whole or not at all */
export async function applyOpsCommand(deps: DocEditorDeps, payload: unknown): Promise<unknown> {
  const { editor, ctx } = requireEditor(deps)
  const input = (payload ?? {}) as ApplyOpsInput
  // Route through executeTool rather than executeOps: the wrapper is what
  // refuses a block-indexed batch after the user edited the document (the
  // stale-index guard) and supplies numbering ids, so an external batch
  // behaves exactly like the built-in agent's.
  const outcome = await executeTool(
    editor,
    {
      id: 'mcp',
      name: 'apply_ops',
      input: { ops: input.ops, ...(input.dryRun === true ? { dryRun: true } : {}) },
    },
    numIdsFor(ctx),
  )
  if (outcome.isError) throw new Error(mcpErrorText(outcome.output))
  if (input.dryRun !== true) clearAiChangedFlags(editor)
  return { summary: outcome.summary, output: outcome.output, mutated: outcome.mutated }
}

/**
 * read_document: with no location argument, the live overview (block list
 * preview + stats + selection — context only an open editor has). A range/full/
 * maxChars/offset switches to a lossless, paged read of the range as restricted
 * HTML.
 */
export async function readDocument(deps: DocEditorDeps, payload: unknown): Promise<unknown> {
  const { editor, ctx } = requireEditor(deps)
  const input = (payload ?? {}) as ReadDocumentInput
  const ranged =
    input.range !== undefined ||
    input.full === true ||
    input.maxChars !== undefined ||
    input.offset !== undefined
  if (!ranged) {
    const outcome = await executeTool(
      editor,
      { id: 'mcp', name: 'get_document_context', input: {} },
      numIdsFor(ctx),
    )
    if (outcome.isError) throw new Error(mcpErrorText(outcome.output))
    return { text: outcome.output }
  }
  const count = editor.state.doc.childCount
  const { start, end } = parseBlockRange(input.range, count)
  const html = serializeRangeToHtml(editor, start, end)
  const offset = Math.max(0, Math.trunc(Number(input.offset)) || 0)
  if (offset > 0 && offset >= html.length) {
    throw new Error(`offset ${offset} is beyond the range content (${html.length} characters)`)
  }
  const maxChars = Math.trunc(Number(input.maxChars)) || 0
  if (maxChars > READ_MAX_CHARS) {
    throw new Error(`maxChars must be a positive integer up to ${READ_MAX_CHARS}`)
  }
  // --full lifts the cap; otherwise a page keeps the read small enough to
  // leave context for the rest of the work (maxChars moves it)
  const budget = input.full === true ? Number.POSITIVE_INFINITY : maxChars || READ_PAGE_CHARS
  const slice = html.slice(offset, offset + budget)
  const next = offset + slice.length
  const note =
    next < html.length
      ? `\n…(+${html.length - next} chars; call read_document again with range:"${start}-${end}" offset=${next} to continue)`
      : offset > 0
        ? `\n(end of range: ${html.length} characters in total)`
        : ''
  let text = slice + note
  if (!slice) {
    let deleted = 0
    for (let i = start; i <= end; i++) {
      if (isTrackedDeleted(editor.state.doc.child(i))) deleted++
    }
    text =
      deleted > 0
        ? `(the ${deleted} block(s) in this range are pending tracked deletions — that text is already deleted and hidden from reads; do not rewrite it)`
        : '(range is empty)'
  }
  return {
    text,
    blocks: count,
    range: `${start}-${end}`,
    total: html.length,
    ...(next < html.length ? { offset: next } : {}),
  }
}

/** search_document: every hit addressed by block index (the read_document vocabulary) */
export async function searchDocument(deps: DocEditorDeps, payload: unknown): Promise<unknown> {
  const { editor } = requireEditor(deps)
  const input = (payload ?? {}) as SearchDocumentInput
  if (typeof input.query !== 'string' || input.query === '') {
    throw new Error('search_document requires a non-empty "query"')
  }
  // Same engine the in-app Find panel uses, so hits match what the user sees.
  const opts: FindOptions = {
    matchCase: input.matchCase === true,
    wholeWord: input.wholeWord === true,
    mode: input.regex === true ? 'regex' : 'literal',
  }
  const compiled = compileFind(input.query, opts)
  if (!compiled.ok) throw new Error(`invalid search query: ${compiled.error}`)
  const count = editor.state.doc.childCount
  const scope =
    typeof input.range === 'string' && input.range.trim() !== ''
      ? parseBlockRange(input.range, count)
      : undefined
  const maxResults = clampSearchResults(input.maxResults)
  const context = clampSearchContext(input.context)
  const matches: SearchBlockHit[] = []
  let truncated = false
  for (const m of findInDoc(editor.state.doc, compiled.query, opts)) {
    // index(0) is the top-level block the match sits in — the index
    // read_document / apply_ops address, so a hit feeds straight back in.
    const block = editor.state.doc.resolve(m.from).index(0)
    if (scope && (block < scope.start || block > scope.end)) continue
    if (matches.length >= maxResults) {
      truncated = true
      break
    }
    const hit = matchContext(m, context)
    matches.push({ block, snippet: hit.before + hit.hit + hit.after })
  }
  return { blocks: count, total: matches.length, truncated, matches }
}

/** office_doc_info: the document's block count, headings outline and text stats. */
export function docInfo(deps: DocEditorDeps): DocInfoResult {
  const { editor } = requireEditor(deps)
  let blocks = 0
  const headings: DocInfoResult['headings'] = []
  editor.state.doc.forEach((node) => {
    // a heading is a docHeading node; its level rides on the node (Word's
    // style-numbered headings read as list items, which office_doc_read's own
    // overview also shows as their underlying type)
    if (node.type.name === 'docHeading') {
      const text = node.textContent.trim()
      if (text) {
        headings.push({
          index: blocks,
          level: Math.min(Math.max(Number(node.attrs.level) || 1, 1), 6),
          text,
        })
      }
    }
    blocks++
  })
  const text = documentTextForWordCount(editor.state.doc)
  return { blocks, headings, words: countWords(text), characters: text.length }
}

/**
 * Run one editor tool (`office_doc_apply_func` → `<func>`) on the live editor.
 * The payload IS the tool's input object, so it is handed to the shared
 * `executeTool` unchanged; the tool accessors are read live so a comment / note
 * / header-footer op reaches the App-owned store. A tool that fails is reported
 * as the command's failure.
 */
async function runDocTool(
  deps: DocEditorDeps,
  name: string,
  payload: unknown,
): Promise<Record<string, unknown>> {
  const { editor, ctx } = requireEditor(deps)
  const input = { ...((payload ?? {}) as Record<string, unknown>) }
  // An external caller has no cursor, so insert_chart's cursor-relative default
  // has no meaning here; append at the end as the other insert tools do.
  if (name === 'insert_chart' && input.afterBlockIndex === undefined) {
    input.afterBlockIndex = editor.state.doc.childCount - 1
  }
  const outcome = await executeTool(
    editor,
    { id: 'mcp', name, input },
    numIdsFor(ctx),
    deps.getTrack?.(),
    undefined,
    undefined,
    deps.getComments?.(),
    deps.getHf?.(),
    deps.getWriter?.(),
    deps.getPageSetup?.(),
    deps.getExtras?.(),
    deps.getNotes?.(),
  )
  if (outcome.isError) throw new Error(mcpErrorText(outcome.output))
  if (outcome.mutated) clearAiChangedFlags(editor)
  return { summary: outcome.summary, output: outcome.output, mutated: outcome.mutated }
}

/**
 * Run one daemon command against the live editor. A name this family does not
 * own is not an error here — the bridge's own switch handles it — so it returns
 * `undefined` and the caller can tell the command was unowned. For a command it
 * does own, a thrown error is reported to the caller as that command's failure.
 *
 * `office_doc_apply_ops` has a daemon-only body (the block-op registry batch);
 * `office_doc_apply_func` names its editor tool in `func`. `insert_content` /
 * `replace_blocks` keep their daemon-side bodies (extra validation and the
 * append-at-end default); every other func runs its editor tool directly.
 */
export async function dispatchDaemonCommand(
  command: string,
  payload: Record<string, unknown>,
  deps: DocEditorDeps,
): Promise<unknown | undefined> {
  switch (command) {
    case 'office_doc_info':
      return docInfo(deps)
    case 'office_doc_read':
      return readDocument(deps, payload)
    case 'office_doc_search':
      return searchDocument(deps, payload)
    case 'office_doc_apply_ops':
      return applyOpsCommand(deps, payload)
    case 'office_doc_apply_func': {
      const func = payload.func
      if (typeof func !== 'string' || !DOC_FUNC_NAMES.has(func)) {
        throw new Error(`unknown func "${String(func)}"`)
      }
      switch (func) {
        case 'insert_content':
          return insertContent(deps, payload.args)
        case 'replace_blocks':
          return replaceBlocks(deps, payload.args)
        default:
          return runDocTool(deps, func, payload.args)
      }
    }
  }
  return undefined
}
