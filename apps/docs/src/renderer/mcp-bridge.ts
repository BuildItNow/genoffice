import type { Editor } from '@tiptap/react'
import { BLANK_BULLET_NUM_ID, BLANK_ORDERED_NUM_ID } from '@genoffice/docx-engine'
import type { McpCommandMessage, McpEditorCommand } from '../shared/ipc'
import { executeTool, markDocSeen } from './ai/tools'
import { findNumId, serializeRangeToHtml, type NumIds } from './ai/protocol'
import { isTrackedDeleted } from './ai/doc-utils'
import { save, type FileActionContext } from './file-actions'

/**
 * MCP bridge (renderer half).
 *
 * The shell main process pushes editor commands over `docs:mcp-command`; this
 * module runs them against the *live* Tiptap editor so an external agent drives
 * the same visible editor the built-in agent does, then reports the outcome
 * back on `docs:mcp-result`. Command execution is serialized so a burst cannot
 * interleave two edits into one document.
 *
 * The executors are the built-in agent's own (`executeTool`), so external edits
 * inherit the same parsing, atomicity, formatting rules and stale-index guard.
 */

export interface McpBridgeDeps {
  /** live file-action context (refreshed every render by App) */
  getCtx: () => FileActionContext
}

function numIdsFor(ctx: FileActionContext): NumIds {
  const blocks = ctx.doc?.parsed.blocks ?? []
  const isBlank = ctx.doc?.isBlank === true
  return {
    bullet: findNumId(blocks, 'bullet') ?? (isBlank ? BLANK_BULLET_NUM_ID : null),
    ordered: findNumId(blocks, 'ordered') ?? (isBlank ? BLANK_ORDERED_NUM_ID : null),
  }
}

/** the payload shapes this bridge accepts per command (validated by the executors) */
interface InsertContentInput {
  html: string
  afterBlockIndex?: number
}
interface ReplaceBlocksInput {
  startBlockIndex: number
  endBlockIndex: number
  html: string
}
interface ApplyOpsInput {
  ops: unknown[]
  dryRun?: boolean
}
interface ReadDocumentInput {
  /** 0-based block index range: "3" or "3-10"; omitted = every block */
  range?: string
  /** whole range in one page instead of paging it */
  full?: boolean
  /** page size in characters when not full */
  maxChars?: number
  /** character offset to continue a paged read */
  offset?: number
}

/** read_document paging: the in-app read_blocks page size, and the CLI's --max-chars cap. */
const READ_PAGE_CHARS = 24_000
const READ_MAX_CHARS = 1_000_000

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

/** mirror of AiPanel's clearAiHighlights: auto-accept the external edit, keeping undo history */
function clearAiChangedFlags(editor: Editor): void {
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
function mcpErrorText(output: string): string {
  return output.replaceAll('get_document_context', 'read_document')
}

async function runCommand(
  deps: McpBridgeDeps,
  command: McpEditorCommand,
  payload: unknown,
): Promise<unknown> {
  const ctx = deps.getCtx()
  const editor = ctx.editor
  if (!editor || !ctx.doc) throw new Error('the document is not ready')

  switch (command) {
    case 'insert_content': {
      const input = (payload ?? {}) as InsertContentInput
      if (typeof input.html !== 'string') throw new Error('insert_content requires "html"')
      // The shared executor clamps an out-of-range afterBlockIndex to the last
      // block, so a mistyped index quietly inserts somewhere unintended while
      // replace_blocks rejects the same mistake. Reject it here too — the
      // in-app agent keeps the clamping behavior it was built around.
      if (input.afterBlockIndex !== undefined) {
        const last = editor.state.doc.childCount - 1
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
      const outcome = await executeTool(
        editor,
        { id: 'mcp', name: 'insert_content', input: { ...input } },
        numIdsFor(ctx),
      )
      if (outcome.isError) throw new Error(mcpErrorText(outcome.output))
      clearAiChangedFlags(editor)
      return { summary: outcome.summary, mutated: outcome.mutated }
    }

    case 'replace_blocks': {
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

    case 'apply_ops': {
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

    case 'read_document': {
      const input = (payload ?? {}) as ReadDocumentInput
      // No location argument keeps the live overview (block list preview + stats +
      // selection — context only an open editor has). A range/full/maxChars/offset
      // switches to a lossless, paged read of the range as restricted HTML.
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

    case 'save_document': {
      const input = (payload ?? {}) as { path?: string; overwrite?: boolean }
      if (typeof input.path !== 'string' || !input.path) {
        throw new Error('save_document requires an absolute "path"')
      }
      let reason = ''
      const ok = await save(ctx, false, true, undefined, {
        path: input.path,
        overwrite: input.overwrite === true,
        onError: (message) => (reason = message),
      })
      if (!ok) throw new Error(reason || 'the document could not be saved')
      return { ok: true, path: input.path }
    }

    default: {
      const unreachable: never = command
      throw new Error(`unknown MCP command: ${String(unreachable)}`)
    }
  }
}

/**
 * A fresh tab boots asynchronously (`newFile()` calls setContent, then setDoc),
 * so a command that lands before `doc` exists would be wiped by that blank
 * reset. Wait for the loaded document before announcing readiness.
 *
 * The wait never stops retrying, only slows down: a tab that gave up while its
 * renderer was still booting would look fine in the UI yet stay unaddressable
 * over MCP for the rest of its life, because readiness is announced exactly
 * once and never re-derived.
 */
const READY_POLL_MS = 50
const READY_SLOW_POLL_MS = 1_000
const READY_FAST_WINDOW_MS = 20_000

async function announceWhenLoaded(
  deps: McpBridgeDeps,
  signal: () => void,
  isCancelled: () => boolean,
): Promise<void> {
  const startedAt = Date.now()
  for (;;) {
    if (isCancelled()) return
    const ctx = deps.getCtx()
    if (ctx?.editor && ctx.doc) {
      signal()
      return
    }
    const slow = Date.now() - startedAt > READY_FAST_WINDOW_MS
    await new Promise((resolve) => setTimeout(resolve, slow ? READY_SLOW_POLL_MS : READY_POLL_MS))
  }
}

/** Subscribe the live editor to MCP commands. Returns the unsubscribe function. */
export function installMcpBridge(deps: McpBridgeDeps): () => void {
  const desktop = window.desktop
  if (!desktop?.onMcpCommand || !desktop.reportMcpResult) return () => {}
  let cancelled = false
  let queue: Promise<void> = Promise.resolve()
  const unsubscribe = desktop.onMcpCommand((message: McpCommandMessage) => {
    if (!message || typeof message.requestId !== 'string') return
    queue = queue.then(async () => {
      try {
        const result = await runCommand(deps, message.command, message.payload)
        desktop.reportMcpResult({ requestId: message.requestId, ok: true, result })
      } catch (error) {
        desktop.reportMcpResult({
          requestId: message.requestId,
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        })
      }
    })
  })
  // Let the shell know this tab can accept commands (device for targeted routing).
  void announceWhenLoaded(
    deps,
    () => desktop.signalMcpReady?.(),
    () => cancelled,
  )
  return () => {
    cancelled = true
    unsubscribe()
  }
}
