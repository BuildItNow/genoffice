import { resolveOpenDocumentOfFamily } from '../mcp/tools/open-documents-tools'
import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { DocsControl } from '../mcp/tools/document-tools'
import { asString, listTabs, type DaemonControlDeps } from './shared'

/**
 * The daemon's doc tools (office_doc_info / _read / _search, office_doc_apply_ops
 * and office_doc_apply_func). Once a docx tab is resolved, the work is
 * forwarded to that tab's renderer through the *same* `DocsControl.runCommand`
 * the MCP docx session tools use, so a daemon edit lands in the edit journal,
 * undo history and save pipeline exactly like an in-app or MCP one (the renderer
 * dispatches it in apps/docs/src/renderer/daemon/commands.ts).
 *
 * Unlike a PDF (a viewer read from disk), a Word document's live editor state —
 * the block indexes an edit addresses, the unsaved text — only exists in its
 * renderer, so every tool here requires an open tab and none reads from a path.
 */

/** resolve the docx tab a command targets, and check the family matches */
async function resolveDocTab(deps: DaemonControlDeps, args: Record<string, unknown>) {
  const tabId = asString(args.tabId)
  if (!tabId) throw new Error('"tabId" is required (see office_list)')
  const tabs = deps.tabManager
  const doc = resolveOpenDocumentOfFamily(await listTabs(tabs), tabId, 'docx')
  const contents = tabs?.webContentsForTab(doc.id)
  if (!contents || contents.isDestroyed()) throw new Error(`"${doc.title}" is no longer open`)
  return { doc, wcId: contents.id }
}

/**
 * Run one doc command in the target tab's renderer (payload carries no tabId),
 * and report the tab it ran on so a caller can read the tab's own state too.
 */
async function runDocOnTab(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
  command: string,
): Promise<{ doc: OpenDocumentTab; result: Record<string, unknown> }> {
  const { doc, wcId } = await resolveDocTab(deps, args)
  const { tabId: _tabId, ...payload } = args
  const result = await deps.docs.runCommand(
    wcId,
    command as Parameters<DocsControl['runCommand']>[1],
    payload,
  )
  return {
    doc,
    result: result && typeof result === 'object' ? (result as Record<string, unknown>) : { result },
  }
}

/** run one doc command in the target tab's renderer (payload carries no tabId) */
export async function runDoc(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
  command: string,
): Promise<Record<string, unknown>> {
  return (await runDocOnTab(deps, args, command)).result
}

/**
 * office_doc_info: the renderer's document descriptor (block count, headings
 * outline, text stats) stamped with the tab's identity and unsaved-changes flag —
 * `dirty` is true while the document has edits not yet written to disk, the same
 * flag office_list reports for every tab.
 */
export async function docInfoTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const { doc, result } = await runDocOnTab(deps, args, 'office_doc_info')
  return { ...result, title: doc.title, path: doc.filePath ?? null, dirty: doc.dirty }
}
