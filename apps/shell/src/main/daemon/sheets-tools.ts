import { resolveOpenDocumentOfFamily } from '../mcp/tools/open-documents-tools'
import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { SheetsControl } from '../mcp/tools/sheets-tools'
import { asString, listTabs, type DaemonControlDeps } from './shared'

/**
 * The daemon's sheet tools (office_sheet_info / _read / _read_formats / _search /
 * _edit). Once a sheet tab is resolved, the work is forwarded to that tab's
 * renderer through the *same* `SheetsControl.runCommand` the MCP session tools
 * use, so a daemon edit lands in the edit journal, undo history and save
 * pipeline exactly like an in-app one (the renderer dispatches it in
 * `daemon/commands.ts`).
 */

/** resolve the sheet tab a command targets, and check the whole family matches */
async function resolveSheetTab(deps: DaemonControlDeps, args: Record<string, unknown>) {
  const tabId = asString(args.tabId)
  if (!tabId) throw new Error('"tabId" is required (see office_list)')
  const tabs = deps.tabManager
  const doc = resolveOpenDocumentOfFamily(await listTabs(tabs), tabId, 'xlsx')
  const contents = tabs?.webContentsForTab(doc.id)
  if (!contents || contents.isDestroyed()) throw new Error(`"${doc.title}" is no longer open`)
  return { doc, wcId: contents.id }
}

/**
 * Run one sheet command in the target tab's renderer (payload carries no tabId),
 * and report the tab it ran on so a caller can read the tab's own state too.
 */
async function runSheetOnTab(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
  command: string,
): Promise<{ doc: OpenDocumentTab; result: Record<string, unknown> }> {
  const { doc, wcId } = await resolveSheetTab(deps, args)
  const { tabId: _tabId, ...payload } = args
  const result = await deps.sheets.runCommand(
    wcId,
    command as Parameters<SheetsControl['runCommand']>[1],
    payload,
  )
  return {
    doc,
    result: result && typeof result === 'object' ? (result as Record<string, unknown>) : { result },
  }
}

/** run one sheet command in the target tab's renderer (payload carries no tabId) */
export async function runSheet(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
  command: string,
): Promise<Record<string, unknown>> {
  return (await runSheetOnTab(deps, args, command)).result
}

/**
 * office_sheet_info stamped with the tab's unsaved-changes flag: `dirty` is true
 * while the workbook has edits not yet written to disk, the same flag office_list
 * reports for every tab. The other sheet reads leave it off.
 */
export async function sheetInfoTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const { doc, result } = await runSheetOnTab(deps, args, 'office_sheet_info')
  return { ...result, dirty: doc.dirty }
}
