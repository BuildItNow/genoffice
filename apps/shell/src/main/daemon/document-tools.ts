import { existsSync, rmSync, statSync } from 'node:fs'
import { extname, isAbsolute, join } from 'node:path'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { buildBlankDocx } from '@genoffice/docx-engine'
import { createBlankPptx } from '@genoffice/pptx-engine'
import { blankPdfBuffer } from '../../../../pdf/src/main/blank-pdf'
import { resolveOpenDocument } from '../mcp/tools/open-documents-tools'
import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { TabManager } from '../tab-manager'
import { atomicWriteFile } from '../atomic-write'
import { asString, listTabs, type DaemonControlDeps } from './shared'
import { sheetInfoTool } from './sheets-tools'
import { pdfInfoTool } from './pdf-tools'

/**
 * The daemon's family-agnostic tools — list / create / open / close / save: they
 * act on a tab whatever it holds, unlike the per-family sheet_* and pdf_* sets.
 * create and open need a tab strip to put the file in; close and save reuse the
 * `open_documents` control's per-family writer/cleanup, so they behave exactly
 * like the built-in ones across every editor family.
 */

/** the formats `office_create` can write a blank file for (the editors it opens) */
export const FILE_TYPES = ['xlsx', 'docx', 'pptx', 'pdf'] as const
export type CreateFileType = (typeof FILE_TYPES)[number]

/** the tab manager, or a clear failure while no window is open */
function requireTabs(tabs: TabManager | null): TabManager {
  if (!tabs) throw new Error('GenOffice has no window open')
  return tabs
}

/** write a fresh empty file of `type` at `path` (throws if it already exists) */
async function createBlankFile(path: string, type: CreateFileType): Promise<void> {
  if (existsSync(path)) throw new Error(`file already exists: ${path}`)
  // one blank-file builder per format the shell can open
  let blank: Uint8Array
  switch (type) {
    case 'xlsx':
      blank = await blankXlsxBuffer()
      break
    case 'docx':
      blank = await buildBlankDocx()
      break
    case 'pptx':
      blank = await createBlankPptx()
      break
    case 'pdf':
      blank = await blankPdfBuffer()
      break
    default:
      throw new Error(`unsupported file type: ${type}`)
  }
  await atomicWriteFile(path, blank)
}

/** open a file in a tab: focus the tab already showing it, else open by format */
function openInTab(tabs: TabManager, path: string): boolean {
  const existing = tabs.findTabByPath(path)
  if (existing) {
    tabs.activateTab(existing.id)
    return true
  }
  switch (extname(path).toLowerCase()) {
    case '.docx':
      tabs.openDocsTab(path)
      return true
    case '.xlsx':
      tabs.openSheetsTab(path)
      return true
    case '.pptx':
      tabs.openSlidesTab(path)
      return true
    case '.pdf':
      tabs.openPdfTab(path)
      return true
    default:
      return false
  }
}

/** one `office_list` row (a tab addressed by tabId on later calls) */
function describeTab(doc: OpenDocumentTab): Record<string, unknown> {
  return {
    tabId: doc.id,
    kind: doc.kind,
    title: doc.title,
    path: doc.filePath ?? null,
    active: doc.active,
    dirty: doc.dirty,
  }
}

/** office_list: the open tabs, each as the row later calls address by tabId */
export async function listTool(deps: DaemonControlDeps): Promise<Record<string, unknown>> {
  return { tabs: (await listTabs(deps.tabManager)).map(describeTab) }
}

/** office_create: write a blank file of the requested format and open it */
export async function createTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  // `path` is the folder, `name` the stem, `type` the format: the tool names the
  // three so each can be checked on its own (a directory that exists, a name
  // that is not a path, a format this build can write)
  const dir = asString(args.path)
  if (!dir) throw new Error('"path" is required (the folder to create the file in)')
  if (!isAbsolute(dir)) throw new Error('"path" must be an absolute folder')
  const name = asString(args.name)
  if (!name) throw new Error('"name" is required (the file name without its extension)')
  if (/[\\/]/.test(name)) {
    throw new Error(`"name" must be a plain file name, not a path: ${name}`)
  }
  const fileType = asString(args.type) ?? ''
  if (!(FILE_TYPES as readonly string[]).includes(fileType)) {
    throw new Error(`unsupported type "${fileType}" (this build creates ${FILE_TYPES.join(', ')})`)
  }
  if (!existsSync(dir) || !statSync(dir).isDirectory()) {
    throw new Error(`"path" is not an existing folder: ${dir}`)
  }
  const tabs = requireTabs(deps.tabManager)
  const path = join(dir, `${name}.${fileType}`)
  await createBlankFile(path, fileType as CreateFileType)
  try {
    if (!openInTab(tabs, path)) throw new Error(`GenOffice could not open ${path}`)
    // a created file is only useful once it is a tab the caller can address
    const tabId = tabs.findTabByPath(path)?.id
    if (!tabId) throw new Error(`GenOffice opened ${path} but reported no tab for it`)
    // Opening a tab is synchronous, but mounting its workbook is not.
    // `runCommand` waits for the tab's renderer before it sends, so any command
    // here makes that wait happen now: a tab that never comes up fails `create`
    // instead of the caller's next call. Only a sheet tab has a workbook to
    // mount; the other families have nothing to wait for. (`read_sheet` is the
    // cheapest command that waits.)
    const doc = resolveOpenDocument(await listTabs(tabs), tabId)
    if (doc?.kind === 'sheets') {
      const contents = tabs.webContentsForTab(tabId)
      if (contents) {
        try {
          await deps.sheets.runCommand(contents.id, 'read_sheet', {})
        } catch (error) {
          // the wait failed for some other reason (a closed tab, a renderer that
          // never came up) — name the real problem, not the read we only used to
          // force the wait
          throw new Error(
            `GenOffice opened ${path} but its spreadsheet did not become ready: ${
              error instanceof Error ? error.message : String(error)
            }`,
          )
        }
      }
    }
    return { path, tabId }
  } catch (error) {
    // the tab never became usable — don't leave the file we just wrote
    try {
      rmSync(path, { force: true })
    } catch (removeError) {
      console.warn('[office] could not remove the unusable created file:', removeError)
    }
    throw error
  }
}

/** office_open: open an existing file in a tab and return its tabId; a workbook
 *  also reports its sheet info, and a PDF its page count/metadata/preview, so
 *  the caller can act without a separate info call */
export async function openTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  // the complement of `create`; the file is the user's, so unlike `create` a
  // failure never removes it
  const path = asString(args.path)
  if (!path) throw new Error('"path" is required (the file to open)')
  if (!isAbsolute(path)) throw new Error('"path" must be an absolute path')
  if (!existsSync(path)) throw new Error(`no such file: ${path}`)
  const tabs = requireTabs(deps.tabManager)
  if (!openInTab(tabs, path)) throw new Error(`GenOffice could not open ${path}`)
  // `openInTab` focuses the tab it found or opened; read its id back so the
  // caller can address it with the per-tab office_* tools
  const tabId = tabs.findTabByPath(path)?.id
  if (!tabId) throw new Error(`GenOffice opened ${path} but reported no tab for it`)
  // a sheet tab is only usable once its workbook has mounted; this info both
  // forces that wait (so the caller's next sheet call does not fail instead) and
  // hands back the workbook's sheets. Other families have nothing to wait for.
  const doc = resolveOpenDocument(await listTabs(tabs), tabId)
  if (doc?.kind === 'sheets') {
    try {
      const info = await sheetInfoTool(deps, { tabId })
      return { path, tabId, info }
    } catch (error) {
      throw new Error(
        `GenOffice opened ${path} but its spreadsheet did not become ready: ${
          error instanceof Error ? error.message : String(error)
        }`,
      )
    }
  }
  // a PDF is a viewer: its info reads the file straight from disk, so unlike the
  // sheet case there is no renderer to wait for — just hand back the descriptor,
  // naming the file if the read fails
  if (doc?.kind === 'pdf') {
    try {
      const info = await pdfInfoTool(deps, { tabId })
      return { path, tabId, info }
    } catch (error) {
      throw new Error(
        `GenOffice opened ${path} but its PDF could not be read: ${
          error instanceof Error ? error.message : String(error)
        }`,
      )
    }
  }
  return { path, tabId }
}

/** office_close: close a tab, saving it on the way out unless discard is set */
export async function closeTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const tabId = asString(args.tabId)
  if (!tabId) throw new Error('"tabId" is required (see office_list)')
  // any editor family can close, so match on tab id/path without the family
  // check; office_list already lists what is open
  const doc = resolveOpenDocument(await listTabs(deps.tabManager), tabId)
  if (!doc) throw new Error(`no open tab matches "${tabId}" (see office_list)`)
  // `discard` defaults to false: an omitted flag saves the document on its way
  // out, the same default as the built-in open_documents close. The per-family
  // settle/save/cleanup is the built-in control's, so a daemon close lands
  // exactly like the one the app offers.
  const { savedPath } = await deps.documents.close(doc, {
    unsaved: args.discard === true ? 'discard' : 'save',
  })
  return { tabId: doc.id, closed: true, ...(savedPath ? { savedPath } : {}) }
}

/** office_save: write a tab's live content back to its own file */
export async function saveTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const tabId = asString(args.tabId)
  if (!tabId) throw new Error('"tabId" is required (see office_list)')
  // save is family-agnostic like close: the open_documents control owns the
  // per-family writer (docs/sheets/slides/markdown/html; a PDF is a viewer and
  // refuses), so route through it rather than the sheets bridge
  const doc = resolveOpenDocument(await listTabs(deps.tabManager), tabId)
  if (!doc) throw new Error(`no open tab matches "${tabId}" (see office_list)`)
  const { savedPath } = await deps.documents.save(doc)
  return { tabId: doc.id, savedPath }
}
