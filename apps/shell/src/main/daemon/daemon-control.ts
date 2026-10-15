import { existsSync, rmSync, statSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { extname, isAbsolute, join } from 'node:path'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { buildBlankDocx } from '@genoffice/docx-engine'
import { createBlankPptx } from '@genoffice/pptx-engine'
import type { ControlReply } from '@genoffice/cli/control-protocol'
import { blankPdfBuffer } from '../../../../pdf/src/main/blank-pdf'
import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { SheetsControl } from '../mcp/tools/sheets-tools'
import type { OpenDocumentsControl } from '../mcp/tools/open-documents-tools'
import { resolveOpenDocument, resolveOpenDocumentOfFamily } from '../mcp/tools/open-documents-tools'
import { readPdfFile, searchPdfFile } from '../mcp/tools/pdf-tools'
import type { TabManager } from '../tab-manager'
import { atomicWriteFile } from '../atomic-write'

/**
 * The shell's implementation of the daemon's own `office_*` tool channel (the
 * `daemon` control command). These tools reach the running app directly over
 * the control socket rather than through the app's built-in MCP server, so the
 * built-in server and its tool list are untouched.
 *
 * The whole surface is built from three collaborators the shell injects: the
 * tab manager (the tab list, opening a file, resolving a `tabId` to its
 * webContents), the sheets renderer channel, and the `open_documents` control
 * (whose per-family writer/cleanup `save` and `close` reuse). Every API the
 * tools need — list, open, create, resolve — is derived here from those three;
 * the only callback the shell injects is `revealShellWindow`, to put a window
 * back when the user closed the last one, and the module pulls in no app main.
 *
 * Once a sheet tab is resolved, the work is forwarded to that tab's renderer
 * through the *same* `SheetsControl.runCommand` the MCP session tools use, so a
 * daemon edit lands in the edit journal, undo history and save pipeline exactly
 * like an in-app one (the renderer dispatches it in `daemon/commands.ts`).
 *
 * The PDF reads (`pdf_read` / `pdf_search`) run right here instead: pdfium is a
 * main-process WASM module, so a PDF is read in the shell from its file (an open
 * tab's, or any `.pdf` path) with the same helpers the built-in `read_pdf` /
 * `search_pdf` use — no renderer round-trip, and no app-side tool is touched.
 */

export interface DaemonControlDeps {
  /** the shell's tab strip; null only while no window is open (macOS, last window closed) */
  tabManager: TabManager | null
  /** the sheets renderer channel (createSheetsControl in index.ts) */
  sheets: SheetsControl
  /**
   * the `open_documents` control (createOpenDocumentsControl in index.ts): save
   * and close reuse its per-family writer/cleanup, so they behave exactly like
   * the built-in ones, across every editor family (not just sheets).
   */
  documents: OpenDocumentsControl
  /**
   * the shell's revealShellWindow: bring the shell window up, rebuilding it when
   * the user closed it (macOS keeps the process alive). Used by the tools that
   * create a tab, which need a window to put it in.
   */
  revealShellWindow: () => void
}

export interface OfficeControl {
  /** run one `office_*` tool (the bare name, e.g. `sheet_read`) against the live app */
  handle: (tool: string, args: Record<string, unknown>) => Promise<ControlReply>
}

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

const ok = (result: Record<string, unknown>): ControlReply => ({ ok: true, result })
const fail = (message: string): ControlReply => ({
  ok: false,
  error: { reason: 'app_unavailable', message },
})

/** the tabs the user can see, or none while no window is open */
async function listTabs(tabs: TabManager | null): Promise<OpenDocumentTab[]> {
  return tabs ? tabs.openDocuments() : []
}

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

/** run one sheet command in the target tab's renderer (payload carries no tabId) */
async function runSheet(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
  command: string,
): Promise<Record<string, unknown>> {
  const { wcId } = await resolveSheetTab(deps, args)
  const { tabId: _tabId, ...payload } = args
  const result = await deps.sheets.runCommand(
    wcId,
    command as Parameters<SheetsControl['runCommand']>[1],
    payload,
  )
  return result && typeof result === 'object' ? (result as Record<string, unknown>) : { result }
}

/** office_list: the open tabs, each as the row later calls address by tabId */
async function listTool(deps: DaemonControlDeps): Promise<Record<string, unknown>> {
  return { tabs: (await listTabs(deps.tabManager)).map(describeTab) }
}

/** office_create: write a blank file of the requested format and open it */
async function createTool(
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
 *  also reports its sheet info, so the caller can address it without a
 *  separate office_sheet_info call */
async function openTool(
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
      const info = await runSheet(deps, { tabId }, 'office_sheet_info')
      return { path, tabId, info }
    } catch (error) {
      throw new Error(
        `GenOffice opened ${path} but its spreadsheet did not become ready: ${
          error instanceof Error ? error.message : String(error)
        }`,
      )
    }
  }
  return { path, tabId }
}

/** office_close: close a tab, saving it on the way out unless discard is set */
async function closeTool(
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
async function saveTool(
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

/**
 * Resolve the file a PDF tool targets: an open PDF tab (a `tabId`, or a `path`
 * that names one) read from its own file, else an absolute `.pdf` path on disk
 * that need not be open. A PDF has no live editor state (the pdf app is a
 * viewer), so both address the file — matching how office_open resolves a tab.
 */
async function resolvePdfPath(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<string> {
  const ref = asString(args.tabId) ?? asString(args.path)
  if (!ref) throw new Error('pass the PDF as "tabId" (see office_list) or an absolute "path"')
  const doc = deps.tabManager ? resolveOpenDocument(await listTabs(deps.tabManager), ref) : undefined
  if (doc) {
    if (doc.kind !== 'pdf') throw new Error(`"${doc.title}" is not a PDF`)
    if (doc.filePath === undefined) {
      throw new Error(`"${doc.title}" has never been saved, so there is no file to read`)
    }
    return doc.filePath
  }
  if (isAbsolute(ref) && extname(ref).toLowerCase() === '.pdf') {
    if (!existsSync(ref)) throw new Error(`no such file: ${ref}`)
    return ref
  }
  throw new Error(
    `no open tab matches "${ref}"; open a PDF in GenOffice, or pass an absolute .pdf path`,
  )
}

/** office_pdf_read: read a PDF's text (a preview, or the named pages) */
async function pdfReadTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const path = await resolvePdfPath(deps, args)
  const bytes = new Uint8Array(await readFile(path))
  return readPdfFile(bytes, {
    path,
    pages: typeof args.pages === 'string' ? args.pages : undefined,
  })
}

/** office_pdf_search: find text in a PDF's text layer, hits addressed by page */
async function pdfSearchTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const query = asString(args.query)
  if (!query) throw new Error('"query" is required (the text to find)')
  const path = await resolvePdfPath(deps, args)
  const bytes = new Uint8Array(await readFile(path))
  const result = await searchPdfFile(bytes, {
    query,
    pages: typeof args.pages === 'string' ? args.pages : undefined,
    matchCase: args.matchCase === true,
    maxResults: typeof args.maxResults === 'number' ? args.maxResults : undefined,
    context: typeof args.context === 'number' ? args.context : undefined,
  })
  return { ...result, total: result.matches.length }
}

export function createDaemonControl(getDeps: () => DaemonControlDeps): OfficeControl {
  return {
    handle: async (tool, args) => {
      try {
        // create and open need a tab strip to put the file in, so make sure the
        // shell window is up first — the daemon starts the app (upstream.ts), but
        // on macOS the process outlives the last window it closed. Re-read the
        // deps afterwards: rebuilding the window rebuilds the tab manager too.
        if ((tool === 'create' || tool === 'open') && !getDeps().tabManager) {
          getDeps().revealShellWindow()
        }
        // read the deps per call: the shell rebuilds them (a new tabManager after
        // the last window closes), so a captured object would go stale
        const deps = getDeps()
        switch (tool) {
          case 'list':
            return ok(await listTool(deps))
          case 'create':
            return ok(await createTool(deps, args))
          case 'open':
            return ok(await openTool(deps, args))
          case 'close':
            return ok(await closeTool(deps, args))
          case 'save':
            return ok(await saveTool(deps, args))
          case 'sheet_info':
            return ok(await runSheet(deps, args, 'office_sheet_info'))
          case 'sheet_read':
            return ok(await runSheet(deps, args, 'office_sheet_read'))
          case 'sheet_read_formats':
            return ok(await runSheet(deps, args, 'office_sheet_read_formats'))
          case 'sheet_search':
            return ok(await runSheet(deps, args, 'office_sheet_search'))
          case 'sheet_edit':
            return ok(await runSheet(deps, args, 'office_sheet_edit'))
          case 'pdf_read':
            return ok(await pdfReadTool(deps, args))
          case 'pdf_search':
            return ok(await pdfSearchTool(deps, args))
          default:
            return fail(`unknown office tool "${tool}"`)
        }
      } catch (error) {
        return fail(error instanceof Error ? error.message : String(error))
      }
    },
  }
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() !== '' ? value.trim() : undefined
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
