/**
 * The shell's daemon's own `office_*` channel (createDaemonControl): the handler
 * an out-of-process daemon reaches over the control socket. The tab list, save
 * and close each go through the open-documents control that already has its own
 * coverage; this pins `create` (writes a blank file of the requested format and
 * opens it) and `open` (opens an existing file in a tab), each returning the
 * tabId later calls address.
 */
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import type { ControlReply } from '@genoffice/cli/control-protocol'
import type { OpenDocumentTab } from '../src/shared/tabs-api'
import type { SheetsControl } from '../src/main/mcp/tools/sheets-tools'
import type { OpenDocumentsControl } from '../src/main/mcp/tools/open-documents-tools'
import type { TabManager } from '../src/main/tab-manager'
import { createDaemonControl, type DaemonControlDeps } from '../src/main/daemon/daemon-control'

function tempFile(name: string): string {
  const path = join(mkdtempSync(join(tmpdir(), 'genoffice-office-open-')), name)
  writeFileSync(path, '')
  return path
}

/** a one-page PDF carrying `text`, written with pdf-lib (the pdf app's own dep) */
async function pdfFile(name: string, text: string): Promise<string> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  doc.addPage([400, 300]).drawText(text, { x: 40, y: 200, size: 14, font })
  const path = join(mkdtempSync(join(tmpdir(), 'genoffice-office-pdf-')), name)
  writeFileSync(path, await doc.save())
  return path
}

function tab(id: string, kind: OpenDocumentTab['kind'], filePath: string): OpenDocumentTab {
  return { id, kind, title: 'x', filePath, active: true, dirty: false }
}

function errorOf(reply: ControlReply): { message: string } {
  if (reply.ok) throw new Error('expected a failure reply')
  return reply.error
}

function resultOf(reply: ControlReply): Record<string, unknown> {
  if (!reply.ok) throw new Error(`expected an ok reply, got: ${reply.error.message}`)
  return reply.result
}

/** a TabManager stand-in: the methods the control derives its APIs from */
function tabs(overrides: Partial<TabManager> = {}): TabManager {
  return {
    openDocuments: async () => [],
    webContentsForTab: () => undefined,
    findTabByPath: () => undefined,
    activateTab: vi.fn(),
    openDocsTab: vi.fn(() => 't1'),
    openSheetsTab: vi.fn(() => 't1'),
    openSlidesTab: vi.fn(() => 't1'),
    openPdfTab: vi.fn(() => 't1'),
    ...overrides,
  } as unknown as TabManager
}

function depsWith(overrides: Partial<DaemonControlDeps> = {}): DaemonControlDeps {
  return {
    tabManager: tabs(),
    sheets: { runCommand: vi.fn() } as unknown as SheetsControl,
    documents: {} as unknown as OpenDocumentsControl,
    revealShellWindow: vi.fn(),
    ...overrides,
  }
}

/** the control under test, wired the way the shell wires it: deps read per call */
function control(overrides: Partial<DaemonControlDeps> = {}) {
  return createDaemonControl(() => depsWith(overrides))
}

describe('office_create', () => {
  it('writes the requested format, then waits for a workbook to mount', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-office-create-'))
    const file = join(dir, 'Book1.xlsx')
    const openSheetsTab = vi.fn()
    const runCommand = vi.fn().mockResolvedValue({ id: 'sheet-1' })
    const reply = await control({
      tabManager: tabs({
        openSheetsTab,
        findTabByPath: () => ({ id: 't1', kind: 'sheets', webContents: { id: 7 } as never }),
        openDocuments: async () => [tab('t1', 'sheets', file)],
        webContentsForTab: () => ({ id: 7, isDestroyed: () => false }) as never,
      }),
      sheets: { runCommand } as unknown as SheetsControl,
    }).handle('create', { path: dir, name: 'Book1', type: 'xlsx' })
    expect(reply).toEqual({ ok: true, result: { path: file, tabId: 't1' } })
    // the blank file is written, then opened by format
    expect(existsSync(file)).toBe(true)
    expect(openSheetsTab).toHaveBeenCalledWith(file)
    expect(runCommand).toHaveBeenCalledWith(7, 'read_sheet', {})
  })

  it('writes a document, presentation or PDF without a workbook wait', async () => {
    for (const [type, kind, opener] of [
      ['docx', 'docs', 'openDocsTab'],
      ['pptx', 'slides', 'openSlidesTab'],
      ['pdf', 'pdf', 'openPdfTab'],
    ] as const) {
      const dir = mkdtempSync(join(tmpdir(), `genoffice-office-create-${type}-`))
      const file = join(dir, `New.${type}`)
      const runCommand = vi.fn()
      const open = vi.fn()
      const reply = await control({
        tabManager: tabs({
          ...({ [opener]: open } as Partial<TabManager>),
          findTabByPath: () => ({ id: 't1', kind, webContents: { id: 9 } as never }),
          openDocuments: async () => [tab('t1', kind, file)],
          webContentsForTab: () => ({ id: 9, isDestroyed: () => false }) as never,
        }),
        sheets: { runCommand } as unknown as SheetsControl,
      }).handle('create', { path: dir, name: 'New', type })
      expect(reply).toEqual({ ok: true, result: { path: file, tabId: 't1' } })
      expect(open).toHaveBeenCalledWith(file)
      // only a sheet tab has a workbook to mount; the others have nothing to wait for
      expect(runCommand).not.toHaveBeenCalled()
    }
  })

  it('refuses a format this build cannot write', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-office-create-'))
    const reply = await control().handle('create', { path: dir, name: 'X', type: 'odt' })
    expect(errorOf(reply).message).toContain('unsupported type')
  })
})

describe('office_open', () => {
  it('opens an existing file, waits for a workbook to mount, and returns its sheets', async () => {
    const file = tempFile('Book1.xlsx')
    const activateTab = vi.fn()
    const runCommand = vi.fn().mockResolvedValue({ currentSheetId: 's1', sheets: [] })
    const reply = await control({
      tabManager: tabs({
        activateTab,
        findTabByPath: () => ({ id: 't1', kind: 'sheets', webContents: { id: 7 } as never }),
        openDocuments: async () => [tab('t1', 'sheets', file)],
        webContentsForTab: () => ({ id: 7, isDestroyed: () => false }) as never,
      }),
      sheets: { runCommand } as unknown as SheetsControl,
    }).handle('open', { path: file })
    // the workbook's sheet info rides along with the tab id
    expect(reply).toEqual({
      ok: true,
      result: { path: file, tabId: 't1', info: { currentSheetId: 's1', sheets: [] } },
    })
    // an already-open file focuses its tab instead of opening a second copy
    expect(activateTab).toHaveBeenCalledWith('t1')
    // the readiness wait doubles as the workbook read the caller wants
    expect(runCommand).toHaveBeenCalledWith(7, 'office_sheet_info', {})
  })

  it('does not wait on a tab that is not a workbook', async () => {
    const file = tempFile('Deck.pptx')
    const runCommand = vi.fn()
    const reply = await control({
      tabManager: tabs({
        findTabByPath: () => ({ id: 't2', kind: 'slides', webContents: { id: 8 } as never }),
        openDocuments: async () => [tab('t2', 'slides', file)],
        webContentsForTab: () => ({ id: 8, isDestroyed: () => false }) as never,
      }),
      sheets: { runCommand } as unknown as SheetsControl,
    }).handle('open', { path: file })
    expect(reply).toEqual({ ok: true, result: { path: file, tabId: 't2' } })
    expect(runCommand).not.toHaveBeenCalled()
  })

  it('refuses a missing path, a relative path, or a file nothing can open', async () => {
    const handle = control().handle
    expect(errorOf(await handle('open', {})).message).toContain('required')
    expect(errorOf(await handle('open', { path: 'Book1.xlsx' })).message).toContain('absolute')
    expect(
      errorOf(await handle('open', { path: join(tmpdir(), 'genoffice-nope-xyz.xlsx') })).message,
    ).toContain('no such file')
    // the file exists but no editor owns its format
    expect(errorOf(await handle('open', { path: tempFile('notes.xyz') })).message).toContain(
      'could not open',
    )
  })

  it('reports a tab that never came back', async () => {
    // the file opens (openSheetsTab runs) but no tab can be found afterwards
    const file = tempFile('Book1.xlsx')
    const reply = await control().handle('open', { path: file })
    expect(errorOf(reply).message).toContain('no tab')
  })
})

describe('office_pdf_read / office_pdf_search', () => {
  it('reads a PDF by an absolute path, no window needed', async () => {
    const file = await pdfFile('report.pdf', 'Hello PDF')
    const result = resultOf(await control().handle('pdf_read', { path: file }))
    // the file is read directly (the pdf app is a viewer): path, page count, text
    expect(result.path).toBe(file)
    expect(result.pageCount).toBe(1)
    expect((result.pages as Array<{ text: string }>)[0]!.text).toContain('Hello PDF')
  })

  it('reads an open PDF tab by its tabId', async () => {
    const file = await pdfFile('open.pdf', 'Open tab text')
    const result = resultOf(
      await control({
        tabManager: tabs({ openDocuments: async () => [tab('p1', 'pdf', file)] }),
      }).handle('pdf_read', { tabId: 'p1' }),
    )
    expect(result.path).toBe(file)
    expect((result.pages as Array<{ text: string }>)[0]!.text).toContain('Open tab text')
  })

  it('searches a PDF and addresses hits by 1-based page', async () => {
    const file = await pdfFile('search.pdf', 'find the needle here')
    const result = resultOf(await control().handle('pdf_search', { path: file, query: 'needle' }))
    expect(result.total).toBe(1)
    expect((result.matches as Array<{ page: number }>)[0]!.page).toBe(1)
  })

  it('refuses a missing target, a non-PDF path, and an unsaved PDF tab', async () => {
    const handle = control().handle
    expect(errorOf(await handle('pdf_read', {})).message).toContain('tabId')
    expect(errorOf(await handle('pdf_read', { path: tempFile('notes.txt') })).message).toContain(
      'no open tab',
    )
    const unsaved = control({
      tabManager: tabs({
        openDocuments: async () => [
          { id: 'p1', kind: 'pdf', title: 'Untitled', filePath: undefined, active: true, dirty: true },
        ],
      }),
    })
    expect(errorOf(await unsaved.handle('pdf_read', { tabId: 'p1' })).message).toContain(
      'never been saved',
    )
  })

  it('requires a query to search', async () => {
    expect(errorOf(await control().handle('pdf_search', {})).message).toContain('query')
  })
})
