import { describe, expect, it } from 'vitest'
import { z } from 'zod'
import { GUI_TOOLS, GUI_TOOL_NAMES } from '../src/gui-tools'
import { CLI_PREFIX, GUI_PREFIX } from '../src/names'

describe('gui tool table', () => {
  it('names every GUI tool office_gui_* with a slug-safe name', () => {
    expect(GUI_TOOLS.length).toBeGreaterThan(0)
    for (const tool of GUI_TOOLS) {
      expect(tool.name.startsWith(GUI_PREFIX)).toBe(true)
      expect(tool.name).toMatch(/^office_gui_[a-z0-9_]+$/)
    }
  })

  it('keeps the CLI prefix and the GUI prefix distinct', () => {
    expect(CLI_PREFIX).toBe('office_cli_')
    expect(GUI_PREFIX).toBe('office_gui_')
    expect(CLI_PREFIX).not.toBe(GUI_PREFIX)
  })

  it('uses unique exposed names and unique upstream names', () => {
    const exposed = GUI_TOOLS.map((t) => t.name)
    const upstream = GUI_TOOLS.map((t) => t.upstream)
    expect(new Set(exposed).size).toBe(exposed.length)
    expect(new Set(upstream).size).toBe(upstream.length)
  })

  it('leaves upstream names bare (no prefix leaks into the app)', () => {
    for (const tool of GUI_TOOLS) {
      expect(tool.upstream.startsWith('office_')).toBe(false)
      expect(tool.upstream).toMatch(/^[a-z][a-z0-9_]*$/)
    }
  })

  it('maps every exposed name to its upstream name', () => {
    expect(GUI_TOOL_NAMES.size).toBe(GUI_TOOLS.length)
    for (const tool of GUI_TOOLS) {
      expect(GUI_TOOL_NAMES.get(tool.name)).toBe(tool.upstream)
    }
    expect(GUI_TOOL_NAMES.get('office_gui_open')).toBe('open_in_genoffice')
  })

  it('describes every tool and gives each a zod input shape', () => {
    for (const tool of GUI_TOOLS) {
      expect(tool.description.trim().length).toBeGreaterThan(0)
      expect(tool.shape).toBeTypeOf('object')
      for (const value of Object.values(tool.shape)) expect(value).toBeInstanceOf(z.ZodType)
    }
  })

  it('leaves the headless generator to the CLI family and keeps only the visible session', () => {
    const upstream = GUI_TOOLS.map((t) => t.upstream)
    // `create` is the CLI's headless generator (office_cli_create); the visible
    // session opener is a different tool the app serves.
    expect(upstream).not.toContain('create')
    expect(upstream).toContain('create_session')
  })

  it('leaves the file-only reader to the CLI family and keeps the GUI reads path-free', () => {
    const upstream = GUI_TOOLS.map((t) => t.upstream)
    // read_docx reads a file on disk with no window; the daemon serves it headlessly
    // as office_cli_docs_read, so it is not re-exposed here.
    expect(upstream).not.toContain('read_docx')
    // a `path` in the GUI family is only ever the whole file a window opens or
    // writes: `open` (input) and `save_session` (output). Every read/edit target
    // is a live tab named by `document`, never a path.
    const pathTools = new Set(['open_in_genoffice', 'save_session'])
    for (const tool of GUI_TOOLS) {
      if (Object.keys(tool.shape).includes('path')) {
        expect(pathTools).toContain(tool.upstream)
      }
    }
  })

  it('exposes read_pdf as a document-addressed read whose adapter resolves the file', () => {
    const tool = GUI_TOOLS.find((t) => t.name === 'office_gui_read_pdf')
    expect(tool?.upstream).toBe('read_pdf')
    expect(tool?.readOnly).toBe(true)
    expect(tool?.shape).toHaveProperty('document')
    expect(tool?.shape).not.toHaveProperty('path')
    // required (a PDF has no session to fall back on)
    expect(z.object(tool!.shape).safeParse({}).success).toBe(false)
    expect(tool?.adapt).toBeTypeOf('function')
  })

  it('keeps read_pdf’s bounded read: pages is accepted and forwarded', () => {
    const tool = GUI_TOOLS.find((t) => t.name === 'office_gui_read_pdf')!
    // the upstream read_pdf takes `pages`; dropping it here would silently
    // force every daemon read to extract the whole document
    expect(tool.shape).toHaveProperty('pages')
    expect(z.object(tool.shape).safeParse({ document: 'x', pages: '1-5,8' }).success).toBe(true)
  })

  it('keeps read_sheet’s range read: range and maxRows are accepted and forwarded', () => {
    const tool = GUI_TOOLS.find((t) => t.name === 'office_gui_read_sheet')!
    // the upstream read_sheet takes `range`/`maxRows`; dropping them here would
    // force every daemon range read back to the ≤100-address path
    expect(tool.shape).toHaveProperty('range')
    expect(tool.shape).toHaveProperty('maxRows')
    expect(z.object(tool.shape).safeParse({ range: 'A1:D20', maxRows: 5 }).success).toBe(true)
  })

  it('exposes one search tool per family, each shaped like its read tool', () => {
    const byName = new Map(GUI_TOOLS.map((t) => [t.name, t]))
    const specs = [
      {
        name: 'office_gui_search_document',
        upstream: 'search_document',
        extra: ['wholeWord', 'regex', 'context'],
      },
      {
        name: 'office_gui_search_sheet',
        upstream: 'search_sheet',
        extra: ['sheet', 'sheetId', 'lookIn'],
      },
      { name: 'office_gui_search_deck', upstream: 'search_deck', extra: ['slide', 'context'] },
      { name: 'office_gui_search_pdf', upstream: 'search_pdf', extra: ['pages', 'context'] },
    ]
    for (const spec of specs) {
      const tool = byName.get(spec.name)
      expect(tool?.upstream).toBe(spec.upstream)
      expect(tool?.readOnly).toBe(true)
      // a GUI search names a live document, never a bare path
      expect(tool?.shape).toHaveProperty('document')
      expect(tool?.shape).toHaveProperty('query')
      expect(tool?.shape).not.toHaveProperty('path')
      for (const field of spec.extra) expect(tool?.shape).toHaveProperty(field)
      // query is required on every one
      expect(z.object(tool!.shape).safeParse({}).success).toBe(false)
    }
    // the three session-backed searches fall back to the session's document; the
    // pdf search has no session, so its document is required
    for (const name of [
      'office_gui_search_document',
      'office_gui_search_sheet',
      'office_gui_search_deck',
    ]) {
      expect(z.object(byName.get(name)!.shape).safeParse({ query: 'x' }).success).toBe(true)
    }
    const pdf = byName.get('office_gui_search_pdf')!
    expect(z.object(pdf.shape).safeParse({ query: 'x' }).success).toBe(false)
    expect(z.object(pdf.shape).safeParse({ document: 'd', query: 'x' }).success).toBe(true)
    // only the pdf search reshapes document -> path (its upstream is path-based,
    // like read_pdf); the others forward their arguments verbatim
    expect(pdf.adapt).toBeTypeOf('function')
    for (const name of [
      'office_gui_search_document',
      'office_gui_search_sheet',
      'office_gui_search_deck',
    ]) {
      expect(byName.get(name)?.adapt).toBeUndefined()
    }
  })
})
