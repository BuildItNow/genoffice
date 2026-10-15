import { describe, expect, it } from 'vitest'
import { z } from 'zod'
import { OFFICE_TOOLS, OFFICE_TOOL_NAMES } from '../src/office/tools'
import { CLI_PREFIX, OFFICE_PREFIX } from '../src/names'

describe('office tool table', () => {
  it('keeps bare names slug-safe and exposes them as office_*', () => {
    expect(OFFICE_TOOLS.length).toBeGreaterThan(0)
    for (const tool of OFFICE_TOOLS) {
      // the table holds the bare name the shell routes on; the exposed name is prefixed
      expect(tool.name).toMatch(/^[a-z][a-z0-9_]*$/)
      expect(tool.name.startsWith('office_')).toBe(false)
      expect(`${OFFICE_PREFIX}${tool.name}`).toMatch(/^office_[a-z0-9_]+$/)
    }
  })

  it('keeps the daemon prefix distinct from the CLI family', () => {
    expect(OFFICE_PREFIX).toBe('office_')
    expect(OFFICE_PREFIX).not.toBe(CLI_PREFIX)
    // the daemon prefix is a strict prefix of the CLI family: office_cli_x must
    // not be mistaken for a bare office_* tool.
    expect(CLI_PREFIX.startsWith(OFFICE_PREFIX)).toBe(true)
  })

  it('uses unique bare names and maps every exposed name to its bare one', () => {
    const names = OFFICE_TOOLS.map((t) => t.name)
    expect(new Set(names).size).toBe(names.length)
    expect(OFFICE_TOOL_NAMES.size).toBe(OFFICE_TOOLS.length)
    for (const tool of OFFICE_TOOLS) {
      expect(OFFICE_TOOL_NAMES.get(`${OFFICE_PREFIX}${tool.name}`)).toBe(tool.name)
    }
    expect(OFFICE_TOOL_NAMES.get('office_sheet_read')).toBe('sheet_read')
  })

  it('describes every tool and gives each a zod input shape', () => {
    for (const tool of OFFICE_TOOLS) {
      expect(tool.description.trim().length).toBeGreaterThan(0)
      expect(tool.shape).toBeTypeOf('object')
      for (const value of Object.values(tool.shape)) expect(value).toBeInstanceOf(z.ZodType)
    }
  })

  it('leaves the writes read-write and the reads read-only', () => {
    const byName = new Map(OFFICE_TOOLS.map((t) => [t.name, t]))
    for (const name of [
      'list',
      'sheet_info',
      'sheet_read',
      'sheet_read_formats',
      'sheet_search',
      'sheet_edit_help',
      'pdf_read',
      'pdf_search',
    ]) {
      expect(byName.get(name)?.readOnly).toBe(true)
    }
    for (const name of ['create', 'open', 'close', 'save', 'sheet_edit']) {
      expect(byName.get(name)?.readOnly).not.toBe(true)
    }
  })

  it('addresses every tool by a bare tabId except list, create, open and the app-free reference', () => {
    const byName = new Map(OFFICE_TOOLS.map((t) => [t.name, t]))
    // list enumerates the tabs, so it takes no target
    expect(byName.get('list')?.shape).not.toHaveProperty('tabId')
    expect(z.object(byName.get('list')!.shape).safeParse({}).success).toBe(true)
    // create names the file, not a tab
    expect(byName.get('create')?.shape).not.toHaveProperty('tabId')
    // open names an existing file, not a tab (it returns the tab it opened)
    expect(byName.get('open')?.shape).not.toHaveProperty('tabId')
    expect(byName.get('open')?.shape).toHaveProperty('path')
    expect(z.object(byName.get('open')!.shape).safeParse({}).success).toBe(false)
    expect(z.object(byName.get('open')!.shape).safeParse({ path: '/tmp/Book1.xlsx' }).success).toBe(
      true,
    )
    // the op reference is static metadata; it names categories/ops, not a tab
    expect(byName.get('sheet_edit_help')?.shape).not.toHaveProperty('tabId')
    expect(byName.get('sheet_edit_help')?.shape).toHaveProperty('list')
    expect(byName.get('sheet_edit_help')?.shape).toHaveProperty('ops')
    // both are optional: no arguments lists every category
    const helpSchema = z.object(byName.get('sheet_edit_help')!.shape)
    expect(helpSchema.safeParse({}).success).toBe(true)
    expect(helpSchema.safeParse({ ops: ['set_cell'] }).success).toBe(true)
    expect(helpSchema.safeParse({ list: ['format'], ops: ['format_range'] }).success).toBe(true)
    // everything else acts on a live tab
    for (const name of ['close', 'save', 'sheet_info', 'sheet_read', 'sheet_read_formats', 'sheet_search', 'sheet_edit']) {
      expect(byName.get(name)?.shape).toHaveProperty('tabId')
      expect(z.object(byName.get(name)!.shape).safeParse({}).success).toBe(false)
    }
  })

  it('serves the sheet edit reference in the daemon, with no app', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_edit_help')!
    expect(help.local).toBeTypeOf('function')
    const result = (await help.local!({ ops: ['format_range', 'set_cell'] })) as {
      categories?: unknown[]
      ops: Array<{ op: string; group: string; signature: string; note?: string }>
      unknown?: { ops: unknown[] }
    }
    expect(result.ops.map((e) => e.op)).toEqual(['format_range', 'set_cell'])
    expect(result.ops[0]).toMatchObject({ op: 'format_range', group: 'format' })
    expect(result.ops[0]!.signature).toContain('range')
    // the catalog's note rides along
    expect(result.ops[0]!.note).toContain('null clears a property')
    // only ops were asked for, and every name was known: nothing else comes back
    expect(result.categories).toBeUndefined()
    expect(result.unknown).toBeUndefined()
  })

  it('lists a category\'s ops by name, and nothing when asked for nothing', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_edit_help')!
    const one = (await help.local!({ list: ['format'] })) as {
      categories: Array<{ name: string; summary?: string; ops: string[] }>
      ops?: unknown[]
    }
    expect(one.categories).toHaveLength(1)
    expect(one.categories[0]).toMatchObject({ name: 'format' })
    expect(one.categories[0]!.ops).toContain('format_range')
    expect(one.ops).toBeUndefined()
    // no arguments ⇒ an empty result
    expect(await help.local!({})).toEqual({})
  })

  it('reports an unknown name on its own, with a suggestion, instead of failing', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_edit_help')!
    const result = (await help.local!({ list: ['formt'], ops: ['set_cel'] })) as {
      ops: unknown[]
      unknown: {
        categories: Array<{ name: string; suggestion?: string }>
        ops: Array<{ op: string; suggestion?: string }>
      }
    }
    expect(result.ops).toEqual([])
    expect(result.unknown.categories).toEqual([{ name: 'formt', suggestion: 'format' }])
    expect(result.unknown.ops).toEqual([{ op: 'set_cel', suggestion: 'set_cell' }])
  })

  it('flags an op only the live editor can run', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_edit_help')!
    const result = (await help.local!({ ops: ['add_table_row'] })) as {
      ops: Array<{ op: string; requiresEditorSession?: string }>
    }
    expect(result.ops[0]!.op).toBe('add_table_row')
    expect(result.ops[0]!.requiresEditorSession).toBeTruthy()
  })

  it('closes a tab by tabId, with discard defaulting to false', () => {
    const close = OFFICE_TOOLS.find((t) => t.name === 'close')!
    const schema = z.object(close.shape)
    // the tab is required; discard is optional, so an omitted flag is the default
    expect(schema.safeParse({}).success).toBe(false)
    const parsed = schema.safeParse({ tabId: 't1' })
    expect(parsed.success).toBe(true)
    expect(parsed.success && parsed.data.discard).toBeUndefined()
    expect(schema.safeParse({ tabId: 't1', discard: true }).success).toBe(true)
  })

  it('creates only the formats this build supports, path/name/type required', () => {
    const create = OFFICE_TOOLS.find((t) => t.name === 'create')!
    expect(create.shape).toHaveProperty('path')
    expect(create.shape).toHaveProperty('name')
    expect(create.shape).toHaveProperty('type')
    const schema = z.object(create.shape)
    expect(schema.safeParse({}).success).toBe(false)
    // the folder alone is not enough: name and type are named separately
    expect(schema.safeParse({ path: '/tmp' }).success).toBe(false)
    expect(schema.safeParse({ path: '/tmp', name: 'Book1' }).success).toBe(false)
    // every editor family this build can open has a blank-file builder
    for (const type of ['xlsx', 'docx', 'pptx', 'pdf']) {
      expect(schema.safeParse({ path: '/tmp', name: 'New', type }).success).toBe(true)
    }
    // a format outside the enum is refused
    expect(schema.safeParse({ path: '/tmp', name: 'Doc1', type: 'odt' }).success).toBe(false)
  })

  it('addresses the PDF reads by an open tab or an absolute path, never a required tabId', () => {
    const byName = new Map(OFFICE_TOOLS.map((t) => [t.name, t]))
    for (const name of ['pdf_read', 'pdf_search']) {
      expect(byName.get(name)?.shape).toHaveProperty('tabId')
      expect(byName.get(name)?.shape).toHaveProperty('path')
    }
    // the read needs only its target; either a tab or a path is enough
    const read = z.object(byName.get('pdf_read')!.shape)
    expect(read.safeParse({ tabId: 't1' }).success).toBe(true)
    expect(read.safeParse({ path: '/tmp/report.pdf' }).success).toBe(true)
    expect(read.safeParse({ tabId: 't1', pages: '1-5' }).success).toBe(true)
    // the search also needs a query
    const search = z.object(byName.get('pdf_search')!.shape)
    expect(search.safeParse({ path: '/tmp/a.pdf' }).success).toBe(false)
    expect(search.safeParse({ path: '/tmp/a.pdf', query: 'x' }).success).toBe(true)
  })

  it('defaults sheetId and scopes the sheet reads to an optional worksheet', () => {
    const byName = new Map(OFFICE_TOOLS.map((t) => [t.name, t]))
    // sheet_info reports the whole workbook; the reads/edits pick a worksheet
    expect(byName.get('sheet_info')?.shape).not.toHaveProperty('sheetId')
    for (const name of ['sheet_read', 'sheet_read_formats', 'sheet_search', 'sheet_edit']) {
      expect(byName.get(name)?.shape).toHaveProperty('sheetId')
    }
    // sheet_read_formats needs a range; sheet_read does not
    const formats = z.object(byName.get('sheet_read_formats')!.shape)
    expect(formats.safeParse({ tabId: 't' }).success).toBe(false)
    expect(formats.safeParse({ tabId: 't', range: 'A1:D20' }).success).toBe(true)
    expect(z.object(byName.get('sheet_read')!.shape).safeParse({ tabId: 't' }).success).toBe(true)
  })
})
