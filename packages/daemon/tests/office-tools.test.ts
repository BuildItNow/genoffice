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
      'sheet_ops_help',
      'doc_info',
      'doc_read',
      'doc_search',
      'doc_ops_help',
      'doc_func_help',
      'pdf_info',
      'pdf_read',
      'pdf_search',
    ]) {
      expect(byName.get(name)?.readOnly).toBe(true)
    }
    for (const name of [
      'create',
      'open',
      'close',
      'save',
      'sheet_apply_ops',
      'doc_apply_ops',
      'doc_apply_func',
    ]) {
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
    expect(byName.get('sheet_ops_help')?.shape).not.toHaveProperty('tabId')
    expect(byName.get('sheet_ops_help')?.shape).toHaveProperty('list')
    expect(byName.get('sheet_ops_help')?.shape).toHaveProperty('ops')
    // both are optional: no arguments lists every category
    const helpSchema = z.object(byName.get('sheet_ops_help')!.shape)
    expect(helpSchema.safeParse({}).success).toBe(true)
    expect(helpSchema.safeParse({ ops: ['set_cell'] }).success).toBe(true)
    expect(helpSchema.safeParse({ list: ['format'], ops: ['format_range'] }).success).toBe(true)
    // the doc op reference is static metadata too: categories/ops, not a tab
    expect(byName.get('doc_ops_help')?.shape).not.toHaveProperty('tabId')
    expect(byName.get('doc_ops_help')?.shape).toHaveProperty('list')
    expect(byName.get('doc_ops_help')?.shape).toHaveProperty('ops')
    const docHelpSchema = z.object(byName.get('doc_ops_help')!.shape)
    expect(docHelpSchema.safeParse({}).success).toBe(true)
    expect(docHelpSchema.safeParse({ ops: ['setFont'] }).success).toBe(true)
    // the func reference is static metadata too: func names, not a tab
    expect(byName.get('doc_func_help')?.shape).not.toHaveProperty('tabId')
    expect(byName.get('doc_func_help')?.shape).toHaveProperty('funcs')
    expect(z.object(byName.get('doc_func_help')!.shape).safeParse({}).success).toBe(true)
    // everything else acts on a live tab
    for (const name of [
      'close',
      'save',
      'sheet_info',
      'sheet_read',
      'sheet_read_formats',
      'sheet_search',
      'sheet_apply_ops',
      'doc_info',
      'doc_read',
      'doc_search',
      'doc_apply_ops',
      'doc_apply_func',
    ]) {
      expect(byName.get(name)?.shape).toHaveProperty('tabId')
      expect(z.object(byName.get(name)!.shape).safeParse({}).success).toBe(false)
    }
  })

  it('serves the sheet ops reference in the daemon, with no app', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_ops_help')!
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

  it("lists a category's ops by name, and nothing when asked for nothing", async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_ops_help')!
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
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_ops_help')!
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
    const help = OFFICE_TOOLS.find((t) => t.name === 'sheet_ops_help')!
    const result = (await help.local!({ ops: ['add_table_row'] })) as {
      ops: Array<{ op: string; requiresEditorSession?: string }>
    }
    expect(result.ops[0]!.op).toBe('add_table_row')
    expect(result.ops[0]!.requiresEditorSession).toBeTruthy()
  })

  it('serves the doc block-op reference in the daemon, with no app', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'doc_ops_help')!
    expect(help.local).toBeTypeOf('function')
    const result = (await help.local!({ ops: ['setFont', 'setHeadingLevel'] })) as {
      categories?: unknown[]
      ops: Array<{ op: string; group: string; signature: string; note?: string }>
      unknown?: { ops: unknown[] }
    }
    expect(result.ops.map((e) => e.op)).toEqual(['setFont', 'setHeadingLevel'])
    expect(result.ops[0]).toMatchObject({ op: 'setFont', group: 'block' })
    expect(result.ops[0]!.signature).toContain('target')
    // the catalog's note rides along
    expect(result.ops[0]!.note).toContain('run-level style')
    // only ops were asked for, and every name was known: nothing else comes back
    expect(result.categories).toBeUndefined()
    expect(result.unknown).toBeUndefined()
  })

  it('serves the doc op category index, expanding one category at a time', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'doc_ops_help')!
    // no arguments ⇒ just the index: names and summaries, no op names
    const index = (await help.local!({})) as {
      categories: Array<{ name: string; summary: string; ops?: string[] }>
    }
    expect(index.categories.map((c) => c.name)).toEqual([
      'text',
      'paragraph',
      'list',
      'structure',
      'table',
      'field',
      'picture',
    ])
    expect(index.categories.every((c) => c.ops === undefined)).toBe(true)
    // a named category carries its op names
    const one = (await help.local!({ list: ['text'] })) as {
      categories: Array<{ name: string; summary?: string; ops: string[] }>
      ops?: unknown[]
    }
    expect(one.categories).toHaveLength(1)
    expect(one.categories[0]).toMatchObject({ name: 'text' })
    expect(one.categories[0]!.ops).toEqual(['setFont', 'setMatchedFont', 'findReplace'])
    expect(one.ops).toBeUndefined()
    // a func name is not an apply_ops op: it comes back under unknown
    const funcAsOp = (await help.local!({ ops: ['insert_image'] })) as {
      ops: unknown[]
      unknown?: { ops: Array<{ op: string; suggestion?: string }> }
    }
    expect(funcAsOp.ops).toEqual([])
    expect(funcAsOp.unknown?.ops).toEqual([{ op: 'insert_image' }])
  })

  it('groups every block op in the catalog exactly once', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'doc_ops_help')!
    const index = (await help.local!({})) as { categories: Array<{ name: string }> }
    const listed = (await help.local!({
      list: index.categories.map((c) => c.name),
    })) as { categories: Array<{ name: string; ops: string[] }> }
    const grouped = listed.categories.flatMap((c) => c.ops)
    // each op appears once across the categories…
    expect(new Set(grouped).size).toBe(grouped.length)
    // …and covers exactly what apply_ops can run (the CLI catalog's block group)
    const { docHelp } = await import('@genoffice/cli/doc-catalog')
    const catalog = await docHelp({ list: ['block'] })
    const blockOps = catalog.categories!.find((c) => c.name === 'block')!.ops
    expect([...grouped].sort()).toEqual([...blockOps].sort())
  })

  it('reports an unknown doc op on its own, with a suggestion', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'doc_ops_help')!
    const result = (await help.local!({ ops: ['setFnt'] })) as {
      ops: unknown[]
      unknown: { ops: Array<{ op: string; suggestion?: string }> }
    }
    expect(result.ops).toEqual([])
    expect(result.unknown.ops).toEqual([{ op: 'setFnt', suggestion: 'setFont' }])
  })

  it('serves the doc func category index, expanding one tier at a time', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'doc_func_help')!
    expect(help.local).toBeTypeOf('function')
    // no arguments ⇒ just the index: names and summaries, no func names, no schemas
    const index = (await help.local!({})) as {
      categories: Array<{ name: string; summary: string; funcs?: string[] }>
      funcs?: unknown[]
    }
    expect(index.categories.map((c) => c.name)).toEqual([
      'content',
      'layout',
      'styles',
      'comments',
      'notes',
      'revisions',
    ])
    expect(index.categories.every((c) => c.funcs === undefined)).toBe(true)
    expect(index.funcs).toBeUndefined()
    // a named category carries its func names, but still no schemas
    const one = (await help.local!({ list: ['comments'] })) as {
      categories: Array<{ name: string; summary: string; funcs: string[] }>
      funcs?: unknown[]
    }
    expect(one.categories).toHaveLength(1)
    expect(one.categories[0]).toMatchObject({ name: 'comments' })
    expect(one.categories[0]!.funcs).toEqual([
      'add_comment',
      'reply_comment',
      'resolve_comment',
      'delete_comment',
      'read_comments',
    ])
    expect(one.funcs).toBeUndefined()
    // named funcs carry their signature with schema and example
    const detail = (await help.local!({ funcs: ['insert_picture'] })) as {
      categories?: unknown
      funcs: Array<{
        func: string
        signature: string
        schema?: { required?: string[] }
        example?: unknown
      }>
      unknown?: unknown
    }
    expect(detail.funcs).toHaveLength(1)
    expect(detail.funcs[0]!.signature).toContain('url')
    expect(detail.funcs[0]!.schema?.required).toContain('url')
    expect(detail.categories).toBeUndefined()
    expect(detail.unknown).toBeUndefined()
  })

  it('groups every func in the enum exactly once', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'doc_func_help')!
    const apply = OFFICE_TOOLS.find((t) => t.name === 'doc_apply_func')!
    const funcNames = (apply.shape.func as unknown as { options: string[] }).options
    const index = (await help.local!({})) as { categories: Array<{ name: string }> }
    const listed = (await help.local!({
      list: index.categories.map((c) => c.name),
    })) as { categories: Array<{ name: string; funcs: string[] }> }
    const grouped = listed.categories.flatMap((c) => c.funcs)
    expect(new Set(grouped).size).toBe(grouped.length)
    expect([...grouped].sort()).toEqual([...funcNames].sort())
    // the op batch is the one doc op that is not a func
    expect(grouped).not.toContain('apply_ops')
  })

  it('reports an unknown func and category with suggestions', async () => {
    const help = OFFICE_TOOLS.find((t) => t.name === 'doc_func_help')!
    const result = (await help.local!({ list: ['comment'], funcs: ['ad_comment'] })) as {
      categories: unknown[]
      funcs: unknown[]
      unknown: {
        categories: Array<{ name: string; suggestion?: string }>
        funcs: Array<{ func: string; suggestion?: string }>
      }
    }
    expect(result.categories).toEqual([])
    expect(result.funcs).toEqual([])
    expect(result.unknown.categories).toEqual([{ name: 'comment', suggestion: 'comments' }])
    expect(result.unknown.funcs).toEqual([{ func: 'ad_comment', suggestion: 'add_comment' }])
  })

  it('addresses the doc reads by a tabId, and doc_search by a query too', () => {
    const byName = new Map(OFFICE_TOOLS.map((t) => [t.name, t]))
    // doc_read's range/full/maxChars/offset are all optional: the overview needs only the tab
    expect(z.object(byName.get('doc_read')!.shape).safeParse({ tabId: 't' }).success).toBe(true)
    expect(
      z.object(byName.get('doc_read')!.shape).safeParse({ tabId: 't', range: '0-2' }).success,
    ).toBe(true)
    // doc_search needs the query on top of the tab
    const search = z.object(byName.get('doc_search')!.shape)
    expect(search.safeParse({ tabId: 't' }).success).toBe(false)
    expect(search.safeParse({ tabId: 't', query: 'x' }).success).toBe(true)
  })

  it('takes a doc apply_ops batch of op objects, with dryRun opt-in', () => {
    const edit = z.object(OFFICE_TOOLS.find((t) => t.name === 'doc_apply_ops')!.shape)
    expect(edit.safeParse({ tabId: 't' }).success).toBe(false)
    const parsed = edit.safeParse({ tabId: 't', ops: [{ op: 'setFont' }] })
    expect(parsed.success).toBe(true)
    expect(parsed.success && parsed.data.dryRun).toBeUndefined()
    expect(edit.safeParse({ tabId: 't', ops: [{ op: 'setFont' }], dryRun: true }).success).toBe(
      true,
    )
  })

  it('runs doc funcs through one tool, validating each against its own spec', async () => {
    const byName = new Map(OFFICE_TOOLS.map((t) => [t.name, t]))
    // every op tool but the batch is a func, not a tool of its own
    for (const name of [
      'doc_insert_content',
      'doc_replace_blocks',
      'doc_read_comments',
      'doc_read_revisions',
      'doc_read_notes',
      'doc_list_styles',
    ]) {
      expect(byName.has(name)).toBe(false)
    }

    // one generic entry: args is a free record until prepare() checks the func
    const apply = byName.get('doc_apply_func')!
    const schema = z.object(apply.shape)
    expect(schema.safeParse({ tabId: 't' }).success).toBe(false)
    expect(schema.safeParse({ tabId: 't', func: 'insert_picture' }).success).toBe(true)
    expect(schema.safeParse({ tabId: 't', func: 'nope' }).success).toBe(false)
    const parsed = (await apply.prepare!({
      tabId: 't',
      func: 'add_comment',
      args: { blockIndex: 2, comment: 'why?' },
    })) as { args: Record<string, unknown> }
    expect(parsed.args).toEqual({ blockIndex: 2, comment: 'why?' })
    // the func's own required fields are enforced
    expect(() =>
      apply.prepare!({ tabId: 't', func: 'add_comment', args: { blockIndex: 2 } }),
    ).toThrow(/comment/)
    // unknown fields are rejected, not silently stripped
    expect(() =>
      apply.prepare!({
        tabId: 't',
        func: 'add_comment',
        args: { blockIndex: 2, comment: 'x', body: 'y' },
      }),
    ).toThrow(/body/)
    expect(() => apply.prepare!({ tabId: 't', func: 'nope', args: {} })).toThrow(/unknown func/)

    // the content writes validate their own fields
    expect(() => apply.prepare!({ tabId: 't', func: 'insert_content', args: {} })).toThrow(/html/)
    expect(
      (await apply.prepare!({
        tabId: 't',
        func: 'replace_blocks',
        args: { startBlockIndex: 0, endBlockIndex: 1, html: '<p>x</p>' },
      })) as { args: Record<string, unknown> },
    ).toMatchObject({ args: { startBlockIndex: 0, endBlockIndex: 1 } })

    // the pure reads need no fields at all
    for (const name of ['read_comments', 'read_revisions', 'read_notes', 'list_styles']) {
      const parsed = (await apply.prepare!({ tabId: 't', func: name })) as {
        args: Record<string, unknown>
      }
      expect(parsed.args).toEqual({})
    }
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
    for (const name of ['pdf_info', 'pdf_read', 'pdf_search']) {
      expect(byName.get(name)?.shape).toHaveProperty('tabId')
      expect(byName.get(name)?.shape).toHaveProperty('path')
    }
    // the info needs only its target; either a tab or a path is enough
    const info = z.object(byName.get('pdf_info')!.shape)
    expect(info.safeParse({ tabId: 't1' }).success).toBe(true)
    expect(info.safeParse({ path: '/tmp/report.pdf' }).success).toBe(true)
    // the read needs its target AND the pages to read (the preview is pdf_info's job)
    const read = z.object(byName.get('pdf_read')!.shape)
    expect(read.safeParse({ tabId: 't1' }).success).toBe(false)
    expect(read.safeParse({ tabId: 't1', pages: '1-5' }).success).toBe(true)
    expect(read.safeParse({ path: '/tmp/report.pdf', pages: '3' }).success).toBe(true)
    // the search also needs a query
    const search = z.object(byName.get('pdf_search')!.shape)
    expect(search.safeParse({ path: '/tmp/a.pdf' }).success).toBe(false)
    expect(search.safeParse({ path: '/tmp/a.pdf', query: 'x' }).success).toBe(true)
  })

  it('addresses the sheet reads by NAME and scopes them to an optional worksheet', () => {
    const byName = new Map(OFFICE_TOOLS.map((t) => [t.name, t]))
    // sheet_info reports the whole workbook; the reads/edits pick a worksheet
    expect(byName.get('sheet_info')?.shape).not.toHaveProperty('sheetId')
    // the sheet tools address the worksheet by NAME (ids are session artifacts)
    for (const name of ['sheet_read', 'sheet_read_formats', 'sheet_search', 'sheet_apply_ops']) {
      expect(byName.get(name)?.shape).toHaveProperty('sheet')
      expect(byName.get(name)?.shape).not.toHaveProperty('sheetId')
    }
    // sheet_read_formats needs a range; sheet_read does not
    const formats = z.object(byName.get('sheet_read_formats')!.shape)
    expect(formats.safeParse({ tabId: 't' }).success).toBe(false)
    expect(formats.safeParse({ tabId: 't', range: 'A1:D20' }).success).toBe(true)
    expect(z.object(byName.get('sheet_read')!.shape).safeParse({ tabId: 't' }).success).toBe(true)
  })

  it('lets the search opt into case-sensitive matching', () => {
    const search = z.object(OFFICE_TOOLS.find((t) => t.name === 'sheet_search')!.shape)
    // optional, so an omitted flag leaves the default case-insensitive scan
    expect(search.safeParse({ tabId: 't', query: 'x' }).success).toBe(true)
    expect(search.safeParse({ tabId: 't', query: 'x', matchCase: true }).success).toBe(true)
  })

  it('scopes sheet_read_formats by a combinable category list, defaulting to the common groups', () => {
    const formats = z.object(OFFICE_TOOLS.find((t) => t.name === 'sheet_read_formats')!.shape)
    // an omitted category defaults to the common groups, leaving the per-edge border opt-in
    expect(formats.parse({ tabId: 't', range: 'A1:D20' }).category).toEqual([
      'font',
      'fill',
      'alignment',
      'number',
    ])
    // several groups combine, and `all` is available for the whole cell
    expect(
      formats.safeParse({ tabId: 't', range: 'A1:D20', category: ['font', 'border'] }).success,
    ).toBe(true)
    expect(formats.safeParse({ tabId: 't', range: 'A1:D20', category: ['all'] }).success).toBe(true)
    // an empty list reads nothing, so it is refused
    expect(formats.safeParse({ tabId: 't', range: 'A1:D20', category: [] }).success).toBe(false)
    // a name outside the enum is refused
    expect(formats.safeParse({ tabId: 't', range: 'A1:D20', category: ['nope'] }).success).toBe(
      false,
    )
  })
})
