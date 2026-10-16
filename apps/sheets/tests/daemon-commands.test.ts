/**
 * The daemon's own `office_sheet_*` dispatch, implemented in daemon/commands.ts
 * and reached over the control socket (never the app's built-in MCP server). It
 * rides the same renderer bridge and reuses the same handlers (`McpSheetHandlers`)
 * as the built-in tools, so these tests exercise the dispatcher in isolation.
 */
import { describe, expect, it, vi } from 'vitest'
import type { McpSheetHandlers } from '../src/renderer/mcp-bridge'
import type {
  RangeFormatsResult,
  RangeReadResult,
  SheetMeta,
} from '../src/renderer/daemon/types'
import { dispatchDaemonCommand, isDaemonCommand } from '../src/renderer/daemon'
import { SHEETLESS_OPS } from '../src/renderer/daemon/sheetless-ops'

const SHEETS = [
  { id: 'sheet-1', name: 'Summary', rows: 3, columns: 4 },
  { id: 'jg_nIT-Asw-tL3tvR3f4q', name: 'Data 2024', rows: 500, columns: 12 },
]

/** a worksheet descriptor carrying a freeze of one row, no columns */
function meta(id: string, name: string, rows: number, columns: number): SheetMeta {
  return { id, name, rows, columns, freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 } }
}

/** a minimal range-read result; override the fields a test cares about */
function rangePage(overrides: Partial<RangeReadResult> = {}): RangeReadResult {
  return {
    range: 'A1:B2',
    rows: [],
    formulas: {},
    merges: [],
    rowsShown: 0,
    rowsTotal: 0,
    truncated: false,
    ...overrides,
  }
}

/** a minimal format-read result; override the fields a test cares about */
function formatsPage(overrides: Partial<RangeFormatsResult> = {}): RangeFormatsResult {
  return { range: 'A1:B2', rows: [], merges: [], rowsShown: 0, rowsTotal: 0, truncated: false, ...overrides }
}

function handlersWith(overrides: Partial<McpSheetHandlers> = {}): McpSheetHandlers {
  return {
    hasWorkbook: () => true,
    context: () => ({ sheetId: 'sheet-1', sheets: SHEETS }),
    // no sheet ⇒ the active sheet (sheet-1); the data sheet resolves by its id
    readSheetMeta: (sheetId) =>
      sheetId === 'jg_nIT-Asw-tL3tvR3f4q'
        ? meta('jg_nIT-Asw-tL3tvR3f4q', 'Data 2024', 500, 12)
        : meta('sheet-1', 'Summary', 3, 4),
    readCells: () => ({}),
    readRange: () => rangePage(),
    findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
    readRangeFormats: () => formatsPage(),
    sheets: () => SHEETS,
    focusSheet: vi.fn(),
    applyOps: vi.fn().mockResolvedValue({ ok: true }),
    saveTo: vi.fn().mockResolvedValue({ ok: true }),
    ...overrides,
  }
}

describe('SHEETLESS_OPS', () => {
  it('is the workbook-level ops the schema gives no sheetId', () => {
    expect([...SHEETLESS_OPS].sort()).toEqual([
      'add_defined_name',
      'add_sheet',
      'delete_defined_name',
      'delete_visual',
      'edit_chart',
      'edit_shape',
    ])
  })
})

describe('isDaemonCommand', () => {
  it('claims the office_sheet_* family and nothing else', () => {
    for (const name of [
      'office_sheet_info',
      'office_sheet_read',
      'office_sheet_read_formats',
      'office_sheet_search',
      'office_sheet_edit',
    ]) {
      expect(isDaemonCommand(name)).toBe(true)
    }
    // the built-in bridge's own commands stay with it
    for (const name of ['read_sheet', 'apply_ops', 'save_sheet']) {
      expect(isDaemonCommand(name)).toBe(false)
    }
  })
})

describe('dispatchDaemonCommand', () => {
  it('returns undefined for a command the family does not own', async () => {
    // the caller's own switch handles such names; the dispatcher stays silent
    await expect(dispatchDaemonCommand('read_sheet', {}, handlersWith())).resolves.toBeUndefined()
    await expect(dispatchDaemonCommand('apply_ops', {}, handlersWith())).resolves.toBeUndefined()
  })

  it('reports the worksheets with their extents and the active sheet name', async () => {
    const result = await dispatchDaemonCommand(
      'office_sheet_info',
      {},
      handlersWith({
        // the renderer context is wide; the command must narrow it to name/extent
        context: () => ({ sheetId: 'sheet-1', sheets: SHEETS, selection: 'A1', knownAddresses: ['A1'] }),
      }),
    )
    expect(result).toEqual({
      current: 'Summary',
      sheets: [
        { id: 'sheet-1', name: 'Summary', rows: 3, columns: 4 },
        { id: 'jg_nIT-Asw-tL3tvR3f4q', name: 'Data 2024', rows: 500, columns: 12 },
      ],
    })
  })

  it('defaults a sheet with no known extent to zero rows/columns', async () => {
    const result = await dispatchDaemonCommand(
      'office_sheet_info',
      {},
      handlersWith({ context: () => ({ sheetId: 's', sheets: [{ id: 's', name: 'Blank' }] }) }),
    )
    expect(result).toEqual({
      current: 'Blank',
      sheets: [{ id: 's', name: 'Blank', rows: 0, columns: 0 }],
    })
  })

  it('resolves a sheet NAME to its id before reading', async () => {
    const readSheetMeta = vi.fn(() => meta('jg_nIT-Asw-tL3tvR3f4q', 'Data 2024', 500, 12))
    const result = await dispatchDaemonCommand(
      'office_sheet_read',
      { sheet: 'Data 2024' },
      handlersWith({ readSheetMeta }),
    )
    expect(readSheetMeta).toHaveBeenCalledWith('jg_nIT-Asw-tL3tvR3f4q')
    expect(result).toEqual({
      id: 'jg_nIT-Asw-tL3tvR3f4q',
      name: 'Data 2024',
      rows: 500,
      columns: 12,
      freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
    })
  })

  it('reads the active sheet when no sheet name is given', async () => {
    const readSheetMeta = vi.fn(() => meta('sheet-1', 'Summary', 3, 4))
    await dispatchDaemonCommand('office_sheet_read', {}, handlersWith({ readSheetMeta }))
    // no sheet ⇒ the sheet the view is on, resolved to its id
    expect(readSheetMeta).toHaveBeenCalledWith('sheet-1')
  })

  it('fails loudly on a sheet name that is no longer open', async () => {
    await expect(
      dispatchDaemonCommand('office_sheet_read', { sheet: 'Gone' }, handlersWith()),
    ).rejects.toThrow(/Gone/)
  })

  it('treats `sheet` as a name only — a sheet id is refused', async () => {
    // ids are session artifacts and never a valid `sheet` value
    await expect(
      dispatchDaemonCommand('office_sheet_read', { sheet: 'sheet-1' }, handlersWith()),
    ).rejects.toThrow(/no worksheet named "sheet-1"/)
  })

  it('matches the sheet name exactly — case and spacing matter', async () => {
    await expect(
      dispatchDaemonCommand('office_sheet_read', { sheet: 'summary' }, handlersWith()),
    ).rejects.toThrow(/no worksheet named "summary"/)
    await expect(
      dispatchDaemonCommand('office_sheet_read', { sheet: ' Summary ' }, handlersWith()),
    ).rejects.toThrow(/no worksheet named " Summary "/)
  })

  it('returns the sheet descriptor alone when no range is given', async () => {
    const readRange = vi.fn()
    const result = await dispatchDaemonCommand(
      'office_sheet_read',
      { sheet: 'Summary' },
      handlersWith({ readRange }),
    )
    expect(result).toEqual({
      id: 'sheet-1',
      name: 'Summary',
      rows: 3,
      columns: 4,
      freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
    })
    expect(readRange).not.toHaveBeenCalled()
  })

  it('zeroes the extent and freeze of a sheet with neither recorded', async () => {
    const result = await dispatchDaemonCommand(
      'office_sheet_read',
      { sheet: 'Summary' },
      handlersWith({
        readSheetMeta: () => ({
          id: 'sheet-1',
          name: 'Summary',
          rows: 0,
          columns: 0,
          freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
        }),
      }),
    )
    expect(result).toEqual({
      id: 'sheet-1',
      name: 'Summary',
      rows: 0,
      columns: 0,
      freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
    })
  })

  it('nests the range grid under the descriptor when a range is given', async () => {
    const readRange = vi.fn().mockResolvedValue(rangePage({ range: 'A1:D20' }))
    const result = await dispatchDaemonCommand(
      'office_sheet_read',
      { sheet: 'Summary', range: 'A1:D20', maxRows: 5 },
      handlersWith({ readRange }),
    )
    expect(readRange).toHaveBeenCalledWith('A1:D20', 'sheet-1', { maxRows: 5 })
    expect(result).toEqual({
      id: 'sheet-1',
      name: 'Summary',
      rows: 3,
      columns: 4,
      freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
      range: rangePage({ range: 'A1:D20' }),
    })
  })

  it('routes the format read to the paged range reader, resolved to the sheet id', async () => {
    const readRangeFormats = vi.fn().mockResolvedValue(
      formatsPage({
        range: 'A1:B2',
        rows: [[{ bold: true }, null], [null, null]],
        rowsShown: 2,
        rowsTotal: 2,
      }),
    )
    const result = await dispatchDaemonCommand(
      'office_sheet_read_formats',
      { sheet: 'Summary', range: 'A1:B2' },
      handlersWith({ readRangeFormats }),
    )
    // the name resolved to the id, and no paging option was asked for
    expect(readRangeFormats).toHaveBeenCalledWith('A1:B2', 'sheet-1', {})
    // the descriptor wraps the format grid exactly like office_sheet_read's value grid
    expect(result).toEqual({
      id: 'sheet-1',
      name: 'Summary',
      rows: 3,
      columns: 4,
      freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
      range: formatsPage({
        range: 'A1:B2',
        rows: [[{ bold: true }, null], [null, null]],
        rowsShown: 2,
        rowsTotal: 2,
      }),
    })
  })

  it('forwards maxRows so the format read pages like the value read', async () => {
    const readRangeFormats = vi.fn().mockResolvedValue(
      formatsPage({
        range: 'A1:B5',
        rowsShown: 5,
        rowsTotal: 20,
        truncated: true,
        note: '5 of 20 rows shown; continue with range:"A6:B20"',
      }),
    )
    await dispatchDaemonCommand(
      'office_sheet_read_formats',
      { sheet: 'Summary', range: 'A1:B20', maxRows: 5 },
      handlersWith({ readRangeFormats }),
    )
    expect(readRangeFormats).toHaveBeenCalledWith('A1:B20', 'sheet-1', { maxRows: 5 })
  })

  it('requires a range for the format read', async () => {
    await expect(
      dispatchDaemonCommand('office_sheet_read_formats', { tabId: 't' }, handlersWith()),
    ).rejects.toThrow(/range/)
  })

  it('maps search hits to the read vocabulary and filters them by range', async () => {
    const findCells = vi.fn().mockResolvedValue({
      matches: [
        { sheetName: 'Summary', address: 'A1', value: 'hit', formula: '=1' },
        { sheetName: 'Summary', address: 'Z9', value: 'far' },
      ],
      truncated: false,
      incompleteSheets: [],
    })
    const result = await dispatchDaemonCommand(
      'office_sheet_search',
      { tabId: 't', sheet: 'Summary', query: 'hit', range: 'A1:D20' },
      handlersWith({ findCells }),
    )
    // the sheet name resolved to its id, and the scan stays case-insensitive by default
    expect(findCells).toHaveBeenCalledWith(expect.objectContaining({ sheetId: 'sheet-1', matchCase: false }))
    expect(result).toEqual({
      total: 1,
      matches: [{ sheet: 'Summary', address: 'A1', value: 'hit', formula: '=1' }],
      truncated: false,
    })
  })

  it('forwards matchCase to the find engine', async () => {
    const findCells = vi.fn().mockResolvedValue({ matches: [], truncated: false, incompleteSheets: [] })
    await dispatchDaemonCommand(
      'office_sheet_search',
      { tabId: 't', query: 'Hit', matchCase: true },
      handlersWith({ findCells }),
    )
    expect(findCells).toHaveBeenCalledWith(expect.objectContaining({ matchCase: true }))
  })

  it('searches every worksheet when no sheet is named', async () => {
    const findCells = vi.fn().mockResolvedValue({ matches: [], truncated: false, incompleteSheets: [] })
    await dispatchDaemonCommand(
      'office_sheet_search',
      { tabId: 't', query: 'x' },
      handlersWith({ findCells }),
    )
    // an unnamed sheet stays unfiltered, unlike the single-sheet reads
    expect(findCells).toHaveBeenCalledWith(expect.objectContaining({ sheetId: undefined }))
  })

  it('requires a non-empty query for the search', async () => {
    await expect(
      dispatchDaemonCommand('office_sheet_search', { tabId: 't', query: '  ' }, handlersWith()),
    ).rejects.toThrow(/query/)
  })

  it('decides seeding by op name: a sheet-scoped op without an address is seeded, a workbook-level one is not', async () => {
    const applyOps = vi.fn().mockResolvedValue({ ok: true })
    await dispatchDaemonCommand(
      'office_sheet_edit',
      {
        tabId: 't',
        sheet: 'Summary',
        ops: [
          // no address field, but set_freeze is sheet-scoped and takes a sheetId
          { op: 'set_freeze', rows: 1, columns: 0 },
          // add_sheet is workbook-level: it must stay untouched
          { op: 'add_sheet', name: 'New' },
        ],
      },
      handlersWith({ applyOps }),
    )
    expect(applyOps).toHaveBeenCalledWith(
      [
        { op: 'set_freeze', rows: 1, columns: 0, sheetId: 'sheet-1' },
        { op: 'add_sheet', name: 'New' },
      ],
      false,
    )
  })

  it('seeds an op whose sheet ref is empty or blank, not only an absent one', async () => {
    const applyOps = vi.fn().mockResolvedValue({ ok: true })
    await dispatchDaemonCommand(
      'office_sheet_edit',
      {
        tabId: 't',
        sheet: 'Summary',
        ops: [
          { op: 'set_cell', sheet: '', address: 'A1', value: 1 },
          { op: 'set_formula', sheet: '   ', address: 'B1', formula: '=1' },
        ],
      },
      handlersWith({ applyOps }),
    )
    // an empty ref is stripped like an absent one; the default sheet takes over
    expect(applyOps).toHaveBeenCalledWith(
      [
        { op: 'set_cell', address: 'A1', value: 1, sheetId: 'sheet-1' },
        { op: 'set_formula', address: 'B1', formula: '=1', sheetId: 'sheet-1' },
      ],
      false,
    )
  })

  it('seeds the default sheet into unrefed ops, focuses and applies', async () => {
    const applyOps = vi.fn().mockResolvedValue({ ok: true })
    const focusSheet = vi.fn()
    const result = await dispatchDaemonCommand(
      'office_sheet_edit',
      {
        tabId: 't',
        sheet: 'Data 2024',
        ops: [
          { op: 'set_cell', address: 'C7', value: 5 },
          { op: 'set_cell', sheet: 'Summary', address: 'A1', value: 1 },
        ],
      },
      handlersWith({ applyOps, focusSheet }),
    )
    expect(result).toEqual({ ok: true })
    // the unrefed op got the default sheet id; the explicit name resolved, not seeded
    expect(applyOps).toHaveBeenCalledWith(
      [
        { op: 'set_cell', address: 'C7', value: 5, sheetId: 'jg_nIT-Asw-tL3tvR3f4q' },
        { op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 1 },
      ],
      false,
    )
    // focus follows the first op, on the sheet it changes
    expect(focusSheet).toHaveBeenCalledWith('jg_nIT-Asw-tL3tvR3f4q', 'C7')
  })

  it('seeds the active sheet into unrefed ops when no sheet is named', async () => {
    const applyOps = vi.fn().mockResolvedValue({ ok: true })
    await dispatchDaemonCommand(
      'office_sheet_edit',
      { tabId: 't', ops: [{ op: 'set_cell', address: 'A1', value: 1 }] },
      handlersWith({ applyOps }),
    )
    // no `sheet` ⇒ the tab's active sheet (sheet-1), following the field's semantics
    expect(applyOps).toHaveBeenCalledWith(
      [{ op: 'set_cell', address: 'A1', value: 1, sheetId: 'sheet-1' }],
      false,
    )
  })

  it('plans without touching the grid or focus under dryRun', async () => {
    const applyOps = vi.fn().mockResolvedValue({ ok: true, planned: true })
    const focusSheet = vi.fn()
    await dispatchDaemonCommand(
      'office_sheet_edit',
      { tabId: 't', sheet: 'Summary', ops: [{ op: 'set_cell', address: 'A1', value: 1 }], dryRun: true },
      handlersWith({ applyOps, focusSheet }),
    )
    expect(focusSheet).not.toHaveBeenCalled()
    expect(applyOps).toHaveBeenCalledWith(
      [{ op: 'set_cell', address: 'A1', value: 1, sheetId: 'sheet-1' }],
      true,
    )
  })

  it('rejects an empty batch', async () => {
    await expect(
      dispatchDaemonCommand('office_sheet_edit', { tabId: 't', ops: [] }, handlersWith()),
    ).rejects.toThrow(/ops/)
  })
})
