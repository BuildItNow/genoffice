/**
 * The MCP apply_ops boundary validates raw ops against the workbook DSL schema
 * (the same gate the built-in AI path applies). Without it, an op missing its
 * sheetId surfaces as a cryptic "Unknown sheet: undefined" from the planner
 * instead of naming the op and pointing at read_sheet.
 */
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { installSheetsMcpBridge, type McpSheetHandlers } from '../src/renderer/mcp-bridge'

type CommandListener = (message: {
  requestId: number
  command: string
  payload: Record<string, unknown>
}) => void

let commandListener: CommandListener | undefined
const reportMcpResult = vi.fn()

beforeEach(() => {
  commandListener = undefined
  reportMcpResult.mockReset()
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: {
      onMcpCommand: (listener: CommandListener) => {
        commandListener = listener
        return () => {}
      },
      reportMcpResult,
      signalMcpReady: () => {},
    },
  }
})

function bridgeWith(overrides: Partial<McpSheetHandlers> = {}): McpSheetHandlers {
  return {
    hasWorkbook: () => true,
    context: () => ({}),
    readCells: () => ({}),
    readRange: () => ({}),
    findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
    sheets: () => [{ id: 'sheet-1', name: 'Sheet1' }],
    focusSheet: () => {},
    applyOps: vi.fn().mockResolvedValue({ ok: true }),
    saveTo: vi.fn().mockResolvedValue({ ok: true }),
    ...overrides,
  }
}

async function sendCommand(
  command: string,
  payload: Record<string, unknown>,
  overrides: Partial<McpSheetHandlers> = {},
): Promise<void> {
  const uninstall = installSheetsMcpBridge(bridgeWith(overrides))
  commandListener!({ requestId: 1, command, payload })
  uninstall()
  await vi.waitFor(() => expect(reportMcpResult).toHaveBeenCalledTimes(1))
}

async function applyOps(payload: Record<string, unknown>): Promise<void> {
  await sendCommand('apply_ops', payload)
}

describe('mcp bridge apply_ops validation', () => {
  it('rejects an op missing sheetId with the op named and a read_sheet hint', async () => {
    await applyOps({ ops: [{ op: 'set_cell', address: 'A1', value: 'x' }] })
    const reply = reportMcpResult.mock.calls[0]![0] as { ok: boolean; error?: string }
    expect(reply.ok).toBe(false)
    expect(reply.error).toContain('operations[0] (set_cell)')
    expect(reply.error).toContain('none of the 1 operation(s) were applied')
    expect(reply.error).toContain('sheetId')
    expect(reply.error).toContain('read_sheet')
  })

  it('passes schema-valid ops through to the apply handler', async () => {
    const applyOpsFn = vi.fn().mockResolvedValue({ ok: true, applied: true })
    const uninstall = installSheetsMcpBridge(bridgeWith({ applyOps: applyOpsFn }))
    commandListener!({
      requestId: 2,
      command: 'apply_ops',
      payload: { ops: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'x' }] },
    })
    uninstall()
    await vi.waitFor(() => expect(reportMcpResult).toHaveBeenCalledTimes(1))
    expect(applyOpsFn).toHaveBeenCalledTimes(1)
    const reply = reportMcpResult.mock.calls[0]![0] as { ok: boolean }
    expect(reply.ok).toBe(true)
  })
})

describe('mcp bridge search_sheet', () => {
  it('forwards query, regex, lookIn and maxResults to the find engine', async () => {
    const findCells = vi.fn().mockReturnValue({
      matches: [{ sheetName: 'Sheet1', address: 'A1', value: 'sum' }],
      truncated: false,
      incompleteSheets: [],
    })
    await sendCommand(
      'search_sheet',
      { query: 'sum', regex: true, lookIn: 'formulas', maxResults: 5 },
      { findCells },
    )
    expect(findCells).toHaveBeenCalledWith({
      query: 'sum',
      regex: true,
      lookIn: 'formulas',
      errorsOnly: false,
      sheetId: undefined,
      maxResults: 5,
    })
    const reply = reportMcpResult.mock.calls[0]![0] as { ok: boolean; result?: unknown }
    expect(reply.ok).toBe(true)
    // the reply renames the engine's sheetName to the read_sheet vocabulary (`sheet`)
    expect(reply.result).toEqual({
      total: 1,
      matches: [{ sheet: 'Sheet1', address: 'A1', value: 'sum' }],
      truncated: false,
    })
  })

  it('scopes the hits to an A1 range after the engine scans', async () => {
    const findCells = vi.fn().mockReturnValue({
      matches: [
        { sheetName: 'Sheet1', address: 'A1', value: 'x' },
        { sheetName: 'Sheet1', address: 'Z9', value: 'x' },
      ],
      truncated: false,
      incompleteSheets: [],
    })
    await sendCommand('search_sheet', { query: 'x', range: 'A1:D20' }, { findCells })
    const reply = reportMcpResult.mock.calls[0]![0] as {
      result: { total: number; matches: Array<{ address: string }> }
    }
    expect(reply.result.total).toBe(1)
    expect(reply.result.matches).toEqual([{ sheet: 'Sheet1', address: 'A1', value: 'x' }])
  })

  it('refuses an empty query before reaching the engine', async () => {
    const findCells = vi.fn()
    await sendCommand('search_sheet', { query: '  ' }, { findCells })
    const reply = reportMcpResult.mock.calls[0]![0] as { ok: boolean; error?: string }
    expect(reply.ok).toBe(false)
    expect(reply.error).toContain('non-empty')
    expect(findCells).not.toHaveBeenCalled()
  })

  it('rejects an unknown worksheet name instead of scanning the wrong sheet', async () => {
    const findCells = vi.fn()
    await sendCommand('search_sheet', { query: 'x', sheet: 'Nope' }, { findCells })
    const reply = reportMcpResult.mock.calls[0]![0] as { ok: boolean; error?: string }
    expect(reply.ok).toBe(false)
    expect(reply.error).toContain('no worksheet named "Nope"')
    expect(findCells).not.toHaveBeenCalled()
  })
})
