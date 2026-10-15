import { beforeEach, describe, expect, it, vi } from 'vitest'
import { AppUnavailableError, AppUpstream } from '../src/upstream'

/**
 * The daemon's connection to the app is exercised with the app's control socket
 * and the launcher stubbed, so the tests pin down the daemon's own behaviour:
 * how it reaches the app over the control socket, when it starts the app, and
 * how it reports failures.
 */

const h = vi.hoisted(() => ({
  endpoint: null as { pid: number; endpoint: string; token: string } | null,
  endpointAfterWait: null as { pid: number; endpoint: string; token: string } | null,
  controlReply: null as Record<string, unknown> | null,
  controlThrows: null as Error | null,
  waitCalls: [] as number[],
  spawnCalls: [] as Array<{ env: NodeJS.ProcessEnv; path?: string }>,
  spawnThrows: null as Error | null,
  requests: [] as Array<{ cmd: string; tool: string; args: Record<string, unknown> }>,
}))

vi.mock('@genoffice/cli/control', () => ({
  controlEndpoint: vi.fn(() => h.endpoint),
  controlRequest: vi.fn(async (_endpoint: unknown, req: unknown) => {
    h.requests.push(req as { cmd: string; tool: string; args: Record<string, unknown> })
    if (h.controlThrows) throw h.controlThrows
    return h.controlReply
  }),
  waitForControlEndpoint: vi.fn(async (_env: unknown, ms: number) => {
    h.waitCalls.push(ms)
    return h.endpointAfterWait
  }),
}))

vi.mock('@genoffice/cli/app-launch', () => ({
  spawnApp: vi.fn(async (opts: { env: NodeJS.ProcessEnv; path?: string }) => {
    h.spawnCalls.push(opts)
    if (h.spawnThrows) throw h.spawnThrows
    return 'genoffice'
  }),
}))

const live = { pid: 4242, endpoint: '/tmp/control.sock', token: 'secret' }
const env = { GENOFFICE_USER_DATA: '/tmp/genoffice-test' }

function upstream() {
  return new AppUpstream({ env, log: () => {}, launchTimeoutMs: 1000 })
}

beforeEach(() => {
  h.endpoint = null
  h.endpointAfterWait = null
  h.controlReply = null
  h.controlThrows = null
  h.waitCalls = []
  h.spawnCalls = []
  h.spawnThrows = null
  h.requests = []
})

describe('AppUpstream.callOffice', () => {
  it('sends the bare tool name and its args over the control socket', async () => {
    h.endpoint = live
    h.controlReply = { tabs: [] }
    const result = await upstream().callOffice('list', {})
    expect(h.requests[0]).toMatchObject({ tool: 'list', args: {} })
    expect(result).toEqual({ tabs: [] })
    // an app that is already up is used as-is, with no launch
    expect(h.spawnCalls).toHaveLength(0)
  })

  it('starts the app when none is running, without forcing its MCP server on', async () => {
    h.endpointAfterWait = live
    h.controlReply = { tabs: [] }
    await upstream().callOffice('list', {})
    expect(h.spawnCalls).toHaveLength(1)
    // the app is brought up for the control socket only; its built-in MCP server
    // stays whatever the Settings switch says
    expect(h.spawnCalls[0]!.env.GENOFFICE_MCP_FORCE).toBeUndefined()
    expect(h.waitCalls).toEqual([1000])
    expect(h.requests[0]).toMatchObject({ tool: 'list' })
  })

  it('fails when the app it launched never comes up', async () => {
    h.endpointAfterWait = null
    await expect(upstream().callOffice('list', {})).rejects.toBeInstanceOf(AppUnavailableError)
    expect(h.spawnCalls).toHaveLength(1)
  })

  it('names the tool when the control socket rejects the call', async () => {
    h.endpoint = live
    h.controlThrows = new Error('socket closed')
    await expect(upstream().callOffice('sheet_read', {})).rejects.toThrow(/sheet_read/)
  })
})
