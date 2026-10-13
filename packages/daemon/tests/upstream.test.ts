import { beforeEach, describe, expect, it, vi } from 'vitest'
import { AppUnavailableError, AppUpstream } from '../src/upstream'

/**
 * The upstream connection is exercised with the app's control socket, the app
 * launcher, and the SDK client/transport all stubbed, so the tests pin down the
 * daemon's own behaviour: how it discovers the app's MCP port, when it starts
 * the app, how it reports failures, and how it retries a dead connection.
 */

const h = vi.hoisted(() => ({
  endpoint: null as { pid: number; endpoint: string; token: string } | null,
  endpointAfterWait: null as { pid: number; endpoint: string; token: string } | null,
  controlReply: null as Record<string, unknown> | null,
  controlThrows: null as Error | null,
  waitCalls: [] as number[],
  spawnCalls: [] as Array<{ env: NodeJS.ProcessEnv; path?: string }>,
  spawnThrows: null as Error | null,
  connectThrows: null as Error | null,
  callToolImpl: null as ((params: unknown) => unknown) | null,
  callToolCalls: [] as Array<{ name: string }>,
  clients: [] as any[],
  closed: [] as any[],
  transportUrls: [] as string[],
}))

vi.mock('@genoffice/cli/control', () => ({
  controlEndpoint: vi.fn(() => h.endpoint),
  controlRequest: vi.fn(async () => {
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

vi.mock('@modelcontextprotocol/sdk/client/streamableHttp.js', () => ({
  StreamableHTTPClientTransport: vi.fn(function (this: any, url: URL) {
    h.transportUrls.push(url.toString())
    this.url = url
  }),
}))

vi.mock('@modelcontextprotocol/sdk/client/index.js', () => ({
  Client: vi.fn(function (this: any) {
    this.connect = vi.fn(async () => {
      if (h.connectThrows) throw h.connectThrows
    })
    this.callTool = vi.fn(async (params: any) => {
      h.callToolCalls.push(params)
      if (h.callToolImpl) return h.callToolImpl(params)
      return { content: [{ type: 'text', text: 'ok' }] }
    })
    this.close = vi.fn(async () => {
      h.closed.push(this)
    })
    h.clients.push(this)
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
  h.connectThrows = null
  h.callToolImpl = null
  h.callToolCalls = []
  h.clients = []
  h.closed = []
  h.transportUrls = []
})

describe('AppUpstream.call', () => {
  it('forwards the bare upstream name to the app and returns its result', async () => {
    h.endpoint = live
    h.controlReply = { running: true, port: 3093 }
    const result = await upstream().call('open_in_genoffice', { path: '/a.docx' })
    expect(h.callToolCalls[0]).toMatchObject({ name: 'open_in_genoffice', arguments: { path: '/a.docx' } })
    expect(result.content[0]).toMatchObject({ text: 'ok' })
    expect(h.transportUrls[0]).toBe('http://127.0.0.1:3093/mcp')
  })

  it('starts the app when none is running, forcing its MCP server on', async () => {
    h.endpointAfterWait = live
    h.controlReply = { running: true, port: 3093 }
    await upstream().call('open_documents', { action: 'list' })
    expect(h.spawnCalls).toHaveLength(1)
    expect(h.spawnCalls[0].env.GENOFFICE_MCP_FORCE).toBe('1')
    expect(h.waitCalls).toEqual([1000])
    expect(h.callToolCalls[0].name).toBe('open_documents')
  })

  it('fails when the app it launched never comes up', async () => {
    h.endpointAfterWait = null
    await expect(upstream().call('open_documents', {})).rejects.toBeInstanceOf(AppUnavailableError)
    expect(h.spawnCalls).toHaveLength(1)
  })

  it('fails when the app is up but its MCP server is off', async () => {
    h.endpoint = live
    h.controlReply = { running: false, port: 0 }
    await expect(upstream().call('open_documents', {})).rejects.toThrow(/Settings ▸ Integrations/)
  })

  it('fails when the control socket cannot report the MCP endpoint', async () => {
    h.endpoint = live
    h.controlThrows = new Error('socket closed')
    await expect(upstream().call('open_documents', {})).rejects.toThrow(/did not report its MCP endpoint/)
  })

  it('fails when the transport refuses the connection', async () => {
    h.endpoint = live
    h.controlReply = { running: true, port: 3093 }
    h.connectThrows = new Error('ECONNREFUSED')
    await expect(upstream().call('open_documents', {})).rejects.toThrow(/could not connect/)
  })

  it('retries once against a fresh connection when a call fails', async () => {
    h.endpoint = live
    h.controlReply = { running: true, port: 3093 }
    let first = true
    h.callToolImpl = () => {
      if (first) {
        first = false
        throw new Error('transport gone')
      }
      return { content: [{ type: 'text', text: 'recovered' }] }
    }
    const result = await upstream().call('read_document', {})
    expect(h.clients).toHaveLength(2)
    expect(h.closed).toHaveLength(1)
    expect(result.content[0]).toMatchObject({ text: 'recovered' })
  })

  it('closes the client on dispose', async () => {
    h.endpoint = live
    h.controlReply = { running: true, port: 3093 }
    const app = upstream()
    await app.call('read_document', {})
    app.dispose()
    expect(h.closed).toHaveLength(1)
  })
})
