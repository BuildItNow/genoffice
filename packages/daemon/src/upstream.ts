import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import { spawnApp } from '@genoffice/cli/app-launch'
import { controlEndpoint, controlRequest, waitForControlEndpoint } from '@genoffice/cli/control'
import { DAEMON_VERSION } from './version'

/**
 * The connection from the daemon to the GenOffice app's own built-in MCP server,
 * used to serve the `office_gui_*` tools.
 *
 * The app's server only exists while the app runs, and its port is whatever the
 * user configured. The daemon therefore discovers it through the control socket
 * (a `mcp` command answered by the running app) and, when nothing answers,
 * starts the app itself and waits for it to publish its endpoint. The MCP client
 * session is opened lazily and dropped the moment a call fails, so a later call
 * re-discovers a relaunched app instead of talking to a dead port.
 */

/** Thrown when the app (or its built-in MCP server) cannot be reached. */
export class AppUnavailableError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'AppUnavailableError'
  }
}

export interface AppUpstreamOptions {
  env: NodeJS.ProcessEnv
  log: (message: string) => void
  /** how long to wait for the app's control endpoint after launching it */
  launchTimeoutMs?: number
}

const DEFAULT_LAUNCH_TIMEOUT_MS = 45_000

export class AppUpstream {
  private client: Client | null = null
  private inflight: Promise<Client> | null = null
  /** set once the daemon has started the app, so a later call does not treat it as never-launched */
  private launched = false

  constructor(private readonly opts: AppUpstreamOptions) {}

  /**
   * A live client to the app's MCP server, starting the app if it is not
   * running. Concurrent callers share one connection attempt.
   */
  async ensure(): Promise<Client> {
    if (this.client) return this.client
    if (this.inflight) return this.inflight
    this.inflight = this.connectOnce().finally(() => {
      this.inflight = null
    })
    return this.inflight
  }

  private async connectOnce(): Promise<Client> {
    const port = await this.discoverPort()
    const transport = new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`))
    const client = new Client(
      { name: 'genoffice-daemon', version: DAEMON_VERSION },
      { capabilities: {} },
    )
    try {
      await client.connect(transport)
    } catch (error) {
      throw new AppUnavailableError(
        `could not connect to GenOffice's MCP server on port ${port}: ${
          error instanceof Error ? error.message : String(error)
        }`,
      )
    }
    this.opts.log(`[daemon] connected to GenOffice MCP on port ${port}`)
    this.client = client
    return client
  }

  /** Where the app's built-in MCP server listens, starting the app if needed. */
  private async discoverPort(): Promise<number> {
    let endpoint = controlEndpoint(this.opts.env)
    if (!endpoint) {
      this.opts.log('[daemon] no GenOffice app is running; starting it for this GUI call')
      await spawnApp({ env: { ...this.opts.env, GENOFFICE_MCP_FORCE: '1' } })
      this.launched = true
      endpoint = await waitForControlEndpoint(
        this.opts.env,
        this.opts.launchTimeoutMs ?? DEFAULT_LAUNCH_TIMEOUT_MS,
      )
      if (!endpoint) {
        throw new AppUnavailableError(
          'GenOffice was started but did not come up in time; retry once its window is ready',
        )
      }
    }
    let reply: Record<string, unknown>
    try {
      reply = await controlRequest(endpoint, { cmd: 'mcp' })
    } catch (error) {
      throw new AppUnavailableError(
        `GenOffice is running but did not report its MCP endpoint: ${
          error instanceof Error ? error.message : String(error)
        }`,
      )
    }
    const running = reply.running === true
    const port = Number(reply.port)
    if (!running || !Number.isInteger(port) || port <= 0) {
      throw new AppUnavailableError(
        'the built-in MCP server is off in GenOffice; enable it under Settings ▸ Integrations',
      )
    }
    return port
  }

  /**
   * Call one of the app's tools. A tool-level failure comes back as an isError
   * result; only a transport failure throws, and that is retried once against a
   * fresh connection (the app may have quit between calls and been relaunched).
   */
  async call(name: string, args: Record<string, unknown>): Promise<CallToolResult> {
    const client = await this.ensure()
    try {
      return (await client.callTool({ name, arguments: args })) as CallToolResult
    } catch (error) {
      this.reset()
      this.opts.log(
        `[daemon] upstream "${name}" failed (${error instanceof Error ? error.message : String(error)}); reconnecting`,
      )
      const retried = await this.ensure()
      return (await retried.callTool({ name, arguments: args })) as CallToolResult
    }
  }

  private reset(): void {
    const client = this.client
    this.client = null
    void client?.close().catch(() => undefined)
  }

  dispose(): void {
    this.reset()
  }
}
