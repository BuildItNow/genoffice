import { spawnApp } from '@genoffice/cli/app-launch'
import { controlEndpoint, controlRequest, waitForControlEndpoint } from '@genoffice/cli/control'
import { DAEMON_CMD, type ControlEndpoint } from '@genoffice/cli/control-protocol'

/**
 * The connection from the daemon to the GenOffice app that backs the `office_*`
 * family.
 *
 * The `office_*` tools ride the app's control socket directly (`callOffice`), so
 * the app's built-in MCP server and its tool list are never involved. The
 * control endpoint only exists while the app runs, so the daemon discovers it
 * (`controlEndpoint`) and, when nothing answers, starts the app itself and waits
 * for it to publish its endpoint.
 *
 * Starting the app does NOT switch its built-in MCP server on: the daemon only
 * needs the control socket, so a locked-down install never opens the app's
 * editing surface just because a tool was called.
 */

/** Thrown when the app cannot be reached. */
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
  constructor(private readonly opts: AppUpstreamOptions) {}

  /**
   * The running app's control-socket endpoint, starting the app if nothing
   * answers.
   */
  private async ensureEndpoint(): Promise<ControlEndpoint> {
    let endpoint = controlEndpoint(this.opts.env)
    if (!endpoint) {
      this.opts.log('[daemon] no GenOffice app is running; starting it for this call')
      await spawnApp({ env: this.opts.env })
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
    return endpoint
  }

  /**
   * Call one of the daemon's own `office_*` tools. This rides the control socket
   * straight into the app's daemon handler — the app's built-in MCP server is
   * never involved, so its tool list is unaffected. A `{ok:false}` reply becomes
   * a thrown `AppUnavailableError`; the result is the tool's payload.
   */
  async callOffice(tool: string, args: Record<string, unknown>): Promise<Record<string, unknown>> {
    try {
      const endpoint = await this.ensureEndpoint()
      return await controlRequest(endpoint, { cmd: DAEMON_CMD, tool, args })
    } catch (error) {
      throw new AppUnavailableError(
        `GenOffice did not answer the "${tool}" call: ${
          error instanceof Error ? error.message : String(error)
        }`,
      )
    }
  }
}
