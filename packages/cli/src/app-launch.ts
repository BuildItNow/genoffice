import { spawn } from 'node:child_process'
import { appLaunch } from './resources'
import { CliError, EXIT } from './result'

/**
 * Start the GenOffice app detached: the caller does not wait for it to open and
 * is not its parent for lifecycle purposes (the window outlives the CLI). The
 * app binary is run with `ELECTRON_RUN_AS_NODE` cleared, so it boots the GUI
 * rather than a Node process.
 *
 * `path` is an optional document to open on launch. Shared by `genoffice open`
 * and the MCP daemon's lazy launch, so both bring the app up the same way.
 */
export async function spawnApp(opts: {
  env: NodeJS.ProcessEnv
  path?: string
}): Promise<string> {
  const launch = appLaunch(opts.env)
  if (!launch) {
    throw new CliError(EXIT.app, 'GenOffice app not found', { hint: 'set GENOFFICE_APP_BIN' })
  }
  const env = { ...opts.env }
  delete env.ELECTRON_RUN_AS_NODE
  await new Promise<void>((resolve, reject) => {
    const child = spawn(launch.command, [...launch.args, ...(opts.path ? [opts.path] : [])], {
      detached: true,
      stdio: 'ignore',
      env,
    })
    child.once('error', (err) =>
      reject(new CliError(EXIT.app, `failed to start GenOffice: ${err.message}`)),
    )
    child.once('spawn', () => {
      child.unref()
      resolve()
    })
  })
  return launch.command
}
