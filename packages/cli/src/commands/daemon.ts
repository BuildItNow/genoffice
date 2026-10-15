import { spawn } from 'node:child_process'
import { flagBool } from '../args'
import type { CommandDef } from '../registry'
import { daemonBundlePath } from '../resources'
import { CliError, EXIT } from '../result'

/**
 * How an MCP client reaches the daemon: the `genoffice` command a terminal
 * already has, with `daemon` as its subcommand. The daemon lives in its own
 * bundle (packages/daemon) because it outlives the app — it keeps serving while
 * GenOffice is closed and starts it on demand — but it has no reason to be a
 * second command on the PATH: this reuses the same link the CLI installs and
 * hands stdio straight to the bundle, so the client speaks to it directly.
 */
export const daemonCommand: CommandDef = {
  name: 'daemon',
  summary:
    'Serve the GenOffice MCP server on stdio through the genoffice command already on the PATH: office_* drives the live app (started on demand) and keeps running across its restarts, and --cli adds the headless office_cli_* family. This is what an MCP client config points at — `command: "genoffice", args: ["daemon"]`.',
  usage: 'daemon [--cli] [--compact-schemas]',
  // the MCP transport owns stdout for the whole run; runCli prints nothing after it
  quiet: true,
  options: [
    {
      name: 'cli',
      description:
        'also serve the headless office_cli_* family (tens of tools); off by default, also GENOFFICE_MCP_CLI=1',
    },
    {
      name: 'compact-schemas',
      description:
        "advertise the CLI tools' inline JSON as plain arrays (smaller tools/list; also GENOFFICE_MCP_COMPACT_SCHEMAS=1)",
    },
  ],
  async run(args, ctx) {
    if (args.positionals.length) {
      throw new CliError(
        EXIT.usage,
        `daemon takes no arguments, got "${args.positionals[0]}"`,
        undefined,
        {
          reason: 'invalid_argument',
          suggestion: 'run `genoffice daemon` alone (or with --cli / --compact-schemas)',
        },
      )
    }
    const bundle = daemonBundlePath()
    if (!bundle) {
      throw new CliError(EXIT.app, 'cannot locate the office-daemon bundle', undefined, {
        reason: 'app_unavailable',
        suggestion: 'in a checkout, build it first: npm run build -w @genoffice/daemon',
      })
    }
    // the bundle parses these itself; the daemon only ever ran on a Node, so
    // process.execPath (the app binary under ELECTRON_RUN_AS_NODE) runs it too
    const forwarded: string[] = []
    if (flagBool(args, 'cli')) forwarded.push('--cli')
    if (flagBool(args, 'compact-schemas')) forwarded.push('--compact-schemas')
    let code: number | null
    try {
      code = await new Promise<number | null>((resolve, reject) => {
        const child = spawn(process.execPath, [bundle, ...forwarded], {
          stdio: 'inherit',
          env: ctx.env,
        })
        child.once('error', reject)
        child.once('exit', (exitCode, signal) => resolve(signal ? null : exitCode))
      })
    } catch (err) {
      throw new CliError(
        EXIT.app,
        `failed to start the daemon: ${err instanceof Error ? err.message : String(err)}`,
      )
    }
    if (code !== null && code !== 0) throw new CliError(EXIT.app, `daemon exited with code ${code}`)
    return { summary: 'daemon session ended' }
  },
}
