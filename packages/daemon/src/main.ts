import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { createDaemonServer } from './server'

/**
 * `office-daemon` entry point. Serves the two tool families over stdio, the
 * transport MCP clients (Claude Code, Cursor, …) launch and speak to. The
 * daemon starts no window and no app of its own: the office_gui_* tools bring
 * GenOffice up only when one is actually called.
 */

const HELP = `office-daemon — GenOffice MCP server that outlives the app

Usage:
  office-daemon [--cli] [--compact-schemas]

Serves the GenOffice MCP tool surface over stdio. office_gui_* (drive a live
window; the app is started on demand) is always on. office_cli_* (the headless
family, tens of tools) is off unless --cli is given; it also takes
GENOFFICE_MCP_CLI=1. The cloud commands (search, image, media, capabilities) are
never served. --compact-schemas advertises the CLI tools' inline JSON as plain
arrays (smaller tools/list; also GENOFFICE_MCP_COMPACT_SCHEMAS=1) and has no
effect without --cli.
`

function truthy(value: string | undefined): boolean {
  return /^(1|true|yes|on)$/i.test(value?.trim() ?? '')
}

async function main(argv: string[]): Promise<number> {
  const args = argv.slice(2)
  let compactSchemas = truthy(process.env.GENOFFICE_MCP_COMPACT_SCHEMAS)
  let cli = truthy(process.env.GENOFFICE_MCP_CLI)
  for (const arg of args) {
    switch (arg) {
      case '-h':
      case '--help':
        process.stdout.write(HELP)
        return 0
      case '--cli':
        cli = true
        break
      case '--compact-schemas':
        compactSchemas = true
        break
      default:
        process.stderr.write(`office-daemon: unknown argument "${arg}" (try --help)\n`)
        return 2
    }
  }

  const { server, dispose } = await createDaemonServer({
    cwd: process.cwd(),
    env: process.env,
    // the MCP transport owns stdout; diagnostics go to stderr
    log: (message) => process.stderr.write(`${message}\n`),
    cli,
    compactSchemas,
  })

  const transport = new StdioServerTransport()
  await new Promise<void>((resolve, reject) => {
    server.server.onclose = () => resolve()
    server.connect(transport).catch(reject)
  })
  dispose()
  return 0
}

main(process.argv).then(
  (code) => process.exit(code),
  (error: unknown) => {
    process.stderr.write(`office-daemon: ${error instanceof Error ? error.stack : String(error)}\n`)
    process.exit(1)
  },
)
