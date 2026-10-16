import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { createContext, disposeContext, registerCliTools } from '@genoffice/cli/mcp-server'
import { registerOfficeTools } from './office/register'
import { CLI_PREFIX } from './names'
import { AppUpstream } from './upstream'
import { DAEMON_VERSION } from './version'

/**
 * The daemon's single MCP server, exposing two families:
 *
 * - `office_cli_*` — the bundled genoffice CLI's headless tools, served right
 *   here. They never need the app, so they keep working after the window closes.
 * - `office_*` — the daemon's own live-window family (create/list/save, the
 *   sheet tools and the PDF reads). It reaches the app over the control socket,
 *   so the app's built-in MCP server and its tool list are never involved.
 */

export interface DaemonServerOptions {
  cwd: string
  env: NodeJS.ProcessEnv
  log: (message: string) => void
  /** advertise the CLI tools' inline JSON as plain arrays (smaller tools/list) */
  compactSchemas?: boolean
  /**
   * Register the headless `office_cli_*` family. Off by default: it is the much
   * larger family (tens of tools), and a GUI session does not need it, so a host
   * opts in (`--cli` / GENOFFICE_MCP_CLI=1). The cloud commands are never served
   * either way (see CLOUD_TOOLS).
   */
  cli?: boolean
}

export interface DaemonServer {
  server: McpServer
  upstream: AppUpstream
  dispose: () => void
}

/**
 * The CLI tools that reach the cloud: search, image and media (their work leaves
 * the machine) and capabilities, the probe that reports whether those can run.
 * The daemon is a local Office domain, so they are dropped outright — not gated.
 */
const CLOUD_TOOLS = ['capabilities', 'search', 'image', 'media'] as const

/** What the client shows the model before any tool is called; the CLI line follows `cli`. */
function instructions(cli: boolean): string {
  const lines: string[] = []
  if (cli) {
    lines.push(
      `office_cli_* run headless on this machine with the bundled genoffice CLI (create, read, convert, edit and render docx/xlsx/pptx/pdf/csv). They need no window, keep working whether or not GenOffice is open, and never send data off the machine.`,
    )
  }
  lines.push(
    `office_* are the daemon's own tools for a live window: office_list (the open tabs), office_create / office_open / office_close / office_save, the per-tab sheet tools (office_sheet_info / office_sheet_read / office_sheet_read_formats / office_sheet_search / office_sheet_edit), and the PDF tools (office_pdf_info / office_pdf_read / office_pdf_search, addressed by an open tab or an absolute .pdf path). office_sheet_edit_help returns the op reference for office_sheet_edit and needs no app. They reach the app directly over its local control socket, so they never touch the app's built-in MCP server.`,
  )
  return [
    'GenOffice MCP daemon: the GenOffice tools behind one server, so editing survives closing the app.',
    ...lines,
  ].join('\n')
}

export async function createDaemonServer(opts: DaemonServerOptions): Promise<DaemonServer> {
  const cli = opts.cli === true
  const ctx = createContext({ cwd: opts.cwd, env: opts.env, log: opts.log })
  const upstream = new AppUpstream({ env: opts.env, log: opts.log })
  const server = new McpServer(
    { name: 'genoffice-daemon', version: DAEMON_VERSION },
    { instructions: instructions(cli) },
  )

  if (cli) {
    await registerCliTools(server, ctx, {
      namePrefix: CLI_PREFIX,
      compactSchemas: opts.compactSchemas,
      // the cloud commands are dropped even when the CLI family is on
      omit: CLOUD_TOOLS,
    })
  }
  registerOfficeTools(server, upstream)

  return {
    server,
    upstream,
    dispose: () => disposeContext(ctx),
  }
}
