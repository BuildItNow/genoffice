import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import { createContext, disposeContext, registerCliTools } from '@genoffice/cli/mcp-server'
import { GUI_TOOLS } from './gui-tools'
import { CLI_PREFIX } from './names'
import { AppUnavailableError, AppUpstream } from './upstream'
import { DAEMON_VERSION } from './version'

/**
 * The daemon's single MCP server, exposing two families:
 *
 * - `office_cli_*` — the bundled genoffice CLI's headless tools, served right
 *   here. They never need the app, so they keep working after the window closes.
 * - `office_gui_*` — the tools that drive a live GenOffice window. Each is a
 *   proxy to the app's own built-in MCP server, which the daemon starts on
 *   demand (see upstream.ts).
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
    `office_gui_* drive a live GenOffice window: the visible editing sessions (create_session / save_session plus the family content tools), the documents the user has open (office_gui_opened_documents), and opening a file in a tab (office_gui_open). The daemon starts GenOffice on the first such call if it is not already running, and hands the call to the app's own MCP server.`,
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
  registerGuiTools(server, upstream)

  return {
    server,
    upstream,
    dispose: () => {
      upstream.dispose()
      disposeContext(ctx)
    },
  }
}

function registerGuiTools(server: McpServer, upstream: AppUpstream): void {
  for (const tool of GUI_TOOLS) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: tool.shape,
        annotations: { readOnlyHint: tool.readOnly === true, openWorldHint: false },
      },
      async (args): Promise<CallToolResult> => {
        try {
          const call = args as Record<string, unknown>
          // most tools forward their arguments verbatim; read_pdf resolves the
          // `document` it exposes into the file path its upstream takes
          const forwarded = tool.adapt ? await tool.adapt(call, upstream) : call
          const result = await upstream.call(tool.upstream, forwarded)
          return relayResult(result)
        } catch (error) {
          return errorResult(error, tool.name)
        }
      },
    )
  }
}

/**
 * Relay the app's result to our client: the text content and isError, exactly
 * like the CLI family. Additionally forward `structuredContent` when the app
 * sent one — the daemon declares no `outputSchema`, so the SDK passes it through
 * untouched. (The app serialises its payloads as JSON text and sends none today;
 * this keeps the field if that ever changes.)
 */
export function relayResult(result: CallToolResult): CallToolResult {
  const relayed: CallToolResult = { content: result.content }
  if (result.structuredContent !== undefined) relayed.structuredContent = result.structuredContent
  if (result.isError) relayed.isError = true
  return relayed
}

function errorResult(error: unknown, toolName: string): CallToolResult {
  const message =
    error instanceof AppUnavailableError
      ? `${error.message} — this is needed by ${toolName}`
      : error instanceof Error
        ? error.message
        : String(error)
  return { content: [{ type: 'text', text: message }], isError: true }
}
