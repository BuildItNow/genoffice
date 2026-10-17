import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import { OFFICE_PREFIX } from '../names'
import type { AppUpstream } from '../upstream'
import { OFFICE_TOOLS } from './tools'

/**
 * Register the daemon's own `office_*` family on the MCP server. Each tool is a
 * thin passthrough: the daemon forwards its arguments to the app over the
 * control socket (`AppUpstream.callOffice`) and relays the payload as JSON text.
 * A tool with a `local` handler is answered right here instead (static metadata
 * that needs no app). The app's built-in MCP server is never involved, so its
 * tool list is unaffected.
 */
export function registerOfficeTools(server: McpServer, upstream: AppUpstream): void {
  for (const tool of OFFICE_TOOLS) {
    const exposed = `${OFFICE_PREFIX}${tool.name}`
    server.registerTool(
      exposed,
      {
        description: tool.description,
        inputSchema: tool.shape,
        annotations: { readOnlyHint: tool.readOnly === true, openWorldHint: false },
      },
      async (args): Promise<CallToolResult> => {
        try {
          const call = args as Record<string, unknown>
          const prepared = tool.prepare ? await tool.prepare(call) : call
          const result = tool.local
            ? await tool.local(prepared)
            : await upstream.callOffice(tool.name, prepared)
          return { content: [{ type: 'text', text: JSON.stringify(result) }] }
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error)
          return {
            content: [{ type: 'text', text: `${exposed} failed: ${message}` }],
            isError: true,
          }
        }
      },
    )
  }
}
