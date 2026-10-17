import type { ControlReply } from '@genoffice/cli/control-protocol'
import { createTool, closeTool, listTool, openTool, saveTool } from './document-tools'
import { runSheet, sheetInfoTool } from './sheets-tools'
import { docInfoTool, runDoc } from './docs-tools'
import { pdfInfoTool, pdfReadTool, pdfSearchTool } from './pdf-tools'
import type { DaemonControlDeps } from './shared'

/**
 * The shell's implementation of the daemon's own `office_*` tool channel (the
 * `daemon` control command). These tools reach the running app directly over
 * the control socket rather than through the app's built-in MCP server, so the
 * built-in server and its tool list are untouched.
 *
 * This module is only the dispatcher: it owns the tool-name switch and the
 * `ControlReply` envelope. The handlers live with their category — the family-
 * agnostic tab/file tools in `./document-tools`, the per-family reads and edits
 * in `./sheets-tools`, `./docs-tools` and `./pdf-tools` — over the deps the shell
 * injects (the tab manager, the sheets and docs renderer channels, and the
 * `open_documents` control whose per-family writer/cleanup save and close reuse;
 * see `./shared`).
 *
 * The PDF reads (`pdf_info` / `pdf_read` / `pdf_search`) run right here instead
 * of in a renderer: pdfium is a main-process WASM module, so a PDF is read in
 * the shell from its file (an open tab's, or any `.pdf` path) with the same
 * helpers the built-in `read_pdf` / `search_pdf` use.
 */

export type { DaemonControlDeps } from './shared'

export interface OfficeControl {
  /** run one `office_*` tool (the bare name, e.g. `sheet_read`) against the live app */
  handle: (tool: string, args: Record<string, unknown>) => Promise<ControlReply>
}

const ok = (result: Record<string, unknown>): ControlReply => ({ ok: true, result })
const fail = (message: string): ControlReply => ({
  ok: false,
  error: { reason: 'app_unavailable', message },
})

export function createDaemonControl(getDeps: () => DaemonControlDeps): OfficeControl {
  return {
    handle: async (tool, args) => {
      try {
        // create and open need a tab strip to put the file in, so make sure the
        // shell window is up first — the daemon starts the app (upstream.ts), but
        // on macOS the process outlives the last window it closed. Re-read the
        // deps afterwards: rebuilding the window rebuilds the tab manager too.
        if ((tool === 'create' || tool === 'open') && !getDeps().tabManager) {
          getDeps().revealShellWindow()
        }
        // read the deps per call: the shell rebuilds them (a new tabManager after
        // the last window closes), so a captured object would go stale
        const deps = getDeps()
        switch (tool) {
          case 'list':
            return ok(await listTool(deps))
          case 'create':
            return ok(await createTool(deps, args))
          case 'open':
            return ok(await openTool(deps, args))
          case 'close':
            return ok(await closeTool(deps, args))
          case 'save':
            return ok(await saveTool(deps, args))
          case 'sheet_info':
            return ok(await sheetInfoTool(deps, args))
          case 'sheet_read':
            return ok(await runSheet(deps, args, 'office_sheet_read'))
          case 'sheet_read_formats':
            return ok(await runSheet(deps, args, 'office_sheet_read_formats'))
          case 'sheet_search':
            return ok(await runSheet(deps, args, 'office_sheet_search'))
          case 'sheet_apply_ops':
            return ok(await runSheet(deps, args, 'office_sheet_apply_ops'))
          case 'doc_info':
            return ok(await docInfoTool(deps, args))
          case 'doc_read':
          case 'doc_search':
          case 'doc_apply_ops':
          case 'doc_apply_func':
            // office_doc_apply_func carries the func name in its payload; the
            // renderer runs it through the same editor tool switch
            return ok(await runDoc(deps, args, `office_${tool}`))
          case 'pdf_info':
            return ok(await pdfInfoTool(deps, args))
          case 'pdf_read':
            return ok(await pdfReadTool(deps, args))
          case 'pdf_search':
            return ok(await pdfSearchTool(deps, args))
          default:
            return fail(`unknown office tool "${tool}"`)
        }
      } catch (error) {
        return fail(error instanceof Error ? error.message : String(error))
      }
    },
  }
}
