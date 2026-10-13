import { realpathSync } from 'node:fs'
import { extname, isAbsolute, resolve } from 'node:path'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import type { AppUpstream } from './upstream'

/**
 * Resolving the `document` argument of `office_gui_read_pdf`.
 *
 * Every other visible-document read (read_document / read_sheet / read_deck)
 * resolves its `document` to a live webContents inside the app. A PDF has no
 * live editor state — the pdf app is a viewer — so this read resolves the same
 * `document` reference to the file instead and reads that. The reference is a
 * tab id or a path, matched exactly like the app's own open-document lookup:
 * exact tab id first, then a normalised path.
 */

interface OpenDocumentRow {
  id?: unknown
  kind?: unknown
  title?: unknown
  path?: unknown
}

export async function resolveReadPdfPath(
  upstream: AppUpstream,
  document: unknown,
): Promise<string> {
  const ref = typeof document === 'string' ? document.trim() : ''
  if (!ref) {
    throw new Error('pass the PDF to read as `document` — an open tab id or a file path')
  }

  const match = matchOpen(await openDocuments(upstream), ref)
  if (match) {
    if (match.kind !== 'pdf') {
      throw new Error(
        `"${label(match, ref)}" is not a PDF — read it with read_document, read_sheet or read_deck`,
      )
    }
    if (typeof match.path !== 'string' || !match.path) {
      throw new Error(`"${label(match, ref)}" has never been saved, so there is no file to read`)
    }
    return match.path
  }

  // not an open document: accept a bare path to a PDF on disk
  if (isAbsolute(ref) && extname(ref).toLowerCase() === '.pdf') return ref
  throw new Error(
    `no open document matches "${ref}"; open a PDF in GenOffice, or pass the absolute path of a .pdf`,
  )
}

function matchOpen(rows: readonly OpenDocumentRow[], ref: string): OpenDocumentRow | undefined {
  const byId = rows.find((row) => row.id === ref)
  if (byId) return byId
  const wanted = normalize(ref)
  return rows.find((row) => typeof row.path === 'string' && normalize(row.path) === wanted)
}

async function openDocuments(upstream: AppUpstream): Promise<OpenDocumentRow[]> {
  const result = await upstream.call('open_documents', { action: 'list' })
  if (result.isError) {
    throw new Error(textOf(result) || 'could not list the open GenOffice documents')
  }
  let parsed: { documents?: unknown }
  try {
    parsed = JSON.parse(textOf(result)) as { documents?: unknown }
  } catch {
    throw new Error('GenOffice returned a malformed open-document list')
  }
  return Array.isArray(parsed?.documents) ? (parsed.documents as OpenDocumentRow[]) : []
}

function textOf(result: CallToolResult): string {
  const first = Array.isArray(result.content) ? result.content[0] : undefined
  return first && first.type === 'text' ? first.text : ''
}

function label(row: OpenDocumentRow, fallback: string): string {
  return typeof row.title === 'string' && row.title ? row.title : fallback
}

/** resolved absolute path, realpath'd when it exists, case-folded on Windows */
function normalize(value: string): string {
  let resolved = resolve(value)
  try {
    resolved = realpathSync.native(resolved)
  } catch {
    // unsaved or moved file: compare the resolved path as-is
  }
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}
