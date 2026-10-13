import { sep } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { resolveReadPdfPath } from '../src/document-ref'
import type { AppUpstream } from '../src/upstream'

/**
 * `office_gui_read_pdf` exposes `document` (a tab id or path) but its upstream
 * `read_pdf` takes a file path. These tests pin the resolution: id first, then a
 * normalised path, and a bare on-disk `.pdf` path when nothing is open.
 */

const OPEN_PDF = { id: 'tab-1', kind: 'pdf', title: 'report', path: '/docs/Report.pdf' }

function upstream(documents: unknown, opts: { isError?: boolean } = {}): AppUpstream {
  return {
    call: vi.fn(async () => ({
      content: [{ type: 'text', text: JSON.stringify({ documents }) }],
      ...(opts.isError ? { isError: true } : {}),
    })),
  } as unknown as AppUpstream
}

describe('resolveReadPdfPath', () => {
  it('resolves an open PDF by tab id', async () => {
    await expect(resolveReadPdfPath(upstream([OPEN_PDF]), 'tab-1')).resolves.toBe('/docs/Report.pdf')
  })

  it('resolves an open PDF by its path', async () => {
    await expect(resolveReadPdfPath(upstream([OPEN_PDF]), '/docs/Report.pdf')).resolves.toBe(
      '/docs/Report.pdf',
    )
  })

  it('accepts a bare path to a .pdf when nothing is open', async () => {
    await expect(resolveReadPdfPath(upstream([]), '/tmp/loose.pdf')).resolves.toBe('/tmp/loose.pdf')
  })

  it('refuses a document that is not a PDF', async () => {
    const doc = { id: 't', kind: 'docs', title: 'letter.docx', path: '/docs/letter.docx' }
    await expect(resolveReadPdfPath(upstream([doc]), 't')).rejects.toThrow(/is not a PDF/)
  })

  it('refuses a PDF that has never been saved', async () => {
    const doc = { id: 't', kind: 'pdf', title: 'Untitled', path: undefined }
    await expect(resolveReadPdfPath(upstream([doc]), 't')).rejects.toThrow(/never been saved/)
  })

  it('refuses an unknown reference that is not a path', async () => {
    await expect(resolveReadPdfPath(upstream([OPEN_PDF]), 'nope')).rejects.toThrow(
      /no open document matches/,
    )
  })

  it('refuses a non-pdf path that is not open', async () => {
    await expect(resolveReadPdfPath(upstream([]), `/tmp${sep}notes.txt`)).rejects.toThrow(
      /no open document matches/,
    )
  })

  it('refuses an empty document', async () => {
    await expect(resolveReadPdfPath(upstream([]), '')).rejects.toThrow(/pass the PDF/)
  })

  it('surfaces a failure to list the open documents', async () => {
    const broken = {
      call: vi.fn(async () => ({
        content: [
          { type: 'text', text: 'GenOffice is not running, so its open documents are unreachable' },
        ],
        isError: true,
      })),
    } as unknown as AppUpstream
    await expect(resolveReadPdfPath(broken, 'tab-1')).rejects.toThrow(/not running/)
  })
})
