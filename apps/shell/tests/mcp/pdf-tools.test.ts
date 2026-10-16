/**
 * read_pdf: pdfium text extraction (read-text.ts) + the MCP tool around it.
 * The fixture PDFs are built with pdf-lib (same dependency the pdf editor
 * ships) so the extraction runs against real, minimally-structured files.
 */
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { beforeAll, describe, expect, it } from 'vitest'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { readPdfText } from '../../../pdf/src/main/read-text'
import { createPdfTools, parsePageList, searchPdfFile } from '../../src/main/mcp/tools/pdf-tools'
import { readPdfFile, readPdfInfo } from '../../src/main/daemon/pdf-tools'

let dir: string
let threePages: string
let longFirstPage: string
let manyPages: string
let corrupt: string

async function buildFixturePdf(path: string): Promise<void> {
  const doc = await PDFDocument.create()
  doc.setTitle('MCP Fixture')
  doc.setAuthor('GenOffice Tests')
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const p1 = doc.addPage([400, 300])
  p1.drawText('Hello MCP PDF\nSecond line here', { x: 40, y: 200, size: 14, font })
  doc.addPage([400, 300]) // scanned-page stand-in: no text layer at all
  const p3 = doc.addPage([400, 300])
  p3.drawText('Tail page', { x: 40, y: 200, size: 14, font })
  await writeFile(path, await doc.save())
}

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-pdf-'))
  threePages = join(dir, 'three-pages.pdf')
  await buildFixturePdf(threePages)
  // one long page, so the default preview's per-page clip is exercised
  longFirstPage = join(dir, 'long-first.pdf')
  const long = await PDFDocument.create()
  const longFont = await long.embedFont(StandardFonts.Helvetica)
  long.addPage([400, 300]).drawText('A'.repeat(3000), { x: 10, y: 200, size: 6, font: longFont })
  await writeFile(longFirstPage, await long.save())
  // more pages than one read window, so the 20-page cap is exercised
  manyPages = join(dir, 'many-pages.pdf')
  const many = await PDFDocument.create()
  const manyFont = await many.embedFont(StandardFonts.Helvetica)
  for (let i = 1; i <= 25; i++)
    many.addPage([400, 300]).drawText(`Page ${i}`, { x: 40, y: 200, size: 14, font: manyFont })
  await writeFile(manyPages, await many.save())
  corrupt = join(dir, 'corrupt.pdf')
  await writeFile(corrupt, Buffer.from('this is not a pdf at all'))
})

describe('readPdfText (pdfium extraction)', () => {
  it('extracts page count, metadata and per-page text', async () => {
    const { readFileSync } = await import('node:fs')
    const doc = await readPdfText(new Uint8Array(readFileSync(threePages)))
    expect(doc.pageCount).toBe(3)
    expect(doc.info.title).toBe('MCP Fixture')
    expect(doc.info.author).toBe('GenOffice Tests')
    expect(doc.truncated).toBe(false)
    expect(doc.pages[0]).toMatchObject({ page: 1, widthPt: 400, heightPt: 300 })
    expect(doc.pages[0]!.text).toContain('Hello MCP PDF')
    // pdfium's \r\n line separators come back normalized
    expect(doc.pages[0]!.text).toContain('Second line here')
    expect(doc.pages[0]!.text).not.toContain('\r')
    // a page with no text layer reports itself instead of looking like a bug
    expect(doc.pages[1]).toMatchObject({ chars: 0, text: '', hasTextLayer: false })
    expect(doc.pages[2]!.text).toContain('Tail page')
  })

  it('honors the page span and the char budget', async () => {
    const { readFileSync } = await import('node:fs')
    const bytes = new Uint8Array(readFileSync(threePages))
    const span = await readPdfText(bytes, { fromPage: 3, toPage: 3 })
    expect(span.pages.map((p) => p.page)).toEqual([3])
    const capped = await readPdfText(bytes, { charBudget: 5 })
    expect(capped.pages[0]!.text).toHaveLength(5)
    expect(capped.truncated).toBe(true)
  })

  it('extracts an explicit page list only, in page order', async () => {
    const { readFileSync } = await import('node:fs')
    const bytes = new Uint8Array(readFileSync(threePages))
    const sparse = await readPdfText(bytes, { pages: [3, 1, 3] })
    expect(sparse.pages.map((p) => p.page)).toEqual([1, 3])
    expect(sparse.pages[1]!.text).toContain('Tail page')
  })

  // Regression: a budget spent exactly on a page boundary used to keep pushing
  // empty pages with `truncated:false`, so the result claimed a complete read
  // while pairing `hasTextLayer:true` with empty text.
  it('reports truncation when the budget lands exactly on a page boundary', async () => {
    const { readFileSync } = await import('node:fs')
    const bytes = new Uint8Array(readFileSync(threePages))
    const perPage = (await readPdfText(bytes)).pages[0]!.text.length
    const exact = await readPdfText(bytes, { charBudget: perPage })
    expect(exact.truncated).toBe(true)
    expect(exact.pages[0]!.text).toHaveLength(perPage)
    // no page beyond the budget is reported as an empty-but-readable page
    expect(exact.pages.filter((p) => p.text.length === 0)).toEqual([])
  })
})

describe('readPdfInfo', () => {
  it('reports page count, metadata and the opening pages\' text, as read_pdf previews them', async () => {
    const { readFileSync } = await import('node:fs')
    const out = (await readPdfInfo(new Uint8Array(readFileSync(threePages)), { path: threePages })) as {
      pageCount: number
      info: { title?: string; author?: string }
      pages: Array<{ page: number; text: string; hasTextLayer: boolean }>
    }
    expect(out.pageCount).toBe(3)
    expect(out.info.title).toBe('MCP Fixture')
    expect(out.info.author).toBe('GenOffice Tests')
    // the opening pages, with the same read text a preview carries
    expect(out.pages.map((p) => p.page)).toEqual([1, 2, 3])
    expect(out.pages[0]!.text).toContain('Hello MCP PDF')
    expect(out.pages[0]!.hasTextLayer).toBe(true)
    // the scanned stand-in page reports no text layer
    expect(out.pages[1]).toMatchObject({ chars: 0, hasTextLayer: false })
    // this document is only 3 pages: nothing left to preview
    expect(out).not.toHaveProperty('nextPages')
  })

  it('reports only the opening pages of a long document', async () => {
    const { readFileSync } = await import('node:fs')
    const out = (await readPdfInfo(new Uint8Array(readFileSync(manyPages)), { path: manyPages })) as {
      pageCount: number
      pages: Array<{ page: number; text: string }>
    }
    // the total page count is the whole document, but only the first pages are read
    expect(out.pageCount).toBe(25)
    expect(out.pages.map((p) => p.page)).toEqual([1, 2, 3])
    expect(out.pages[0]!.text).toContain('Page 1')
  })
})

describe('daemon readPdfFile (whole-page paging)', () => {
  it('never splits a page: whole pages up to the cap, the rest deferred', async () => {
    // six pages of 30k chars: past the 80k cap only by whole pages
    const wide = join(dir, 'wide-pages.pdf')
    const doc = await PDFDocument.create()
    const font = await doc.embedFont(StandardFonts.Helvetica)
    for (let i = 0; i < 6; i++) {
      doc.addPage([400, 300]).drawText('A'.repeat(30_000), { x: 5, y: 200, size: 4, font })
    }
    await writeFile(wide, await doc.save())
    const { readFileSync } = await import('node:fs')
    const bytes = new Uint8Array(readFileSync(wide))
    // the full text of one page, to compare against the paged read
    const alone = (await readPdfFile(bytes, { path: wide, pages: '1' })) as {
      pages: Array<{ text: string }>
    }
    const wholeLen = alone.pages[0]!.text.length
    const out = (await readPdfFile(bytes, { path: wide, pages: '1-6' })) as {
      pages: Array<{ page: number; text: string }>
      truncated: boolean
      nextPages?: string
    }
    // two whole pages fit under the cap; page 3 is deferred, not split in
    expect(out.pages.map((p) => p.page)).toEqual([1, 2])
    for (const page of out.pages) expect(page.text).toHaveLength(wholeLen)
    expect(out.truncated).toBe(true)
    expect(out.nextPages).toBe('3-6')
  })
})

describe('parsePageList', () => {
  it('parses single pages, ranges and lists; sorted and unique', () => {
    expect(parsePageList('3', 10)).toEqual([3])
    expect(parsePageList('1-3, 7', 10)).toEqual([1, 2, 3, 7])
    expect(parsePageList('7,1-3,7', 10)).toEqual([1, 2, 3, 7])
  })
  it('rejects junk, inverted ranges and empty specs', () => {
    expect(() => parsePageList('x', 10)).toThrow(/invalid page/)
    expect(() => parsePageList('5-2', 10)).toThrow(/must not precede start/)
    expect(() => parsePageList(' , ', 10)).toThrow(/at least one page/)
  })
  it('drops out-of-range pages instead of refusing them', () => {
    // a fixed window can run past the end of a shorter document
    expect(parsePageList('2-9', 3)).toEqual([2, 3])
    expect(parsePageList('9', 3)).toEqual([])
  })
})

describe('read_pdf tool', () => {
  const handler = createPdfTools()[0]!.handler as (
    args: Record<string, unknown>,
  ) => Promise<unknown>

  it('guards the path like read_docx does', async () => {
    await expect(handler({ path: 'relative.pdf' })).rejects.toThrow(/path must be absolute/)
    await expect(handler({ path: join(dir, 'x.txt') })).rejects.toThrow(/\.pdf/)
    await expect(handler({ path: join(dir, 'nope.pdf') })).rejects.toThrow(/file not found/)
  })

  it('returns a preview (page count + metadata + first pages) by default', async () => {
    const out = (await handler({ path: threePages })) as {
      pageCount: number
      info: { title?: string; author?: string }
      pages: Array<{ page: number; text: string; hasTextLayer: boolean }>
    }
    expect(out.pageCount).toBe(3)
    expect(out.info.title).toBe('MCP Fixture')
    expect(out.info.author).toBe('GenOffice Tests')
    // the opening pages, so the agent learns what the document is
    expect(out.pages.map((p) => p.page)).toEqual([1, 2, 3])
    expect(out.pages[0]!.text).toContain('Hello MCP PDF')
    expect(out.pages[0]!.hasTextLayer).toBe(true)
    // nothing left to preview: this document is only 3 pages
    expect(out).not.toHaveProperty('nextPages')
    expect(out).not.toHaveProperty('note')
  })

  it('clips each preview page with the shared marker', async () => {
    const out = (await handler({ path: longFirstPage })) as {
      pages: Array<{ page: number; chars: number; text: string }>
    }
    // the page holds 3000 chars; the preview keeps 2000 and names the rest
    expect(out.pages[0]!.chars).toBeGreaterThan(2000)
    expect(out.pages[0]!.text).toMatch(/…\(\+\d+ chars\)$/)
    expect(out.pages[0]!.text.slice(0, 2000)).toBe('A'.repeat(2000))
  })

  it('reads only the requested pages', async () => {
    const out = (await handler({ path: threePages, pages: '3' })) as {
      pages: Array<{ page: number; text: string }>
    }
    expect(out.pages.map((p) => p.page)).toEqual([3])
    expect(out.pages[0]!.text).toContain('Tail page')
    const sparse = (await handler({ path: threePages, pages: '1,3' })) as {
      pages: Array<{ page: number }>
    }
    expect(sparse.pages.map((p) => p.page)).toEqual([1, 3])
    // out-of-range pages are dropped, not refused, so a fixed window can run
    // past the end of a shorter document
    const past = (await handler({ path: threePages, pages: '2-9' })) as {
      pages: Array<{ page: number }>
    }
    expect(past.pages.map((p) => p.page)).toEqual([2, 3])
  })

  it('caps one read at 20 pages and names the next span', async () => {
    const out = (await handler({ path: manyPages, pages: '1-25' })) as {
      pageCount: number
      pages: Array<{ page: number }>
      truncated: boolean
      nextPages?: string
    }
    expect(out.pageCount).toBe(25)
    expect(out.pages.map((p) => p.page)).toEqual(Array.from({ length: 20 }, (_, i) => i + 1))
    expect(out.truncated).toBe(true)
    expect(out.nextPages).toBe('21-25')
    // a window inside the cap is a complete read: no truncation, no next span
    const rest = (await handler({ path: manyPages, pages: '21-25' })) as {
      pages: Array<{ page: number }>
      truncated: boolean
      nextPages?: string
    }
    expect(rest.pages.map((p) => p.page)).toEqual([21, 22, 23, 24, 25])
    expect(rest.truncated).toBe(false)
    expect(rest).not.toHaveProperty('nextPages')
  })

  it('names the next span when the preview leaves pages unread', async () => {
    const out = (await handler({ path: manyPages })) as {
      pages: Array<{ page: number }>
      nextPages?: string
    }
    expect(out.pages.map((p) => p.page)).toEqual([1, 2, 3])
    expect(out.nextPages).toBe('4-23')
  })

  it('maps a pdfium load failure to a clean error', async () => {
    await expect(handler({ path: corrupt })).rejects.toThrow(/encrypted or corrupt/)
  })
})

describe('searchPdfFile', () => {
  const bytes = async (path: string): Promise<Uint8Array> => {
    const { readFileSync } = await import('node:fs')
    return new Uint8Array(readFileSync(path))
  }

  it('finds a literal hit and addresses it by 1-based page', async () => {
    const out = await searchPdfFile(await bytes(threePages), { query: 'Hello' })
    expect(out.pageCount).toBe(3)
    expect(out.matches).toHaveLength(1)
    expect(out.matches[0]).toMatchObject({ page: 1 })
    expect(out.matches[0]!.snippet).toContain('Hello')
    expect(out.truncated).toBe(false)
  })

  it('matches case-insensitively by default and exactly with matchCase', async () => {
    const fold = await searchPdfFile(await bytes(threePages), { query: 'hello' })
    expect(fold.matches).toHaveLength(1)
    const exact = await searchPdfFile(await bytes(threePages), { query: 'hello', matchCase: true })
    expect(exact.matches).toHaveLength(0)
  })

  it('scopes the scan to a page span', async () => {
    const wrong = await searchPdfFile(await bytes(threePages), { query: 'Tail', pages: '1' })
    expect(wrong.matches).toHaveLength(0)
    const right = await searchPdfFile(await bytes(threePages), { query: 'Tail', pages: '3' })
    expect(right.matches.map((m) => m.page)).toEqual([3])
  })

  it('caps the hits and reports truncation', async () => {
    const out = await searchPdfFile(await bytes(threePages), { query: 'e', maxResults: 2 })
    expect(out.matches).toHaveLength(2)
    expect(out.truncated).toBe(true)
  })

  it('returns no matches for text that is absent', async () => {
    const out = await searchPdfFile(await bytes(threePages), { query: 'zzzz' })
    expect(out.matches).toEqual([])
    expect(out.truncated).toBe(false)
  })
})

describe('search_pdf tool', () => {
  const handler = createPdfTools()[1]!.handler as (
    args: Record<string, unknown>,
  ) => Promise<unknown>

  it('guards the path and the query', async () => {
    await expect(handler({ path: 'relative.pdf', query: 'x' })).rejects.toThrow(
      /path must be absolute/,
    )
    await expect(handler({ path: join(dir, 'x.txt'), query: 'x' })).rejects.toThrow(/\.pdf/)
    await expect(handler({ path: join(dir, 'nope.pdf'), query: 'x' })).rejects.toThrow(
      /file not found/,
    )
    await expect(handler({ path: threePages, query: '' })).rejects.toThrow(
      /query must not be empty/,
    )
  })

  it('returns hits addressed the way read_pdf addresses pages', async () => {
    const out = (await handler({ path: threePages, query: 'Tail' })) as {
      pageCount: number
      total: number
      matches: Array<{ page: number; snippet: string }>
    }
    expect(out.total).toBe(1)
    expect(out.pageCount).toBe(3)
    expect(out.matches[0]).toMatchObject({ page: 3 })
    expect(out.matches[0]!.snippet).toContain('Tail')
  })

  it('forwards pages, matchCase and maxResults to the scanner', async () => {
    const scoped = (await handler({ path: threePages, query: 'Hello', pages: '3' })) as {
      total: number
    }
    expect(scoped.total).toBe(0)
    const exact = (await handler({ path: threePages, query: 'hello', matchCase: true })) as {
      total: number
    }
    expect(exact.total).toBe(0)
    const capped = (await handler({ path: threePages, query: 'e', maxResults: 1 })) as {
      total: number
      truncated: boolean
    }
    expect(capped.total).toBe(1)
    expect(capped.truncated).toBe(true)
  })
})
