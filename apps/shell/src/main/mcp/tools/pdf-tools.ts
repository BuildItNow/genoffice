import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, extname, isAbsolute } from 'node:path'
import { z } from 'zod'
import { clipText } from '@genoffice/pptx-ops'
import { readPdfText } from '../../../../../pdf/src/main/read-text'
import type { McpToolDefinition } from '../mcp-server'

/**
 * PDF reading — the registry's `pdf` family exposes read-only MCP access
 * (`mcp.read`, no session, no generation): the pdf app is a viewer, so an
 * agent reads a PDF's text layer directly from disk through the same
 * main-process pdfium the editor uses. Like read_docx, this is headless and
 * always registered. Read-only is the final design — MCP does not drive the
 * pdf editor.
 */

const PDF_EXT = '.pdf'
/** hard cap on returned text so one huge PDF cannot flood the agent's context */
const MAX_TOTAL_CHARS = 80_000
/** default preview: the opening pages, so an agent learns what the document is */
const PREVIEW_PAGES = 3
/** per-page clip in the preview; the inline marker says how much is withheld */
const PREVIEW_CHARS_PER_PAGE = 2000
/** one call reads at most this many pages; the result names the next span to continue from */
const MAX_PAGES_PER_READ = 20

/**
 * Parse a 1-based page list like "3" or "1-5,8" into sorted unique pages.
 * Malformed tokens and inverted ranges throw; pages outside 1..pageCount are
 * dropped rather than refused, so a fixed page window (e.g. "21-40") can run
 * past the end of a shorter document and still read its tail.
 */
export function parsePageList(spec: string, pageCount: number): number[] {
  const pages = new Set<number>()
  let sawToken = false
  for (const part of spec.split(',')) {
    const trimmed = part.trim()
    if (!trimmed) continue
    sawToken = true
    const range = trimmed.match(/^(\d+)(?:\s*-\s*(\d+))?$/)
    if (!range) throw new Error(`invalid page "${trimmed}" — use "3" or "1-5,8"`)
    const from = Number(range[1])
    const to = range[2] === undefined ? from : Number(range[2])
    if (to < from) throw new Error(`invalid page range "${trimmed}" — end must not precede start`)
    for (let p = Math.max(from, 1); p <= Math.min(to, pageCount); p++) pages.add(p)
  }
  if (!sawToken) throw new Error('pages must list at least one page, e.g. "1-5,8"')
  return [...pages].sort((a, b) => a - b)
}

/** the next 20-page window after `last`, so a reader knows how to continue */
function nextPagesAfter(last: number, pageCount: number): string | undefined {
  if (last >= pageCount) return undefined
  return `${last + 1}-${Math.min(last + MAX_PAGES_PER_READ, pageCount)}`
}

/** search scans more of the file than a read returns: only snippet lengths leave memory */
const MAX_SEARCH_CHARS = 400_000

export interface PdfSearchResult {
  pageCount: number
  /** one entry per literal occurrence, addressed by 1-based page */
  matches: Array<{ page: number; snippet: string }>
  /** the cap or the scan budget stopped the search before the whole span was covered */
  truncated: boolean
}

/** `radius` characters on each side of a hit, whitespace collapsed, with ellipses */
function pdfSnippet(text: string, index: number, length: number, radius: number): string {
  const start = Math.max(0, index - radius)
  const end = Math.min(text.length, index + length + radius)
  const before = text.slice(start, index).replace(/\s+/g, ' ')
  const after = text.slice(index + length, end).replace(/\s+/g, ' ')
  return `${start > 0 ? '…' : ''}${before}${text.slice(index, index + length)}${after}${
    end < text.length ? '…' : ''
  }`
}

/**
 * Search a PDF's text layer for a literal substring, page by page. The results
 * are addressed the way `read_pdf` addresses content (1-based `page`), so a hit
 * feeds straight back into `read_pdf { pages }`. Scanned pages have no text
 * layer and never match. A `pages` filter narrows the scan; the whole document
 * is scanned otherwise, up to the character budget (reported via `truncated`).
 */
export async function searchPdfFile(
  bytes: Uint8Array,
  options: {
    query: string
    /** 1-based page span to scan, e.g. "1-5,8"; omit for every page */
    pages?: string
    matchCase?: boolean
    maxResults?: number
    /** characters of context on either side of a hit; default 30 */
    context?: number
  },
): Promise<PdfSearchResult> {
  const needle = options.matchCase ? options.query : options.query.toLowerCase()
  const maxResults = Math.max(1, Math.trunc(options.maxResults ?? 100))
  const context = Math.max(0, Math.trunc(options.context ?? 30))
  // A zero-budget probe reports the page count (and surfaces a corrupt/encrypted
  // file) before the page span can be resolved.
  const probe = await readPdfText(bytes, { charBudget: 0 })
  if (!needle) return { pageCount: probe.pageCount, matches: [], truncated: false }
  const pages = options.pages ? parsePageList(options.pages, probe.pageCount) : undefined
  const doc = await readPdfText(bytes, {
    ...(pages ? { pages } : {}),
    charBudget: MAX_SEARCH_CHARS,
  })

  const matches: PdfSearchResult['matches'] = []
  let truncated = false
  outer: for (const pageText of doc.pages) {
    const text = pageText.text
    const haystack = options.matchCase ? text : text.toLowerCase()
    let at = haystack.indexOf(needle)
    while (at !== -1) {
      if (matches.length >= maxResults) {
        truncated = true
        break outer
      }
      matches.push({
        page: pageText.page,
        snippet: pdfSnippet(text, at, needle.length, context),
      })
      at = haystack.indexOf(needle, at + needle.length)
    }
  }
  // the scan budget clipping extraction also means later pages went unsearched
  if (doc.truncated) truncated = true
  return { pageCount: doc.pageCount, matches, truncated }
}

export function createPdfTools(): McpToolDefinition[] {
  return [
    {
      name: 'read_pdf',
      description:
        'Read a PDF file. By default a preview: page count, metadata (title/author), and the first ' +
        "few pages' text, each clipped with a `…(+N chars)` marker — enough to tell what the document " +
        'is. Pass `pages` (e.g. "3" or "1-5,8") to read those pages in full, with page size; one call ' +
        'reads at most 20 pages and the result names the `nextPages` to continue from, so a long ' +
        'document is read window by window. Text is extracted in content order, so complex ' +
        'multi-column layouts may scramble the reading order; scanned pages have no text layer and ' +
        'come back with empty text and hasTextLayer:false.',
      inputSchema: {
        path: z.string().describe('absolute path to a .pdf file'),
        pages: z
          .string()
          .optional()
          .describe(
            '1-based pages to read, e.g. "3" or "1-5,8" (at most 20 per call; pages past the end ' +
              'are ignored); omit for the preview (first pages, clipped)',
          ),
      },
      handler: async (args) => {
        const filePath = String(args.path ?? '')
        if (!isAbsolute(filePath)) throw new Error('path must be absolute')
        if (extname(filePath).toLowerCase() !== PDF_EXT)
          throw new Error('path must point to a .pdf file')
        if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)

        const bytes = new Uint8Array(await readFile(filePath))
        // A zero-budget probe opens the document (surfacing an encrypted or
        // corrupt file as a clean tool error) and reports the page count and
        // metadata. The page span cannot be resolved before the count is known.
        const probe = await readPdfText(bytes, { charBudget: 0 }).catch(() => {
          throw new Error('could not open the PDF — it may be encrypted or corrupt')
        })

        const spec = typeof args.pages === 'string' ? args.pages.trim() : ''
        // No page selection: a bounded preview — page count, metadata and the
        // opening pages' text (each clipped) so the agent learns what the
        // document is, then asks for the pages whose text it wants in full.
        if (!spec) {
          const preview = await readPdfText(bytes, {
            toPage: PREVIEW_PAGES,
            charBudget: MAX_TOTAL_CHARS,
          })
          const last = preview.pages.at(-1)?.page
          const next = last === undefined ? undefined : nextPagesAfter(last, preview.pageCount)
          return {
            path: filePath,
            name: basename(filePath),
            pageCount: preview.pageCount,
            info: preview.info,
            pages: preview.pages.map((p) => ({
              ...p,
              text: clipText(p.text, PREVIEW_CHARS_PER_PAGE).text,
            })),
            ...(next ? { nextPages: next } : {}),
          }
        }

        // Only the requested pages are extracted, so pages the caller never
        // asked for neither spend the budget nor count as truncation. Out-of-range
        // pages were already dropped by parsePageList, and a request wider than a
        // single window is read up to MAX_PAGES_PER_READ pages with the rest named
        // by `nextPages`.
        const wanted = parsePageList(spec, probe.pageCount)
        const window = wanted.slice(0, MAX_PAGES_PER_READ)
        const doc = await readPdfText(bytes, { pages: window, charBudget: MAX_TOTAL_CHARS })
        const last = doc.pages.at(-1)?.page
        const more = window.length < wanted.length || doc.truncated
        const next = last === undefined ? undefined : nextPagesAfter(last, doc.pageCount)
        return {
          path: filePath,
          name: basename(filePath),
          pageCount: doc.pageCount,
          info: doc.info,
          pages: doc.pages,
          truncated: more,
          ...(more && next ? { nextPages: next } : {}),
        }
      },
    },
    {
      name: 'search_pdf',
      description:
        'Search a PDF file for text and return every hit, addressed the way read_pdf addresses ' +
        'content — {page, snippet} with 1-based pages — so a hit feeds straight into ' +
        'read_pdf { pages: String(page) }. `pages` scopes the scan (e.g. "1-5,8"); matchCase makes ' +
        'the match case-sensitive (default is case-insensitive); maxResults caps the hits (default ' +
        '100, max 500); context is the characters of context around each hit (default 30). Scanned ' +
        'pages have no text layer and never match.',
      inputSchema: {
        path: z.string().describe('absolute path to a .pdf file'),
        query: z.string().describe('the text to find'),
        pages: z
          .string()
          .optional()
          .describe('1-based pages to scan, e.g. "1-5,8"; default is every page'),
        matchCase: z.boolean().optional().describe('match case (default is case-insensitive)'),
        maxResults: z
          .number()
          .int()
          .positive()
          .optional()
          .describe('cap on returned hits (default 100, max 500)'),
        context: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe('characters of context around each hit (default 30)'),
      },
      handler: async (args) => {
        const filePath = String(args.path ?? '')
        if (!isAbsolute(filePath)) throw new Error('path must be absolute')
        if (extname(filePath).toLowerCase() !== PDF_EXT)
          throw new Error('path must point to a .pdf file')
        if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)
        const query = String(args.query ?? '')
        if (!query) throw new Error('query must not be empty')
        const bytes = new Uint8Array(await readFile(filePath))
        const result = await searchPdfFile(bytes, {
          query,
          pages: typeof args.pages === 'string' ? args.pages : undefined,
          matchCase: args.matchCase === true,
          maxResults: typeof args.maxResults === 'number' ? args.maxResults : undefined,
          context: typeof args.context === 'number' ? args.context : undefined,
        })
        return { ...result, total: result.matches.length }
      },
    },
  ]
}
