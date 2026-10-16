import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { extname, isAbsolute } from 'node:path'
import { clipText } from '@genoffice/pptx-ops'
import { readPdfText } from '../../../../pdf/src/main/read-text'
import { resolveOpenDocument } from '../mcp/tools/open-documents-tools'
import { asString, listTabs, type DaemonControlDeps } from './shared'

/**
 * The daemon's PDF tools (office_pdf_info / _read / _search) — the daemon's own
 * copy of the readers in `../mcp/tools/pdf-tools`, so the daemon's PDF reads live
 * with the daemon instead of reaching into the built-in MCP server's tool folder.
 * The pdf app is a viewer, so an agent reads a PDF's text layer directly from
 * disk through the same main-process pdfium the editor uses; read-only — the
 * daemon never drives the pdf editor, and the reads run here rather than round-
 * tripping to a renderer.
 */

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

/**
 * Read the named pages of a PDF's text — whole pages only, as many as fit under
 * the cap with the rest named by `nextPages`, so a page never comes back split.
 * No preview: the opening pages are `readPdfInfo`'s job. Backs the daemon's
 * `office_pdf_read`.
 */
export async function readPdfFile(
  bytes: Uint8Array,
  options: { path: string; pages: string },
): Promise<Record<string, unknown>> {
  const filePath = options.path
  // A zero-budget probe opens the document (surfacing an encrypted or corrupt
  // file as a clean error) and reports the page count and metadata. The page
  // span cannot be resolved before the count is known.
  const probe = await readPdfText(bytes, { charBudget: 0 }).catch(() => {
    throw new Error('could not open the PDF — it may be encrypted or corrupt')
  })

  // Only the requested pages are extracted, so pages the caller never asked for
  // neither spend the budget nor count as truncation. Out-of-range pages were
  // already dropped by parsePageList, and a request wider than a single window is
  // read up to MAX_PAGES_PER_READ pages with the rest named by `nextPages`.
  const wanted = parsePageList(options.pages.trim(), probe.pageCount)
  const window = wanted.slice(0, MAX_PAGES_PER_READ)
  // The extractor's own budget cuts the page it crosses, so read the window plain
  // and page it here in whole pages: a page is never split — as many complete
  // pages as fit under the cap come back and the rest are deferred to
  // `nextPages`. The first page always comes back whole, so a reader that asked
  // for one page too large for the cap still makes progress.
  const doc = await readPdfText(bytes, { pages: window })
  const pages: typeof doc.pages = []
  let spent = 0
  for (const page of doc.pages) {
    if (pages.length > 0 && spent + page.text.length > MAX_TOTAL_CHARS) break
    pages.push(page)
    spent += page.text.length
  }
  const last = pages.at(-1)?.page
  const more = pages.length < wanted.length
  const next = last === undefined ? undefined : nextPagesAfter(last, doc.pageCount)
  return {
    path: filePath,
    pageCount: doc.pageCount,
    info: doc.info,
    pages,
    truncated: more,
    ...(more && next ? { nextPages: next } : {}),
  }
}

/**
 * Describe one PDF: its total page count, metadata and the opening pages the
 * same way a read's preview reads them (each page's text, clipped). Backs the
 * daemon's `office_pdf_info`, so an agent can see what a document is — its size
 * and its opening content — in one call, before a fuller read or a search.
 */
export async function readPdfInfo(
  bytes: Uint8Array,
  options: { path: string },
): Promise<Record<string, unknown>> {
  const filePath = options.path
  const preview = await readPdfText(bytes, {
    toPage: PREVIEW_PAGES,
    charBudget: MAX_TOTAL_CHARS,
  }).catch(() => {
    throw new Error('could not open the PDF — it may be encrypted or corrupt')
  })
  return {
    path: filePath,
    pageCount: preview.pageCount,
    info: preview.info,
    pages: preview.pages.map((p) => ({
      ...p,
      text: clipText(p.text, PREVIEW_CHARS_PER_PAGE).text,
    })),
  }
}

/**
 * Resolve the file a PDF tool targets: an open PDF tab (a `tabId`, or a `path`
 * that names one) read from its own file, else an absolute `.pdf` path on disk
 * that need not be open. A PDF has no live editor state (the pdf app is a
 * viewer), so both address the file — matching how office_open resolves a tab.
 */
async function resolvePdfPath(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<string> {
  const ref = asString(args.tabId) ?? asString(args.path)
  if (!ref) throw new Error('pass the PDF as "tabId" (see office_list) or an absolute "path"')
  const doc = deps.tabManager ? resolveOpenDocument(await listTabs(deps.tabManager), ref) : undefined
  if (doc) {
    if (doc.kind !== 'pdf') throw new Error(`"${doc.title}" is not a PDF`)
    if (doc.filePath === undefined) {
      throw new Error(`"${doc.title}" has never been saved, so there is no file to read`)
    }
    return doc.filePath
  }
  if (isAbsolute(ref) && extname(ref).toLowerCase() === '.pdf') {
    if (!existsSync(ref)) throw new Error(`no such file: ${ref}`)
    return ref
  }
  throw new Error(
    `no open tab matches "${ref}"; open a PDF in GenOffice, or pass an absolute .pdf path`,
  )
}

/** office_pdf_info: a PDF's descriptor — page count, metadata and the opening pages' preview */
export async function pdfInfoTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const path = await resolvePdfPath(deps, args)
  const bytes = new Uint8Array(await readFile(path))
  return readPdfInfo(bytes, { path })
}

/** office_pdf_read: read a PDF's named pages (the preview is office_pdf_info's job) */
export async function pdfReadTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const pages = asString(args.pages)
  if (!pages) throw new Error('"pages" is required (e.g. "3" or "1-5,8")')
  const path = await resolvePdfPath(deps, args)
  const bytes = new Uint8Array(await readFile(path))
  return readPdfFile(bytes, { path, pages })
}

/** office_pdf_search: find text in a PDF's text layer, hits addressed by page */
export async function pdfSearchTool(
  deps: DaemonControlDeps,
  args: Record<string, unknown>,
): Promise<Record<string, unknown>> {
  const query = asString(args.query)
  if (!query) throw new Error('"query" is required (the text to find)')
  const path = await resolvePdfPath(deps, args)
  const bytes = new Uint8Array(await readFile(path))
  const result = await searchPdfFile(bytes, {
    query,
    pages: typeof args.pages === 'string' ? args.pages : undefined,
    matchCase: args.matchCase === true,
    maxResults: typeof args.maxResults === 'number' ? args.maxResults : undefined,
    context: typeof args.context === 'number' ? args.context : undefined,
  })
  return { ...result, total: result.matches.length }
}

