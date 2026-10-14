import { z } from 'zod'
import { outlineToOps, parsePptxOutline, type PptxSourceFormat } from '../pptx-outline'
import { createPptxViaCli } from '../headless-cli'
import { resolveOutputPath } from './document-tools'
import { generateExtension } from './formats'
import type { FamilyDriver, SessionHost, TargetResolver } from './session-tools'
import { contentTarget } from './session-tools'
import type { McpToolDefinition } from '../mcp-server'
import type { CliRunner } from '../cli-runner'

/**
 * Slides (pptx) tool surface for the MCP server.
 *
 * Two paths, mirroring the docx tools:
 * - headless `create_pptx`: outline -> ops, handed to the bundled `genoffice`
 *   CLI (`create --type pptx --ops`), gated behind the "background generation"
 *   setting. No deck-building code lives here;
 * - a visible deck session: the tools drive a real slides tab the user watches.
 *   Slides sessions live in the slides main process, so the control talks to
 *   main-process state directly (no renderer bridge) — see
 *   `apps/shell/src/main/mcp/slides-bridge.ts`.
 */

export interface SlidesToolDeps {
  /** directory generated files land in when the caller gives no path */
  defaultSaveDir: () => string
  /** expose the headless create_pptx tool; default true — the shell passes the user's setting */
  background?: boolean
  /** visible-deck control; absent in headless/unit runs, which drops the session tools */
  slides?: SlidesControl
  /** the bundled genoffice CLI, used by the headless tool */
  cli?: CliRunner
  /**
   * Resolve a caller-supplied document reference (tab id or path) to the
   * webContents of that open tab, so the deck tools can edit a presentation the
   * *user* has open rather than only the session's own blank one. Absent in
   * headless/unit runs, where a `document` argument is refused.
   */
  resolveTarget?: TargetResolver
}

/** transaction request for the visible deck (executor semantics, ops are EMU-space) */
export interface SlidesTxnRequest {
  ops: unknown[]
  isolation?: 'atomic' | 'per_op'
  dryRun?: boolean
}

/** how `read_deck` pages and clips, mirroring the CLI's `slides read --slide/--full/--max-chars`. */
export interface ReadDeckOptions {
  /** 0-based slide index to read; omit for every slide */
  slide?: number
  /** return whole text and every table row instead of the preview clip */
  full?: boolean
  /** clip text to this many characters (default 300); ignored when `full` */
  maxChars?: number
}

/**
 * Visible-deck control implemented in the shell main process against the slides
 * main-process sessions (`slides-bridge.ts`).
 */
export interface SlidesControl {
  /** open a fresh blank slides tab; resolves to its webContents id */
  openBlankTab: () => Promise<number>
  /** apply one transaction to the tab's session (history + journal + live render included) */
  runTxn: (wcId: number, req: SlidesTxnRequest) => Promise<unknown>
  /** read the visible deck (slides/elements with ids, text and geometry); `opts` pages/clips like the CLI's `slides read` */
  readDeck: (wcId: number, opts?: ReadDeckOptions) => Promise<unknown>
  /** save the session's deck to an absolute path */
  saveDeck: (wcId: number, path: string, overwrite: boolean) => Promise<{ path: string }>
}

const PPTX_EXT = `.${generateExtension('pptx')}`

/** the headless tool: outline -> ops -> pptx bytes written straight to disk */
function createHeadlessPptxTool(deps: SlidesToolDeps): McpToolDefinition {
  return {
    name: 'create_pptx',
    description:
      'Create a PowerPoint .pptx file from an outline and save it to disk without opening the app UI. ' +
      'By default `outline` is Markdown where `# Title` starts each slide, `## text` adds a bold line, ' +
      '`- text` a bullet (indent two spaces per level) and `1. text` a numbered bullet. Use ' +
      'format:"json" to pass {"slides":[{"title":"...","bullets":["...","...",{"text":"...","level":1}]}]} instead. ' +
      'Returns the absolute path of the written file.',
    inputSchema: {
      title: z.string().describe('presentation title, used as the file name'),
      outline: z
        .string()
        .describe(
          'Markdown outline (default) or a JSON string of {"slides":[...]} when format is "json"',
        ),
      format: z.enum(['markdown', 'json']).optional().describe('outline format; default markdown'),
      path: z
        .string()
        .optional()
        .describe('absolute output path; default is a new file in the default save folder'),
      overwrite: z
        .boolean()
        .optional()
        .describe('allow replacing an existing file at `path`; default false'),
    },
    handler: async (args) => {
      const title = String(args.title ?? '').trim()
      if (!title) throw new Error('title must not be empty')
      if (typeof args.outline !== 'string') {
        throw new Error('outline must be a string (markdown by default, JSON with format:"json")')
      }
      if (!deps.cli) throw new Error('headless generation is not available in this build')
      const format: PptxSourceFormat = args.format === 'json' ? 'json' : 'markdown'
      const ops = outlineToOps(parsePptxOutline(format, args.outline))

      const targetPath = resolveOutputPath({
        defaultSaveDir: deps.defaultSaveDir,
        ext: PPTX_EXT,
        title,
        requestedPath: typeof args.path === 'string' ? args.path : undefined,
        overwrite: args.overwrite === true,
      })

      const { summary, outputPath } = await createPptxViaCli(deps.cli, {
        ops,
        out: targetPath,
        overwrite: args.overwrite === true,
      })
      return { path: outputPath, summary }
    },
  }
}

export function createSlidesTools(deps: SlidesToolDeps, host: SessionHost): McpToolDefinition[] {
  return [
    // headless generation is opt-in, same rule as create_docx
    ...(deps.background === false ? [] : [createHeadlessPptxTool(deps)]),
    ...createDeckContentTools(deps, host),
  ]
}

/**
 * The pptx session lifecycle as seen by the shared create_session / save_session
 * tools. The content tools below address the tab this driver opened.
 */
export function slidesDriver(slides: SlidesControl): FamilyDriver {
  return {
    family: 'pptx',
    openBlankTab: () => slides.openBlankTab(),
    save: (wcId, path, overwrite) => slides.saveDeck(wcId, path, overwrite),
  }
}

/** search_deck result-count/context bounds (default 100 / 30, hard cap 500 / 200) */
const DECK_SEARCH_DEFAULT_RESULTS = 100
const DECK_SEARCH_MAX_RESULTS = 500
const DECK_SEARCH_CONTEXT_DEFAULT = 30
const DECK_SEARCH_CONTEXT_MAX = 200

function clampSearchResults(value: unknown): number {
  const n = Math.trunc(Number(value))
  if (!Number.isFinite(n) || n < 1) return DECK_SEARCH_DEFAULT_RESULTS
  return Math.min(n, DECK_SEARCH_MAX_RESULTS)
}

function clampSearchContext(value: unknown): number {
  const n = Math.trunc(Number(value))
  if (!Number.isFinite(n) || n < 0) return DECK_SEARCH_CONTEXT_DEFAULT
  return Math.min(n, DECK_SEARCH_CONTEXT_MAX)
}

/** the `read_deck {full:true}` surface this scan reads (see pptx-ops describe-deck.ts) */
interface DeckElementLike {
  id?: string | null
  type?: string
  text?: string
  children?: DeckElementLike[]
}
interface DeckPageLike {
  index?: number
  elements?: DeckElementLike[]
  notes?: string
}

/** a pptx hit, addressed by 0-based slide + element id (the apply_slide_ops vocabulary) */
export interface DeckSearchHit {
  slide: number
  el: string | null
  type: string
  snippet: string
}

/** `radius` characters on each side of a hit, whitespace collapsed, with ellipses */
function deckSnippet(text: string, index: number, length: number, radius: number): string {
  const start = Math.max(0, index - radius)
  const end = Math.min(text.length, index + length + radius)
  const before = text.slice(start, index).replace(/\s+/g, ' ')
  const after = text.slice(index + length, end).replace(/\s+/g, ' ')
  return `${start > 0 ? '…' : ''}${before}${text.slice(index, index + length)}${after}${
    end < text.length ? '…' : ''
  }`
}

/**
 * Scan a full deck read for a literal substring. Text lives on leaf elements
 * (a group has no text of its own), so a group is walked into and each hit is
 * attributed to the child whose id apply_slide_ops addresses. Speaker notes are
 * scanned too, reported with a null `el` and type "notes".
 */
export function scanDeck(
  deck: unknown,
  opts: { query: string; slide?: number; matchCase: boolean; maxResults: number; context: number },
): { slides: number; total: number; truncated: boolean; matches: DeckSearchHit[] } {
  const pages = ((deck as { pages?: DeckPageLike[] } | null | undefined)?.pages ??
    []) as DeckPageLike[]
  const needle = opts.matchCase ? opts.query : opts.query.toLowerCase()
  const matches: DeckSearchHit[] = []
  let truncated = false
  if (!needle) return { slides: pages.length, total: 0, truncated, matches }

  const push = (slide: number, el: string | null, type: string, text: string): boolean => {
    const hay = opts.matchCase ? text : text.toLowerCase()
    let at = hay.indexOf(needle)
    while (at !== -1) {
      if (matches.length >= opts.maxResults) {
        truncated = true
        return false
      }
      matches.push({ slide, el, type, snippet: deckSnippet(text, at, needle.length, opts.context) })
      at = hay.indexOf(needle, at + needle.length)
    }
    return true
  }

  const walk = (el: DeckElementLike, slide: number): boolean => {
    if (typeof el.text === 'string' && el.text) {
      if (!push(slide, el.id ?? null, el.type ?? 'element', el.text)) return false
    }
    for (const child of el.children ?? []) {
      if (!walk(child, slide)) return false
    }
    return true
  }

  outer: for (const page of pages) {
    const slide = typeof page.index === 'number' ? page.index : 0
    if (opts.slide !== undefined && slide !== opts.slide) continue
    for (const el of page.elements ?? []) {
      if (!walk(el, slide)) break outer
    }
    if (typeof page.notes === 'string' && page.notes) {
      if (!push(slide, null, 'notes', page.notes)) break outer
    }
  }
  return { slides: pages.length, total: matches.length, truncated, matches }
}

/**
 * Visible-deck content tools: edit the deck the shared session opened, so the
 * user watches each page take shape. The control mutates the main-process
 * session (the same one the app's own editing and built-in AI use), so every
 * step lands in the app's undo history and the tab re-renders live.
 *
 * Only registered when the shell wired a SlidesControl — headless/unit runs
 * keep the file-only surface.
 */
function createDeckContentTools(deps: SlidesToolDeps, host: SessionHost): McpToolDefinition[] {
  const slides = deps.slides
  if (!slides) return []

  const target = contentTarget(host, deps.resolveTarget, 'pptx')
  const documentField = z
    .string()
    .optional()
    .describe(
      "tab id or path of an open presentation to edit instead of the session's own tab " +
        '(from open_documents {action:"list"}); the edit lands in that live deck, and a ' +
        'mutating tool switches the UI to that tab first so the user sees it happen',
    )

  return [
    {
      name: 'read_deck',
      description:
        'Read the visible presentation: every slide with its elements (ids, types, text, geometry) ' +
        'so you can target follow-up edits. Element ids and slide indexes are what apply_slide_ops ' +
        'addresses; geometry is document-space EMU (1 px = 9525 EMU on a standard 16:9 deck). Text ' +
        'and table rows are previewed (clipped) by default; pass full:true for the whole text and ' +
        'every table row, maxChars to move the clip, and slide to read one slide (0-based). ' +
        'Same shape as the genoffice CLI `slides read`.',
      inputSchema: {
        document: documentField,
        slide: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe('0-based slide index to read; default is every slide'),
        full: z
          .boolean()
          .optional()
          .describe('return whole text and every table row (no preview clipping)'),
        maxChars: z
          .number()
          .int()
          .positive()
          .max(1_000_000)
          .optional()
          .describe('clip text to this many characters; default 300; ignored when full is true'),
      },
      handler: async (args) =>
        slides.readDeck(await target(args.document), {
          slide: typeof args.slide === 'number' ? args.slide : undefined,
          full: args.full === true,
          maxChars: typeof args.maxChars === 'number' ? args.maxChars : undefined,
        }),
    },
    {
      name: 'search_deck',
      description:
        'Search inside the visible presentation and return every hit, addressed the way read_deck ' +
        'and apply_slide_ops address content — {slide, el, type, snippet}, a 0-based slide plus the ' +
        'element id — so a hit feeds straight into read_deck { slide } or apply_slide_ops ' +
        "{ slide, el }. Every slide's element text and speaker notes are scanned; pass slide to " +
        'search one 0-based slide. matchCase makes the match case-sensitive (default is ' +
        'case-insensitive); maxResults caps the hits (default 100, max 500); context is the ' +
        'characters of context around each hit (default 30).',
      inputSchema: {
        document: documentField,
        query: z.string().describe('the text to find'),
        slide: z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe('0-based slide to search; default is every slide'),
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
        const query = String(args.query ?? '')
        if (!query) throw new Error('query must not be empty')
        const deck = await slides.readDeck(await target(args.document), { full: true })
        return scanDeck(deck, {
          query,
          slide: typeof args.slide === 'number' ? args.slide : undefined,
          matchCase: args.matchCase === true,
          maxResults: clampSearchResults(args.maxResults),
          context: clampSearchContext(args.context),
        })
      },
    },
    {
      name: 'apply_slide_ops',
      description:
        'Apply slide editing ops to the visible presentation as one transaction (atomic by default, ' +
        'plan-validated with snapshot rollback; per_op isolation skips failures instead). Ops address ' +
        '{slide, el} and use document-space EMU. Common ops: setText, setFont, setParagraphFormat, ' +
        'addElement (textbox/shapes/lines), addTable, addChart, addBlankSlide, deleteSlide, moveSlide, ' +
        'duplicateSlide, deleteElement, setFill, setStroke, setTransform, setBackground, applyTheme, ' +
        'findReplace, setTableCell, tableStructure, setTransition, setNotes. A failing atomic batch ' +
        'changes nothing — fix the op named in the error and resend the whole batch. ' +
        'Use read_deck for ids and geometry first.',
      inputSchema: {
        ops: z.array(z.any()).describe('array of op objects (at most 50 per transaction)').max(50),
        isolation: z
          .enum(['atomic', 'per_op'])
          .optional()
          .describe('atomic (default) applies all-or-nothing; per_op applies independently'),
        dryRun: z
          .boolean()
          .optional()
          .describe('validate and plan the batch without changing the deck'),
        document: documentField,
      },
      handler: async (args) => {
        const wc = await target(args.document, { focus: true })
        const result = (await slides.runTxn(wc, {
          ops: (Array.isArray(args.ops) ? args.ops : []) as unknown[],
          isolation:
            args.isolation === 'per_op' || args.isolation === 'atomic' ? args.isolation : undefined,
          dryRun: args.dryRun === true,
        })) as { applied?: boolean; failures?: Array<{ index: number; error: string }> }
        // a rejected transaction is a tool-level error so the caller reacts to it;
        // dry-run results (plan + failures) stay a normal response
        if (result?.applied === false && args.dryRun !== true) {
          const details = (result.failures ?? []).map((f) => `[${f.index}] ${f.error}`).join('\n')
          throw new Error(details || 'the transaction could not be applied')
        }
        return result
      },
    },
  ]
}
