import { existsSync, readFileSync } from 'node:fs'
import { extname } from 'node:path'
import { assertAllowed, type PathContext } from '../fs'
import { isAbsolute, resolve } from 'node:path'
import { createBlankPptx, openPptx, savePptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { runTxn, type Op, type TxnResult } from '@genoffice/pptx-ops'
import { CliError, EXIT } from '../result'

// The deck readout lives in `@genoffice/pptx-ops` so `slides read` and the app's
// MCP `read_deck` share one implementation; re-exported for this module's callers.
export {
  describeDeck,
  describeLayouts,
  PREVIEW_CHARS,
  type DeckSummary,
  type EffectiveStyle,
  type ElementSummary,
  type LayoutSummary,
  type SlideSummary,
} from '@genoffice/pptx-ops'

/** Image/media ops that take bytes, with the (possibly nested) field each one reads. A local file path there is read for the caller. */
const BYTES_FIELDS: Record<string, string[]> = {
  addPicture: ['bytes'],
  replacePicture: ['bytes'],
  addMedia: ['bytes', 'poster.bytes'],
  setImageFill: ['source.bytes'],
  setBackground: ['source.bytes'],
}

export async function openDeck(bytes: Uint8Array): Promise<OpenedPptx> {
  return openPptx(bytes)
}

export async function blankDeck(): Promise<OpenedPptx> {
  return openPptx(await createBlankPptx())
}

export function saveDeck(opened: OpenedPptx): Promise<Uint8Array> {
  return savePptx(opened)
}

export function parseOps(text: string, source: string): Op[] {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new CliError(
      EXIT.usage,
      `${source}: not valid JSON (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  const ops = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!ops) {
    throw new CliError(
      EXIT.usage,
      `${source}: expected an array of ops or { "ops": [...] }`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  const bad = ops.findIndex(
    (op) =>
      !op ||
      typeof op !== 'object' ||
      Array.isArray(op) ||
      typeof (op as { op?: unknown }).op !== 'string',
  )
  if (bad !== -1) {
    throw new CliError(
      EXIT.usage,
      `${source}: ops[${bad}] must be an object with a string "op"`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  return ops as Op[]
}

/** Resolves `bytes`-style fields that name a local file into the bytes the ops expect. */
export function inlineLocalFiles(ops: Op[], ctx: PathContext): Op[] {
  return ops.map((op) => {
    const fields = Object.hasOwn(BYTES_FIELDS, op.op) ? BYTES_FIELDS[op.op] : undefined
    if (!fields) return op
    const next = structuredClone(op) as Op
    for (const field of fields) {
      const value = getPath(next, field)
      if (typeof value !== 'string' || value.startsWith('data:')) continue
      const path = isAbsolute(value) ? value : resolve(ctx.cwd, value)
      if (!existsSync(path)) continue
      assertAllowed(path, ctx.env, 'read')
      setPath(next, field, new Uint8Array(readFileSync(path)))
      // the op's sibling "ext" is implied by the file name
      const extField = field.replace(/bytes$/, 'ext')
      if (!getPath(next, extField)) setPath(next, extField, extname(path).slice(1).toLowerCase())
    }
    return next
  })
}

function getPath(obj: Record<string, unknown>, dotted: string): unknown {
  return dotted.split('.').reduce<unknown>((cur, key) => {
    return cur && typeof cur === 'object' ? (cur as Record<string, unknown>)[key] : undefined
  }, obj)
}

function setPath(obj: Record<string, unknown>, dotted: string, value: unknown): void {
  const keys = dotted.split('.')
  let cur = obj
  for (const key of keys.slice(0, -1)) cur = cur[key] as Record<string, unknown>
  cur[keys[keys.length - 1]!] = value
}

export interface ApplyOptions {
  isolation?: 'atomic' | 'per_op'
  dryRun?: boolean
}

export function applyOps(opened: OpenedPptx, ops: Op[], opts: ApplyOptions): TxnResult {
  return runTxn(opened, { ops, isolation: opts.isolation, dryRun: opts.dryRun })
}
