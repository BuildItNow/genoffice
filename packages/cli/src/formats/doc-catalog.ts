import { didYouMean } from '../suggest'
import type { JsonSchema, OpEntry } from '../op-catalog'

/**
 * The reference the live app's `office_doc_apply_ops_help` serves: `list` names
 * categories and returns each one's op names, `ops` names ops and returns each
 * one's signature, note and (for block ops) how the op is addressed. Mirrors the
 * sheet catalog's `sheetHelp`.
 *
 * Built on `docsCatalog()` — the very object `genoffice guide docs` prints — so
 * the reference cannot drift from what `docs apply --ops` accepts. The engine
 * module is imported lazily: the daemon bundles this file and must not pull the
 * whole docx engine onto its startup path.
 */

export interface DocOpHelp {
  op: string
  group: string
  signature: string
  note?: string
  /** block ops: whether a `target` is required, optional or refused for this op */
  target?: 'required' | 'optional' | 'none'
  /** block ops: the keys each target may carry (blockIndexes / nodeType / headingLevel …) */
  keys?: string[]
  /** content ops: the op object's input schema */
  schema?: JsonSchema
}

export interface DocCategoryHelp {
  name: string
  summary?: string
  ops: string[]
}

export interface DocHelp {
  /** present when categories were asked for (via `list`) */
  categories?: DocCategoryHelp[]
  /** present when ops were asked for (via `ops`) */
  ops?: DocOpHelp[]
  /** names that matched nothing; present only when at least one name was unknown */
  unknown?: {
    categories: { name: string; suggestion?: string }[]
    ops: { op: string; suggestion?: string }[]
  }
}

/** one catalog entry as the help reply carries it */
function opHelp(entry: OpEntry): DocOpHelp {
  return {
    op: entry.op,
    group: entry.group,
    signature: entry.signature,
    ...(entry.note ? { note: entry.note } : {}),
    ...(entry.target ? { target: entry.target } : {}),
    ...(entry.keys ? { keys: entry.keys } : {}),
    ...(entry.schema ? { schema: entry.schema } : {}),
  }
}

/**
 * `list` names categories and returns each one's op names, `ops` names ops and
 * returns each one's signature/note/target (the same text `guide docs <op>`
 * prints). Nothing given returns nothing. An unknown name is reported on its own
 * with a suggestion rather than failing the whole call.
 */
export async function docHelp(req: {
  list?: readonly string[]
  ops?: readonly string[]
}): Promise<DocHelp> {
  const { docsCatalog } = await import('./docx')
  const { catalog } = await docsCatalog()
  const groups = new Map(catalog.groups.map((g) => [g.name, g]))
  const byOp = new Map(catalog.ops.map((e) => [e.op, e]))
  const unknownCategories: { name: string; suggestion?: string }[] = []
  const unknownOps: { op: string; suggestion?: string }[] = []
  const out: DocHelp = {}

  if (req.list) {
    const categories: DocCategoryHelp[] = []
    for (const name of req.list) {
      const group = groups.get(name)
      if (!group) {
        const suggestion = didYouMean(name, groups.keys())
        unknownCategories.push(suggestion ? { name, suggestion } : { name })
        continue
      }
      categories.push({
        name: group.name,
        ...(group.summary ? { summary: group.summary } : {}),
        ops: group.ops,
      })
    }
    out.categories = categories
  }

  if (req.ops) {
    const ops: DocOpHelp[] = []
    for (const name of req.ops) {
      const entry = byOp.get(name)
      if (!entry) {
        const suggestion = didYouMean(name, byOp.keys())
        unknownOps.push(suggestion ? { op: name, suggestion } : { op: name })
        continue
      }
      ops.push(opHelp(entry))
    }
    out.ops = ops
  }

  if (unknownCategories.length || unknownOps.length) {
    out.unknown = { categories: unknownCategories, ops: unknownOps }
  }

  return out
}
