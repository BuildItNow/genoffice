import { z, type ZodRawShape } from 'zod'

/**
 * Deferred function tools for `office_doc_apply_func`.
 *
 * A spec looks exactly like an MCP tool (name + description + Zod shape) but is
 * not registered: one generic tool names it, the daemon validates the call with
 * the same shape, and `office_doc_func_help` expands it on demand. That keeps
 * `tools/list` small without giving up strict, per-function schemas.
 */

export interface DocFuncSpec {
  name: string
  description: string
  shape: ZodRawShape
  /** the help category this func belongs to (see `DocFuncCategory`) */
  category: string
  example?: Record<string, unknown>
}

/** one help category; `summary` is what the unnamed help index lists */
export interface DocFuncCategory {
  name: string
  summary: string
}

export interface DocFuncHelpEntry {
  func: string
  description: string
  signature: string
  schema?: FuncJsonSchema
  example?: Record<string, unknown>
}

export interface DocFuncCategoryHelpEntry extends DocFuncCategory {
  /** present when the category was asked for via `list` */
  funcs?: string[]
}

export interface DocFuncHelp {
  /** always present: the category index, plus `funcs` names when they were asked for */
  categories?: DocFuncCategoryHelpEntry[]
  /** present when func names were asked for */
  funcs?: DocFuncHelpEntry[]
  /** names that matched nothing; present only when at least one name was unknown */
  unknown?: {
    categories: { name: string; suggestion?: string }[]
    funcs: { func: string; suggestion?: string }[]
  }
}

/** the JSON Schema subset zod's `toJSONSchema` emits for the func shapes */
export interface FuncJsonSchema {
  type?: string | string[]
  const?: unknown
  enum?: unknown[]
  properties?: Record<string, FuncJsonSchema>
  required?: string[]
  items?: FuncJsonSchema
  anyOf?: FuncJsonSchema[]
  oneOf?: FuncJsonSchema[]
  additionalProperties?: boolean | FuncJsonSchema
}

export interface DocFuncs {
  specs: DocFuncSpec[]
  /** a non-empty tuple, ready for `z.enum` */
  names: [string, ...string[]]
  byName: ReadonlyMap<string, DocFuncSpec>
  /**
   * Validate one call against its func's own shape (strict: an unknown field is
   * an error, not silently dropped) and return the normalized args.
   */
  prepare(input: Record<string, unknown>): Record<string, unknown>
  /**
   * Tiered expansion: no names → the category index only; `list` → each named
   * category's func names; `funcs` → each named func's full field schema and
   * example. One call never dumps every func.
   */
  help(req: { list?: readonly string[]; funcs?: readonly string[] }): DocFuncHelp
}

export function createDocFuncs(specs: DocFuncSpec[], categories: DocFuncCategory[]): DocFuncs {
  const byName = new Map(specs.map((spec) => [spec.name, spec]))
  const names = specs.map((spec) => spec.name) as [string, ...string[]]
  const schemas = new Map<string, FuncJsonSchema>()
  const categoryByName = new Map(categories.map((category) => [category.name, category]))
  const funcsByCategory = new Map(categories.map((category) => [category.name, [] as string[]]))
  // fail at daemon startup, not on a help call, when the tables drift apart
  for (const spec of specs) {
    const bucket = funcsByCategory.get(spec.category)
    if (!bucket) {
      throw new Error(`doc func "${spec.name}" has unknown category "${spec.category}"`)
    }
    bucket.push(spec.name)
  }
  for (const category of categories) {
    if (!funcsByCategory.get(category.name)!.length) {
      throw new Error(`doc func category "${category.name}" has no funcs`)
    }
  }

  /** a func's args as JSON Schema — the very shape the daemon validates with */
  function schemaFor(spec: DocFuncSpec): FuncJsonSchema {
    let schema = schemas.get(spec.name)
    if (!schema) {
      schema = z.toJSONSchema(z.strictObject(spec.shape), {
        io: 'input',
        unrepresentable: 'any',
      }) as FuncJsonSchema
      schemas.set(spec.name, schema)
    }
    return schema
  }

  return {
    specs,
    names,
    byName,
    prepare(input) {
      const name = String(input.func ?? '')
      const spec = byName.get(name)
      if (!spec) {
        throw new Error(`unknown func "${name}"; call office_doc_func_help for the list`)
      }
      const parsed = z.strictObject(spec.shape).safeParse(input.args ?? {})
      if (!parsed.success) {
        const issues = parsed.error.issues
          .map((issue) => `${issue.path.join('.') || 'args'}: ${issue.message}`)
          .join('; ')
        throw new Error(
          `invalid args for func "${name}": ${issues}. Signature: ` +
            `${funcSignature(schemaFor(spec))} — call ` +
            `office_doc_func_help({ funcs: ["${name}"] }) for every field.`,
        )
      }
      return { ...input, args: parsed.data }
    },
    help(req) {
      // no names at all: the category index, so the first call stays tiny
      if (!req.list && !req.funcs) {
        return {
          categories: categories.map((category) => ({
            name: category.name,
            summary: category.summary,
          })),
        }
      }
      const categoriesOut: DocFuncCategoryHelpEntry[] = []
      const unknownCategories: { name: string; suggestion?: string }[] = []
      for (const name of req.list ?? []) {
        const category = categoryByName.get(name)
        if (!category) {
          const suggestion = nearestName(name, categoryByName.keys())
          unknownCategories.push(suggestion ? { name, suggestion } : { name })
          continue
        }
        categoriesOut.push({
          name: category.name,
          summary: category.summary,
          funcs: funcsByCategory.get(category.name)!,
        })
      }
      const funcs: DocFuncHelpEntry[] = []
      const unknownFuncs: { func: string; suggestion?: string }[] = []
      for (const name of req.funcs ?? []) {
        const spec = byName.get(name)
        if (!spec) {
          const suggestion = nearestName(name, names)
          unknownFuncs.push(suggestion ? { func: name, suggestion } : { func: name })
          continue
        }
        funcs.push({
          func: spec.name,
          description: spec.description,
          signature: funcSignature(schemaFor(spec)),
          schema: schemaFor(spec),
          ...(spec.example ? { example: spec.example } : {}),
        })
      }
      return {
        ...(req.list ? { categories: categoriesOut } : {}),
        ...(req.funcs ? { funcs } : {}),
        ...(unknownCategories.length || unknownFuncs.length
          ? { unknown: { categories: unknownCategories, funcs: unknownFuncs } }
          : {}),
      }
    },
  }
}

/** a compact field signature, like the op catalogs print: `{a, b?, c: n|[string]}` */
function funcSignature(schema: FuncJsonSchema): string {
  const required = new Set(schema.required ?? [])
  const fields = Object.entries(schema.properties ?? {}).map(([key, value]) => {
    const type = schemaTypeText(value)
    return `${key}${required.has(key) ? '' : '?'}${type ? `: ${type}` : ''}`
  })
  return `{${fields.join(', ')}}`
}

function schemaTypeText(schema: FuncJsonSchema, inUnion = false): string {
  if (schema.const !== undefined) return JSON.stringify(schema.const)
  if (schema.enum) return schema.enum.map(String).join('|')
  const variants = schema.anyOf ?? schema.oneOf
  if (variants) return variants.map((v) => schemaTypeText(v, true)).join('|')
  const types = Array.isArray(schema.type) ? schema.type : schema.type ? [schema.type] : []
  if (types.length === 0) return 'any'
  if (types.length > 1) return types.map((t) => scalarTypeText(t, true)).join('|')
  const [type] = types
  if (type === 'array') return `[${schema.items ? schemaTypeText(schema.items, true) : 'any'}]`
  if (type === 'object') {
    if (
      !schema.properties &&
      schema.additionalProperties &&
      typeof schema.additionalProperties === 'object'
    ) {
      return `{"<key>": ${schemaTypeText(schema.additionalProperties, true)}}`
    }
    return funcSignature(schema)
  }
  return scalarTypeText(type!, inUnion)
}

function scalarTypeText(type: string, inUnion: boolean): string {
  switch (type) {
    case 'string':
      return inUnion ? 'string' : ''
    case 'number':
    case 'integer':
      return 'n'
    case 'boolean':
      return 'bool'
    default:
      return type
  }
}

/** optimal-string-alignment distance; one transposition counts as one edit */
function editDistance(a: string, b: string): number {
  const rows = a.length + 1
  const cols = b.length + 1
  const d: number[] = new Array(rows * cols)
  for (let i = 0; i < rows; i++) d[i * cols] = i
  for (let j = 0; j < cols; j++) d[j] = j
  for (let i = 1; i < rows; i++) {
    for (let j = 1; j < cols; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1
      let v = Math.min(
        d[(i - 1) * cols + j]! + 1,
        d[i * cols + j - 1]! + 1,
        d[(i - 1) * cols + j - 1]! + cost,
      )
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        v = Math.min(v, d[(i - 2) * cols + j - 2]! + 1)
      }
      d[i * cols + j] = v
    }
  }
  return d[rows * cols - 1]!
}

/** the closest candidate when the typo is small enough to be one */
export function nearestName(input: string, candidates: Iterable<string>): string | undefined {
  const needle = input.toLowerCase()
  if (!needle) return undefined
  const budget = Math.max(2, Math.floor(needle.length / 3))
  let best: { name: string; d: number } | undefined
  for (const name of candidates) {
    const d = editDistance(needle, name.toLowerCase())
    if (d > budget || d >= needle.length || (best && d >= best.d)) continue
    best = { name, d }
  }
  return best?.name
}
