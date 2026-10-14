import { flagBool, flagString, type ParsedArgs } from './args'
import { CliError, EXIT } from './result'

// The clip marker is shared with the app's MCP reads (`@genoffice/pptx-ops`), so a
// preview reads the same whether it came from the CLI or an open tab.
export { clipText } from '@genoffice/pptx-ops'

/** `--full` lifts the cap, `--max-chars n` moves it. */
export function previewChars(args: ParsedArgs, fallback: number): number {
  if (flagBool(args, 'full')) return Infinity
  const raw = flagString(args, 'max-chars')
  if (raw === undefined) return fallback
  const n = Number(raw)
  if (!Number.isInteger(n) || n < 1 || n > MAX_PREVIEW_CHARS)
    throw new CliError(
      EXIT.usage,
      `--max-chars must be a positive integer up to ${MAX_PREVIEW_CHARS} (got ${raw}); use --full for uncapped output`,
      undefined,
      {
        reason: 'invalid_argument',
      },
    )
  return n
}

/**
 * Upper bound for --max-chars: the flag exists to keep previews small, and an
 * unbounded value defeats it (agents dumping whole documents into context).
 * --full remains the escape hatch for legitimately uncapped output.
 */
export const MAX_PREVIEW_CHARS = 1_000_000
