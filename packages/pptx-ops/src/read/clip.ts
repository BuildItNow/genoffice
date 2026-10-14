/** Clip to `max` characters; the marker names how much is missing so a reader can decide whether to fetch it. */
export function clipText(text: string, max: number): { text: string; truncated: boolean } {
  if (text.length <= max) return { text, truncated: false }
  return { text: `${text.slice(0, max)}…(+${text.length - max} chars)`, truncated: true }
}
