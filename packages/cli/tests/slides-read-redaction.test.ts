import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

/**
 * The CLI half of the shared read contract: `slides read` and the app's MCP
 * `read_deck` both funnel through `describeDeck`, so a deck read from disk must
 * withhold exactly as one read from an open tab does. This exercises the file
 * path end to end (write the mark, save, re-open, read).
 */

const INCH = 914400
const LABEL = 'client phone'
const SECRET = '13800138000'

async function deckWithWithheldRun(): Promise<string> {
  const dir = tempDir()
  const ops = join(dir, 'create.json')
  writeFileSync(
    ops,
    JSON.stringify([
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: INCH, y: INCH, cx: 5 * INCH, cy: INCH },
        paragraphs: [
          { runs: [{ text: 'call ' }, { text: SECRET, redact: LABEL, underline: true }] },
        ],
      },
    ]),
  )
  const out = join(dir, 'deck.pptx')
  const created = await run(['create', '--type', 'pptx', '--ops', ops, '--out', out, '--json'])
  expect(created.code).toBe(0)
  return out
}

describe('slides read withholds marked runs', () => {
  it('shows the placeholder and never the words, in preview and full', async () => {
    const file = await deckWithWithheldRun()
    for (const flags of [[], ['--full']]) {
      const r = await run(['slides', 'read', file, ...flags, '--json'])
      expect(r.code).toBe(0)
      const text = JSON.stringify(r.json().detail)
      expect(text).toContain('{{client phone}}')
      expect(text).not.toContain(SECRET)
    }
  })
})
