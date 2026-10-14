import { describe, it, expect, beforeAll } from 'vitest'
import {
  addElement,
  addPicture,
  createBlankPptx,
  openPptx,
  setElementRedaction,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { describeDeck } from '../src/read/describe-deck'

/**
 * The readout is the last place a withheld span could leak: `describeDeck` turns
 * a deck into text a model reads, so a redacted run must become its `{{label}}`
 * marker — never the words it covers — and a redacted picture must surface its
 * marker rather than disappear. `read_deck` (open tab) and `slides read` (file)
 * share this one function, so guarding it here covers both.
 */

const LABEL = 'client phone'
const SECRET = '13800138000'
const PIC_LABEL = 'id card'

// 1x1 red PNG
const PNG_1PX = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

let opened: OpenedPptx

beforeAll(async () => {
  opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  addElement(slide, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    paragraphs: [{ runs: [{ text: 'call ' }, { text: SECRET, redact: LABEL, underline: true }] }],
  })
  const pic = addPicture(opened, slide, {
    bytes: PNG_1PX,
    ext: 'png',
    offset: { x: 0, y: 914400, cx: 914400, cy: 914400 },
  })!
  setElementRedaction(slide, pic.id, PIC_LABEL)
})

const elementOf = (type: string, full = false) =>
  describeDeck(opened, undefined, full).pages[0]!.elements.find((e) => e.type === type)!

describe('describeDeck honours withholding', () => {
  it('a withheld run reads as its marker, never its words', () => {
    const el = elementOf('text')
    expect(el.text).toBe('call {{client phone}}')
    expect(el.text).not.toContain(SECRET)
  })

  it('a withheld picture surfaces its marker instead of vanishing', () => {
    expect(elementOf('picture').text).toBe('{{id card}}')
  })

  it('a full read still withholds both', () => {
    expect(elementOf('text', true).text).toBe('call {{client phone}}')
    expect(elementOf('picture', true).text).toBe('{{id card}}')
  })
})
