import { describe, expect, it } from 'vitest'
import { relayResult } from '../src/server'

/**
 * Every office_gui_* call relays the app's result through `relayResult`, like
 * the CLI family maps its own: text content and isError. On top of that, a
 * `structuredContent` the app sent is forwarded rather than dropped.
 */

describe('relayResult', () => {
  it('forwards structuredContent alongside the text content', () => {
    const relayed = relayResult({
      content: [{ type: 'text', text: '{"a":1}' }],
      structuredContent: { a: 1 },
    })
    expect(relayed.content).toEqual([{ type: 'text', text: '{"a":1}' }])
    expect(relayed.structuredContent).toEqual({ a: 1 })
    expect(relayed.isError).toBeUndefined()
  })

  it('omits structuredContent when the app sent none', () => {
    const relayed = relayResult({ content: [{ type: 'text', text: 'hi' }] })
    expect(relayed).not.toHaveProperty('structuredContent')
  })

  it('carries an isError result through', () => {
    const relayed = relayResult({ content: [{ type: 'text', text: 'boom' }], isError: true })
    expect(relayed.isError).toBe(true)
  })
})
