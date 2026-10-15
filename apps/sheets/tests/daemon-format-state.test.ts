import { BooleanNumber, BorderStyleTypes, type IStyleData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  borderStateOf,
  cellFormatStateOf,
  DEFAULT_FORMAT_CATEGORIES,
  FORMAT_CATEGORIES,
  FORMAT_CATEGORY_KEYS,
} from '../src/renderer/daemon/format-state'

/** a stand-in range carrying `style` and a fixed number format */
function range(style: Record<string, unknown>, pattern = 'General') {
  return {
    getCellStyleData: () => style as unknown as IStyleData,
    getNumberFormat: () => pattern,
  }
}

describe('borderStateOf', () => {
  it('reads each bordered edge as a line style and its plain-hex color', () => {
    const borders = borderStateOf({
      bd: {
        t: { s: BorderStyleTypes.THIN, cl: { rgb: '#c00000' } },
        b: { s: BorderStyleTypes.DOUBLE, cl: { rgb: 'rgb(0, 128, 255)' } },
      },
    } as unknown as IStyleData)
    expect(borders).toEqual({
      top: { style: 'thin', color: '#C00000' },
      bottom: { style: 'double', color: '#0080FF' },
    })
  })

  it('omits an edge with no color, and edges that are not bordered', () => {
    // a line style with no color still reads as a border (color left off)
    expect(
      borderStateOf({ bd: { l: { s: BorderStyleTypes.MEDIUM } } } as unknown as IStyleData),
    ).toEqual({ left: { style: 'medium' } })
    // s = none is not a border; a style-free cell has none either
    expect(borderStateOf({ bd: { l: { s: 0 } } } as unknown as IStyleData)).toBeUndefined()
    expect(borderStateOf({} as IStyleData)).toBeUndefined()
  })
})

describe('cellFormatStateOf', () => {
  it('folds the per-edge borders into the cell format state', () => {
    const format = cellFormatStateOf(
      range({ bl: BooleanNumber.TRUE, bd: { r: { s: BorderStyleTypes.MEDIUM, cl: { rgb: '#000000' } } } }),
    )
    expect(format).toEqual({
      bold: true,
      border: { right: { style: 'medium', color: '#000000' } },
    })
  })

  it('leaves the border off a cell that has none', () => {
    expect(cellFormatStateOf(range({}))).toEqual({})
  })

  it('emits rotation in the DSL vocabulary and indent as steps', () => {
    expect(cellFormatStateOf(range({ tr: { a: -30 } }))).toEqual({ textRotation: -30 })
    expect(cellFormatStateOf(range({ tr: { v: BooleanNumber.TRUE } }))).toEqual({
      textRotation: 'vertical',
    })
    expect(cellFormatStateOf(range({ pd: { l: 16 } }))).toEqual({ indent: 2 })
  })

  it('scopes the read to the named categories', () => {
    const cell = range({
      bl: BooleanNumber.TRUE,
      bg: { rgb: '#ff0000' },
      bd: { t: { s: BorderStyleTypes.THIN } },
      pd: { l: 8 },
    })
    expect(cellFormatStateOf(cell, ['font'])).toEqual({ bold: true })
    expect(cellFormatStateOf(cell, ['fill'])).toEqual({ fillColor: '#FF0000' })
    expect(cellFormatStateOf(cell, ['font', 'border'])).toEqual({
      bold: true,
      border: { top: { style: 'thin' } },
    })
    expect(cellFormatStateOf(cell, ['all'])).toEqual(cellFormatStateOf(cell))
  })
})

describe('FORMAT_CATEGORY_KEYS', () => {
  it('partitions the read properties across the five groups, with `all` their union', () => {
    const groups = FORMAT_CATEGORIES.filter((category) => category !== 'all').flatMap(
      (category) => FORMAT_CATEGORY_KEYS[category],
    )
    // the groups are disjoint and together cover every property `all` reports
    expect(new Set(groups).size).toBe(groups.length)
    expect([...groups].sort()).toEqual([...FORMAT_CATEGORY_KEYS.all].sort())
  })

  it('defaults to the common groups, leaving the per-edge border opt-in', () => {
    expect(DEFAULT_FORMAT_CATEGORIES).toEqual(['font', 'fill', 'alignment', 'number'])
    expect(DEFAULT_FORMAT_CATEGORIES).not.toContain('border')
    expect(DEFAULT_FORMAT_CATEGORIES).not.toContain('all')
  })
})
