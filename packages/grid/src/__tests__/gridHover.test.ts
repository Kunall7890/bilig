import { describe, expect, test } from 'vitest'
import { sameGridHoverState } from '../gridHover.js'

describe('gridHover', () => {
  test('compares hover state structurally', () => {
    expect(sameGridHoverState({ cell: [1, 2], header: null, cursor: 'cell' }, { cell: [1, 2], header: null, cursor: 'cell' })).toBe(true)
    expect(sameGridHoverState({ cell: [1, 2], header: null, cursor: 'cell' }, { cell: [1, 3], header: null, cursor: 'cell' })).toBe(false)
    expect(
      sameGridHoverState(
        { cell: null, header: { kind: 'column', index: 4 }, cursor: 'pointer' },
        { cell: null, header: { kind: 'column', index: 4 }, cursor: 'pointer' },
      ),
    ).toBe(true)
  })
})
