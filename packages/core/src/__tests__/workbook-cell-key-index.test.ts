import { describe, expect, it } from 'vitest'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { createCellKeyIndexMap, makeCellKey, makeLogicalCellKey } from '../workbook-cell-key-index.js'

describe('workbook cell key index', () => {
  it('resolves physical cell keys lazily through the current sheet position index', () => {
    const index = createCellKeyIndexMap((sheetId, row, col) => (sheetId === 2 && row === 4 && col === 6 ? 42 : undefined))

    expect(index.get(makeCellKey(2, 4, 6))).toBe(42)
    expect(index.has(makeCellKey(2, 4, 6))).toBe(true)
    expect(index.get(makeCellKey(2, 4, 7))).toBeUndefined()
    expect(index.has(makeCellKey(2, 4, 7))).toBe(false)
  })

  it('falls back to explicit map entries for invalid physical cell keys', () => {
    const index = createCellKeyIndexMap(() => undefined)

    index.set(-1, 17)

    expect(index.get(-1)).toBe(17)
    expect(index.has(-1)).toBe(true)
    expect(index.get(makeCellKey(0, 0, 0))).toBeUndefined()
    expect(index.get(makeCellKey(1, MAX_ROWS, 0))).toBeUndefined()
    expect(index.get(makeCellKey(1, 0, MAX_COLS))).toBeUndefined()
  })

  it('uses a stable logical cell key delimiter', () => {
    expect(makeLogicalCellKey(3, 'row-1', 'column-2')).toBe('3\trow-1\tcolumn-2')
  })
})
