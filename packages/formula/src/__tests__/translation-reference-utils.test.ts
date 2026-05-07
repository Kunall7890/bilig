import { describe, expect, it } from 'vitest'
import {
  formatAxisReference,
  formatCellReference,
  formatSheetPrefix,
  mapInterval,
  mapPointIndex,
  parseAxisReferenceParts,
  parseCellReferenceParts,
  quoteSheetNameIfNeeded,
} from '../translation-reference-utils.js'

describe('translation reference utils', () => {
  it('parses and formats absolute cell and axis references', () => {
    expect(parseCellReferenceParts('$B$12')).toEqual({ row: 11, col: 1, rowAbsolute: true, colAbsolute: true })
    expect(formatCellReference({ row: 11, col: 1, rowAbsolute: true, colAbsolute: true }, 2, 3)).toBe('$D$3')

    expect(parseAxisReferenceParts('$C', 'column')).toEqual({ index: 2, absolute: true })
    expect(formatAxisReference(true, 27, 'column')).toBe('$AB')
    expect(parseAxisReferenceParts('$4', 'row')).toEqual({ index: 3, absolute: true })
    expect(formatAxisReference(false, 4, 'row')).toBe('5')
  })

  it('quotes sheet names only when formulas require it', () => {
    expect(quoteSheetNameIfNeeded('Sheet1')).toBe('Sheet1')
    expect(quoteSheetNameIfNeeded("Owner's Sheet")).toBe("'Owner''s Sheet'")
    expect(formatSheetPrefix("Owner's Sheet")).toBe("'Owner''s Sheet'!")
    expect(formatSheetPrefix()).toBe('')
  })

  it('maps points and intervals across structural transforms', () => {
    expect(mapPointIndex(4, { kind: 'insert', axis: 'row', start: 2, count: 3 })).toBe(7)
    expect(mapPointIndex(3, { kind: 'delete', axis: 'row', start: 2, count: 3 })).toBeUndefined()
    expect(mapInterval(2, 7, { kind: 'delete', axis: 'row', start: 4, count: 2 })).toEqual({ start: 2, end: 5 })
    expect(mapInterval(2, 4, { kind: 'move', axis: 'row', start: 2, count: 3, target: 8 })).toEqual({ start: 8, end: 10 })
  })
})
