import { describe, expect, it } from 'vitest'
import { createFormulaBindingSheetIndex } from '../engine/services/formula-binding-sheet-index.js'

describe('formula binding sheet index', () => {
  it('tracks formula owners and qualified sheet references without duplicate cells', () => {
    const index = createFormulaBindingSheetIndex()

    index.trackFormula(7, 'Sheet1', { deps: ['A1', "'Q1 Report'!B2", "'Owner''s Sheet'!C3", "'Q1 Report'!D4"] })
    index.trackFormula(7, 'Sheet1', { deps: ["'Q1 Report'!B2"] })
    index.trackFormula(9, 'Sheet2', { deps: ['Sheet1!A1'] })

    expect(index.collectOwnedBySheet('Sheet1')).toEqual([7])
    expect(index.collectReferencingSheet('Q1 Report')).toEqual([7])
    expect(index.collectReferencingSheet("Owner's Sheet")).toEqual([7])
    expect(index.collectReferencingSheet('Sheet1')).toEqual([9])
  })

  it('untracks formulas and removes empty sheet entries', () => {
    const index = createFormulaBindingSheetIndex()
    index.trackFormula(7, 'Sheet1', { deps: ['Sheet2!A1'] })
    index.untrackFormula(7, 'Sheet1', { deps: ['Sheet2!A1'] })

    expect(index.collectOwnedBySheet('Sheet1')).toEqual([])
    expect(index.collectReferencingSheet('Sheet2')).toEqual([])
  })

  it('moves owner and reference buckets when a sheet is renamed', () => {
    const index = createFormulaBindingSheetIndex()
    index.trackFormula(1, 'Old', { deps: ['Target!A1'] })
    index.trackFormula(2, 'Other', { deps: ['Old!A1'] })
    index.trackFormula(3, 'New', { deps: ['Old!B1'] })

    const moved = index.moveSheetName('Old', 'New')

    expect([...moved.owners]).toEqual([1])
    expect([...moved.references]).toEqual([2, 3])
    expect(index.collectOwnedBySheet('Old')).toEqual([])
    expect(index.collectOwnedBySheet('New')).toEqual([3, 1])
    expect(index.collectReferencingSheet('Old')).toEqual([])
    expect(index.collectReferencingSheet('New')).toEqual([2, 3])
  })
})
