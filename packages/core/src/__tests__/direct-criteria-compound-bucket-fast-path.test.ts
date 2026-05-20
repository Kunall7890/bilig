import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('direct criteria compound bucket fast path', () => {
  it('keeps exact compound aggregate edits on the direct path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-criteria-compound-bucket-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'E1', 'A')
    engine.setCellValue('Sheet1', 'F1', 'target')

    for (let row = 2; row <= 257; row += 1) {
      const value = row - 1
      engine.setCellValue('Sheet1', `A${row}`, value % 2 === 0 ? 'A' : 'B')
      engine.setCellValue('Sheet1', `B${row}`, value % 16 === 0 ? 'target' : 'other')
      engine.setCellValue('Sheet1', `C${row}`, value)
    }

    engine.setCellFormula('Sheet1', 'G1', 'COUNTIFS(A2:A257,E1,B2:B257,F1)')
    engine.setCellFormula('Sheet1', 'G2', 'SUMIFS(C2:C257,A2:A257,E1,B2:B257,F1)')
    engine.setCellFormula('Sheet1', 'G3', 'AVERAGEIFS(C2:C257,A2:A257,E1,B2:B257,F1)')
    engine.setCellFormula('Sheet1', 'G4', 'MINIFS(C2:C257,A2:A257,E1,B2:B257,F1)')
    engine.setCellFormula('Sheet1', 'G5', 'MAXIFS(C2:C257,A2:A257,E1,B2:B257,F1)')

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'C17', 1600)

    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 16 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 3760 })
    expect(engine.getCellValue('Sheet1', 'G3')).toEqual({ tag: ValueTag.Number, value: 235 })
    expect(engine.getCellValue('Sheet1', 'G4')).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(engine.getCellValue('Sheet1', 'G5')).toEqual({ tag: ValueTag.Number, value: 1600 })
    expect(engine.getPerformanceCounters().nativeDirectCriteriaPredicateAggregateEvaluations).toBe(0)
    expect(engine.getPerformanceCounters().nativeDirectCriteriaAggregateEvaluations).toBe(0)
    expect(engine.getPerformanceCounters().directCriteriaMatchCacheHits).toBe(0)
    expect(engine.getPerformanceCounters().formulasBound).toBe(0)
    expect(engine.getPerformanceCounters().topoRebuilds).toBe(0)
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().columnOwnerBuilds).toBeLessThanOrEqual(1)
  })
})
