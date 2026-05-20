import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'
import { diffWorkbookSemanticSnapshots, workbookSemanticSnapshotsEqual } from '../semantics/index.js'

describe('engine semantic invariants', () => {
  it('keeps edit, snapshot restore, and full recalc semantically equivalent', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'semantic-invariants', replicaId: 'semantic-primary' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [
      [1, 10],
      [2, 20],
      [3, 30],
    ])
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'C2', 'SUM(B1:B3)')
    engine.insertRows('Sheet1', 1, 1)
    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellValue('Sheet1', 'B2', 40)

    const beforeRestore = engine.exportSnapshot()

    const restored = new SpreadsheetEngine({ workbookName: 'semantic-invariants', replicaId: 'semantic-restored' })
    await restored.ready()
    restored.importSnapshot(beforeRestore)
    restored.recalculateNow()

    const afterRestore = restored.exportSnapshot()
    expect(diffWorkbookSemanticSnapshots(beforeRestore, afterRestore)).toEqual([])
    expect(workbookSemanticSnapshotsEqual(beforeRestore, afterRestore)).toBe(true)
  })
})
