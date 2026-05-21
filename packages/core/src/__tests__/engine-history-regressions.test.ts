import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import {
  createEngineSeedSnapshot,
  exportReplaySnapshot,
  normalizeSnapshotForSemanticComparison,
  type CoreAction,
} from './engine-fuzz-helpers.js'

describe('engine history regressions', () => {
  it('does not record undo history for table-preserving structural delete no-ops', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('named-structures', 'table-preserving-delete-noop-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'table-preserving-delete-noop-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.deleteRows('Sheet1', 0, 2)
    const afterDeletingHeaderRows = engine.exportSnapshot()

    engine.deleteRows('Sheet1', 1, 1)

    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(afterDeletingHeaderRows),
    )
    expect(engine.undo()).toBe(true)
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(normalizeSnapshotForSemanticComparison(seedSnapshot))
    expect(engine.undo()).toBe(false)
  })

  it('does not synthesize generated table headers during structural undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('pivot-analytics', 'table-header-delete-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'table-header-delete-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const moveAction = {
      kind: 'move',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      target: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'D1' },
    } satisfies CoreAction
    const deleteAction = {
      kind: 'deleteColumns',
      start: 1,
      count: 1,
    } satisfies CoreAction

    engine.moveRange(moveAction.source, moveAction.target)
    const expectedAfterMove = await exportReplaySnapshot(seedSnapshot, [moveAction])
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(expectedAfterMove),
    )
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Empty })

    engine.deleteColumns('Sheet1', deleteAction.start, deleteAction.count)
    expect(engine.undo()).toBe(true)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Empty })
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(expectedAfterMove),
    )
  })
})
