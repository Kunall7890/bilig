import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'
import {
  applyActionAndCaptureResult,
  createEngineSeedSnapshot,
  exportReplaySnapshot,
  normalizeSnapshotForSemanticComparison,
  type CoreAction,
} from './engine-fuzz-helpers.js'

function undoAll(engine: SpreadsheetEngine): number {
  let count = 0
  while (engine.undo()) {
    count += 1
  }
  return count
}

describe('engine history fuzz regressions', () => {
  it('counts undoable structural deletes even when the exported snapshot is unchanged', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('named-structures', 'history-undoable-invisible-structural-delete')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'history-undoable-invisible-structural-delete',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const actions: CoreAction[] = [
      { kind: 'deleteColumns', start: 0, count: 2 },
      { kind: 'deleteRows', start: 0, count: 1 },
      { kind: 'deleteRows', start: 0, count: 1 },
    ]
    const accepted = actions.map((action) => applyActionAndCaptureResult(engine, action).accepted)

    expect(accepted).toEqual([true, true, true])
    expect(undoAll(engine)).toBe(3)
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(normalizeSnapshotForSemanticComparison(seedSnapshot))
  })

  it('keeps model history aligned when undoing an invisible clear operation', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('pivot-analytics', 'history-invisible-clear-undo')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'history-invisible-clear-undo',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const applied: CoreAction[] = []
    const applyAccepted = (action: CoreAction) => {
      const result = applyActionAndCaptureResult(engine, action)
      if (result.accepted) {
        applied.push(action)
      }
    }

    applyAccepted({ kind: 'deleteRows', start: 0, count: 1 })
    applyAccepted({ kind: 'clear', range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' } })
    expect(applied).toHaveLength(2)

    expect(engine.undo()).toBe(true)
    applied.pop()

    const expectedSnapshot = await exportReplaySnapshot(seedSnapshot, applied)
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(expectedSnapshot),
    )
  })

  it('preserves sheet-qualified invalid defined-name refs when undoing column inserts', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('named-structures', 'history-invalid-defined-name-column-insert-undo')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'history-invalid-defined-name-column-insert-undo',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const applied: CoreAction[] = [
      { kind: 'deleteColumns', start: 0, count: 2 },
      { kind: 'insertColumns', start: 0, count: 1 },
    ]
    for (const action of applied) {
      const result = applyActionAndCaptureResult(engine, action)
      expect(result.accepted).toBe(true)
    }

    expect(engine.undo()).toBe(true)
    applied.pop()

    const expectedSnapshot = await exportReplaySnapshot(seedSnapshot, applied)
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: { kind: 'formula', formula: '=Sheet1!#REF!' },
    })
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(expectedSnapshot),
    )
  })
})
