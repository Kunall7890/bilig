import { describe, expect, it } from 'vitest'
import { spillDependencyKey } from '../engine-metadata-utils.js'
import {
  createOperationDerivedOpApplier,
  type OperationDerivedOp,
  type OperationDerivedWorkbookAccess,
} from '../engine/services/operation-derived-op-helpers.js'
import { createReplicaState, type OpOrder } from '../replica-state.js'
import type { WorkbookPivotRecord } from '../workbook-store.js'

function createHarness(input?: { readonly trackedSpillDependents?: readonly number[] }) {
  const setSpills: { sheetName: string; address: string; rows: number; cols: number }[] = []
  const deletedSpills: { sheetName: string; address: string }[] = []
  const deletedPivots: { sheetName: string; address: string }[] = []
  const versioned: { op: OperationDerivedOp; order: OpOrder }[] = []
  let pivot: WorkbookPivotRecord | undefined

  const workbook = {
    setSpill(sheetName, address, rows, cols) {
      setSpills.push({ sheetName, address, rows, cols })
      return { sheetName, address, rows, cols }
    },
    deleteSpill(sheetName, address) {
      deletedSpills.push({ sheetName, address })
      return true
    },
    setPivot(record) {
      pivot = record
      return record
    },
    getPivot(sheetName, address) {
      return pivot?.sheetName === sheetName && pivot.address === address ? pivot : undefined
    },
    deletePivot(sheetName, address) {
      deletedPivots.push({ sheetName, address })
      pivot = undefined
      return true
    },
  } satisfies OperationDerivedWorkbookAccess

  const materializedPivots: WorkbookPivotRecord[] = []
  const clearedPivots: WorkbookPivotRecord[] = []
  const rebinds: { candidates: readonly number[]; formulaChangedCount: number }[] = []
  const applier = createOperationDerivedOpApplier({
    state: {
      workbook,
      replicaState: createReplicaState('replica-a'),
    },
    reverseSpillEdges: new Map(
      input?.trackedSpillDependents ? [[spillDependencyKey('Sheet1', 'B2'), new Set(input.trackedSpillDependents)]] : [],
    ),
    setEntityVersionForOp(op, order) {
      versioned.push({ op, order })
    },
    materializePivot(record) {
      materializedPivots.push(record)
      return [31, 32]
    },
    clearOwnedPivot(record) {
      clearedPivots.push(record)
      return [41]
    },
    rebindFormulaCells(candidates, formulaChangedCount) {
      rebinds.push({ candidates, formulaChangedCount })
      return formulaChangedCount + candidates.length
    },
  })

  return {
    applier,
    clearedPivots,
    deletedPivots,
    deletedSpills,
    materializedPivots,
    rebinds,
    setSpills,
    versioned,
  }
}

describe('operation derived op helpers', () => {
  it('applies spill range ops and rebinds tracked spill dependents', () => {
    const harness = createHarness({ trackedSpillDependents: [7, 3, 7] })
    const op = { kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'B2', rows: 2, cols: 3 } satisfies OperationDerivedOp

    const changed = harness.applier.applyDerivedOpNow(op)

    expect(changed).toEqual([7, 3])
    expect(harness.setSpills).toEqual([{ sheetName: 'Sheet1', address: 'B2', rows: 2, cols: 3 }])
    expect(harness.rebinds).toEqual([{ candidates: [7, 3], formulaChangedCount: 0 }])
    expect(harness.versioned).toHaveLength(1)
    const firstVersion = harness.versioned[0]
    expect(firstVersion?.op).toBe(op)
    expect(firstVersion?.order).toMatchObject({ counter: 1, replicaId: 'replica-a', opIndex: 0 })
  })

  it('materializes and clears pivot outputs around pivot derived ops', () => {
    const harness = createHarness()
    const upsert = {
      kind: 'upsertPivotTable',
      name: 'Pivot1',
      sheetName: 'Sheet1',
      address: 'D4',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B10' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Amount', aggregate: 'sum' }],
      rows: 4,
      cols: 2,
    } satisfies OperationDerivedOp

    expect(harness.applier.applyDerivedOpNow(upsert)).toEqual([31, 32])
    expect(harness.materializedPivots).toEqual([
      {
        name: 'Pivot1',
        sheetName: 'Sheet1',
        address: 'D4',
        source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B10' },
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Amount', aggregate: 'sum' }],
        rows: 4,
        cols: 2,
      },
    ])

    const deleteOp = { kind: 'deletePivotTable', sheetName: 'Sheet1', address: 'D4' } satisfies OperationDerivedOp
    expect(harness.applier.applyDerivedOpNow(deleteOp)).toEqual([41])
    expect(harness.clearedPivots).toHaveLength(1)
    expect(harness.deletedPivots).toEqual([{ sheetName: 'Sheet1', address: 'D4' }])
    expect(harness.versioned.map((entry) => entry.op.kind)).toEqual(['upsertPivotTable', 'deletePivotTable'])
  })
})
