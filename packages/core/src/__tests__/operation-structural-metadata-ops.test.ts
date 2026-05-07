import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../workbook-store.js'
import type { OpOrder } from '../replica-state.js'
import {
  applyOperationStructuralMetadataOp,
  type OperationStructuralMetadataOp,
} from '../engine/services/operation-structural-metadata-ops.js'

const order: OpOrder = {
  counter: 1,
  replicaId: 'test',
  batchId: 'test:1',
  opIndex: 0,
}

function createHarness() {
  const workbook = new WorkbookStore()
  workbook.createSheet('Sheet1')
  const versionedKinds: string[] = []
  return {
    workbook,
    versionedKinds,
    apply(op: OperationStructuralMetadataOp, source: 'local' | 'restore' = 'local') {
      return applyOperationStructuralMetadataOp({
        workbook,
        op,
        order,
        source,
        setEntityVersionForOp: (versionedOp) => {
          versionedKinds.push(versionedOp.kind)
        },
      })
    },
  }
}

describe('operation structural metadata ops', () => {
  it('applies row metadata and reports row invalidation without full structural invalidation', () => {
    const harness = createHarness()

    const change = harness.apply({
      kind: 'updateRowMetadata',
      sheetName: 'Sheet1',
      start: 2,
      count: 3,
      size: 24,
      hidden: null,
    })

    expect(change.structuralInvalidation).toBe(false)
    expect(change.invalidatedRows).toEqual([{ sheetName: 'Sheet1', startIndex: 2, endIndex: 4 }])
    expect(change.invalidatedColumns).toEqual([])
    expect(harness.workbook.getRowMetadata('Sheet1', 2, 3)?.size).toBe(24)
    expect(harness.versionedKinds).toEqual(['updateRowMetadata'])
  })

  it('applies range metadata and returns the exact range invalidation surface', () => {
    const harness = createHarness()
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }
    harness.workbook.upsertCellStyle({ id: 'style:accent', fill: { backgroundColor: '#ffeeaa' } })

    const styleChange = harness.apply({ kind: 'setStyleRange', range, styleId: 'style:accent' })
    const mergeChange = harness.apply({ kind: 'mergeCells', range })

    expect(styleChange).toMatchObject({
      structuralInvalidation: false,
      invalidatedRanges: [range],
    })
    expect(mergeChange).toMatchObject({
      structuralInvalidation: true,
      invalidatedRanges: [range],
    })
    expect(harness.workbook.listStyleRanges('Sheet1')).toHaveLength(1)
    expect(harness.workbook.listMergeRanges('Sheet1')).toHaveLength(1)
    expect(harness.versionedKinds).toEqual(['setStyleRange', 'mergeCells'])
  })

  it('applies structural sheet metadata and reports a full structural invalidation', () => {
    const harness = createHarness()

    const change = harness.apply({ kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 2 })

    expect(change.structuralInvalidation).toBe(true)
    expect(change.invalidatedRanges).toEqual([])
    expect(harness.workbook.getFreezePane('Sheet1')).toMatchObject({ rows: 1, cols: 2 })
    expect(harness.versionedKinds).toEqual(['setFreezePane'])
  })
})
