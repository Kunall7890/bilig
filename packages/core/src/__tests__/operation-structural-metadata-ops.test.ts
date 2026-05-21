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

  it('invalidates ranges that reference a changed style record', () => {
    const harness = createHarness()
    const range = { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' }
    harness.workbook.upsertCellStyle({ id: 'style:accent', fill: { backgroundColor: '#ffeeaa' } })
    harness.workbook.setStyleRange(range, 'style:accent')

    const change = harness.apply({
      kind: 'upsertCellStyle',
      style: { id: 'style:accent', fill: { backgroundColor: '#34a853' } },
    })

    expect(change).toMatchObject({
      structuralInvalidation: false,
      invalidatedRanges: [range],
    })
    expect(harness.workbook.getCellStyle('style:accent')?.fill?.backgroundColor).toBe('#34a853')
    expect(harness.versionedKinds).toEqual(['upsertCellStyle'])
  })

  it('skips style invalidation when the upserted style is presentation-equivalent', () => {
    const harness = createHarness()
    const range = { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' }
    harness.workbook.upsertCellStyle({ id: 'style:accent', fill: { backgroundColor: '#00ff00' } })
    harness.workbook.setStyleRange(range, 'style:accent')

    const change = harness.apply({
      kind: 'upsertCellStyle',
      style: { id: 'style:accent', fill: { backgroundColor: '#0f0' } },
    })

    expect(change.invalidatedRanges).toEqual([])
    expect(harness.workbook.getCellStyle('style:accent')?.fill?.backgroundColor).toBe('#00ff00')
    expect(harness.versionedKinds).toEqual(['upsertCellStyle'])
  })

  it('invalidates full visible sheets when the default style changes', () => {
    const harness = createHarness()
    harness.workbook.createSheet('Sheet2')

    const change = harness.apply({
      kind: 'upsertCellStyle',
      style: { id: 'style-0', fill: { backgroundColor: '#ffffff' } },
    })

    expect(change.invalidatedRanges).toEqual([
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'XFD1048576' },
      { sheetName: 'Sheet2', startAddress: 'A1', endAddress: 'XFD1048576' },
    ])
    expect(harness.versionedKinds).toEqual(['upsertCellStyle'])
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
