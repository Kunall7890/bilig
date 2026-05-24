import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { DirtyMaskV3 } from '../../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import {
  buildLocalAxisWorkbookDelta,
  buildLocalCellSnapshotWorkbookDelta,
  buildLocalCellSnapshotsWorkbookDelta,
  buildLocalRangeWorkbookDelta,
  LOCAL_CELL_CONTENT_DIRTY_MASK,
  LOCAL_CELL_VISUAL_DIRTY_MASK,
} from '../projected-workbook-local-delta.js'

const identity = {
  sheetId: 7,
  sheetOrdinal: 3,
}
describe('projected workbook local delta builders', () => {
  it('builds cell deltas from the accepted snapshot version and visual shape', () => {
    const batch = buildLocalCellSnapshotWorkbookDelta({
      identity,
      seq: 44,
      snapshot: {
        address: 'B2',
        flags: 0,
        sheetName: 'Sheet1',
        styleId: 'accent',
        value: { tag: ValueTag.Number, value: 17 },
        version: 12,
      },
    })

    expect(batch).toMatchObject({
      calcSeq: 12,
      seq: 44,
      sheetId: 7,
      sheetOrdinal: 3,
      source: 'localOptimistic',
      styleSeq: 12,
      valueSeq: 12,
    })
    expect(batch.dirty.cellRanges).toEqual(new Uint32Array([1, 1, 1, 1, LOCAL_CELL_VISUAL_DIRTY_MASK]))
  })

  it('marks plain local cell writes as full visual damage so moved text cannot retain stale fills', () => {
    const batch = buildLocalCellSnapshotWorkbookDelta({
      identity,
      seq: 47,
      snapshot: {
        address: 'C4',
        flags: 0,
        sheetName: 'Sheet1',
        value: { tag: ValueTag.String, stringId: 9, value: 'moved text' },
        version: 13,
      },
    })

    expect(batch.dirty.cellRanges).toEqual(new Uint32Array([3, 3, 2, 2, LOCAL_CELL_VISUAL_DIRTY_MASK]))
  })

  it('builds multi-cell deltas for optimistic trust flag clears', () => {
    const batch = buildLocalCellSnapshotsWorkbookDelta({
      identity,
      seq: 48,
      snapshots: [
        {
          address: 'B2',
          flags: 0,
          sheetName: 'Sheet1',
          value: { tag: ValueTag.Number, value: 17 },
          version: 12,
        },
        {
          address: 'D4',
          flags: 0,
          sheetName: 'Sheet1',
          value: { tag: ValueTag.String, stringId: 9, value: 'trusted' },
          version: 15,
        },
      ],
    })

    expect(batch).toMatchObject({
      calcSeq: 15,
      seq: 48,
      sheetId: 7,
      sheetOrdinal: 3,
      source: 'localOptimistic',
      styleSeq: 15,
      valueSeq: 15,
    })
    expect(batch.dirty.cellRanges).toEqual(
      new Uint32Array([1, 1, 1, 1, LOCAL_CELL_VISUAL_DIRTY_MASK, 3, 3, 3, 3, LOCAL_CELL_VISUAL_DIRTY_MASK]),
    )
  })

  it('honors precise dirty masks for local style-only updates', () => {
    const batch = buildLocalCellSnapshotWorkbookDelta({
      dirtyMask: DirtyMaskV3.Style,
      identity,
      seq: 51,
      snapshot: {
        address: 'B2',
        flags: 0,
        sheetName: 'Sheet1',
        styleId: 'style-local-bold-empty',
        value: { tag: ValueTag.Empty },
        version: 16,
      },
    })

    expect(batch.dirty.cellRanges).toEqual(new Uint32Array([1, 1, 1, 1, DirtyMaskV3.Style]))
  })

  it('honors precise dirty masks for local content-only updates', () => {
    const batch = buildLocalCellSnapshotWorkbookDelta({
      dirtyMask: LOCAL_CELL_CONTENT_DIRTY_MASK,
      identity,
      seq: 52,
      snapshot: {
        address: 'B2',
        flags: 0,
        sheetName: 'Sheet1',
        value: { tag: ValueTag.String, value: 'typed' },
        version: 17,
      },
    })

    expect(batch.dirty.cellRanges).toEqual(new Uint32Array([1, 1, 1, 1, DirtyMaskV3.Value | DirtyMaskV3.Text]))
  })

  it('builds coarse range deltas for unmaterialized optimistic style overlays', () => {
    const batch = buildLocalRangeWorkbookDelta({
      identity,
      range: {
        startRow: 4,
        endRow: 899,
        startCol: 3,
        endCol: 5,
      },
      seq: 49,
    })

    expect(batch).toMatchObject({
      calcSeq: 49,
      seq: 49,
      sheetId: 7,
      sheetOrdinal: 3,
      source: 'localOptimistic',
      styleSeq: 49,
      valueSeq: 49,
    })
    expect(batch.dirty.axisX).toEqual(new Uint32Array())
    expect(batch.dirty.axisY).toEqual(new Uint32Array())
    expect(batch.dirty.cellRanges).toEqual(new Uint32Array([4, 899, 3, 5, LOCAL_CELL_VISUAL_DIRTY_MASK]))
  })

  it('builds axis deltas with isolated axis damage and clamped indices', () => {
    const columnBatch = buildLocalAxisWorkbookDelta({
      axis: 'column',
      identity,
      index: 2,
      seq: 45,
    })
    const rowBatch = buildLocalAxisWorkbookDelta({
      axis: 'row',
      identity,
      index: 4,
      seq: 46,
    })

    expect(columnBatch).toMatchObject({
      axisSeqX: 45,
      axisSeqY: 0,
      calcSeq: 45,
      source: 'localOptimistic',
    })
    expect(columnBatch.dirty.axisX).toEqual(new Uint32Array([2, 2, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]))
    expect(columnBatch.dirty.axisY).toEqual(new Uint32Array())
    expect(rowBatch).toMatchObject({
      axisSeqX: 0,
      axisSeqY: 46,
      calcSeq: 46,
      source: 'localOptimistic',
    })
    expect(rowBatch.dirty.axisX).toEqual(new Uint32Array())
    expect(rowBatch.dirty.axisY).toEqual(new Uint32Array([4, 4, DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]))
  })
})
