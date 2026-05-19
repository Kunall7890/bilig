import { parseCellAddress } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS, type CellSnapshot } from '@bilig/protocol'
import type { WorkbookDeltaBatchV3 } from '@bilig/worker-transport'
import { DirtyMaskV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'

export interface ProjectedWorkbookLocalDeltaSheetIdentity {
  readonly sheetId: number
  readonly sheetOrdinal: number
}

const LOCAL_AXIS_X_DIRTY_MASK = DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect
const LOCAL_AXIS_Y_DIRTY_MASK = DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect
export const LOCAL_CELL_TEXT_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Text
export const LOCAL_CELL_VISUAL_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border

export function buildLocalCellSnapshotWorkbookDelta(input: {
  readonly identity: ProjectedWorkbookLocalDeltaSheetIdentity
  readonly dirtyMask?: number | undefined
  readonly seq: number
  readonly snapshot: CellSnapshot
}): WorkbookDeltaBatchV3 {
  const parsed = parseCellAddress(input.snapshot.address, input.snapshot.sheetName)
  const valueSeq = Math.max(0, input.snapshot.version)
  return {
    axisSeqX: 0,
    axisSeqY: 0,
    calcSeq: valueSeq,
    dirty: {
      axisX: new Uint32Array(),
      axisY: new Uint32Array(),
      cellRanges: new Uint32Array([parsed.row, parsed.row, parsed.col, parsed.col, input.dirtyMask ?? LOCAL_CELL_VISUAL_DIRTY_MASK]),
    },
    freezeSeq: 0,
    magic: 'bilig.workbook.delta.v3',
    seq: input.seq,
    sheetId: input.identity.sheetId,
    sheetOrdinal: input.identity.sheetOrdinal,
    source: 'localOptimistic',
    styleSeq: valueSeq,
    valueSeq,
    version: 1,
  }
}

export function buildLocalAxisWorkbookDelta(input: {
  readonly axis: 'column' | 'row'
  readonly identity: ProjectedWorkbookLocalDeltaSheetIdentity
  readonly index: number
  readonly seq: number
}): WorkbookDeltaBatchV3 {
  const axisIndex = input.axis === 'column' ? clampAxisIndex(input.index, MAX_COLS) : clampAxisIndex(input.index, MAX_ROWS)
  return {
    axisSeqX: input.axis === 'column' ? input.seq : 0,
    axisSeqY: input.axis === 'row' ? input.seq : 0,
    calcSeq: input.seq,
    dirty: {
      axisX: input.axis === 'column' ? new Uint32Array([axisIndex, axisIndex, LOCAL_AXIS_X_DIRTY_MASK]) : new Uint32Array(),
      axisY: input.axis === 'row' ? new Uint32Array([axisIndex, axisIndex, LOCAL_AXIS_Y_DIRTY_MASK]) : new Uint32Array(),
      cellRanges: new Uint32Array(),
    },
    freezeSeq: 0,
    magic: 'bilig.workbook.delta.v3',
    seq: input.seq,
    sheetId: input.identity.sheetId,
    sheetOrdinal: input.identity.sheetOrdinal,
    source: 'localOptimistic',
    styleSeq: input.seq,
    valueSeq: input.seq,
    version: 1,
  }
}

function clampAxisIndex(index: number, axisLength: number): number {
  if (!Number.isFinite(index)) {
    return 0
  }
  return Math.max(0, Math.min(axisLength - 1, Math.trunc(index)))
}
