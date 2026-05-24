import { parseCellAddress } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS, type CellSnapshot } from '@bilig/protocol'
import type { WorkbookDeltaBatchV3 } from '@bilig/worker-transport'
import { DirtyMaskV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'

export interface ProjectedWorkbookLocalDeltaSheetIdentity {
  readonly sheetId: number
  readonly sheetOrdinal: number
}

export interface ProjectedWorkbookLocalDeltaRange {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

const LOCAL_AXIS_X_DIRTY_MASK = DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect
const LOCAL_AXIS_Y_DIRTY_MASK = DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect
export const LOCAL_CELL_VISUAL_DIRTY_MASK = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border

export function buildLocalCellSnapshotWorkbookDelta(input: {
  readonly identity: ProjectedWorkbookLocalDeltaSheetIdentity
  readonly seq: number
  readonly snapshot: CellSnapshot
  readonly dirtyMask?: number | undefined
}): WorkbookDeltaBatchV3 {
  return buildLocalCellSnapshotsWorkbookDelta({
    dirtyMask: input.dirtyMask,
    identity: input.identity,
    seq: input.seq,
    snapshots: [input.snapshot],
  })
}

export function buildLocalCellSnapshotsWorkbookDelta(input: {
  readonly dirtyMask?: number | ((snapshot: CellSnapshot) => number) | undefined
  readonly identity: ProjectedWorkbookLocalDeltaSheetIdentity
  readonly seq: number
  readonly snapshots: readonly CellSnapshot[]
}): WorkbookDeltaBatchV3 {
  const cellRanges: number[] = []
  let valueSeq = 0
  input.snapshots.forEach((snapshot) => {
    const parsed = parseCellAddress(snapshot.address, snapshot.sheetName)
    valueSeq = Math.max(valueSeq, snapshot.version, 0)
    cellRanges.push(parsed.row, parsed.row, parsed.col, parsed.col, resolveCellSnapshotDirtyMask(snapshot, input.dirtyMask))
  })
  return {
    axisSeqX: 0,
    axisSeqY: 0,
    calcSeq: valueSeq,
    dirty: {
      axisX: new Uint32Array(),
      axisY: new Uint32Array(),
      cellRanges: Uint32Array.from(cellRanges),
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

export function buildLocalRangeWorkbookDelta(input: {
  readonly identity: ProjectedWorkbookLocalDeltaSheetIdentity
  readonly seq: number
  readonly range: ProjectedWorkbookLocalDeltaRange
}): WorkbookDeltaBatchV3 {
  const rowStart = clampAxisIndex(Math.min(input.range.startRow, input.range.endRow), MAX_ROWS)
  const rowEnd = clampAxisIndex(Math.max(input.range.startRow, input.range.endRow), MAX_ROWS)
  const colStart = clampAxisIndex(Math.min(input.range.startCol, input.range.endCol), MAX_COLS)
  const colEnd = clampAxisIndex(Math.max(input.range.startCol, input.range.endCol), MAX_COLS)
  return {
    axisSeqX: 0,
    axisSeqY: 0,
    calcSeq: input.seq,
    dirty: {
      axisX: new Uint32Array(),
      axisY: new Uint32Array(),
      cellRanges: new Uint32Array([rowStart, rowEnd, colStart, colEnd, LOCAL_CELL_VISUAL_DIRTY_MASK]),
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

function resolveCellSnapshotDirtyMask(
  snapshot: CellSnapshot,
  dirtyMask: number | ((snapshot: CellSnapshot) => number) | undefined,
): number {
  const resolved = typeof dirtyMask === 'function' ? dirtyMask(snapshot) : dirtyMask
  return Number.isInteger(resolved) && resolved !== undefined && resolved >= 0 ? resolved : LOCAL_CELL_VISUAL_DIRTY_MASK
}
