import { describe, expect, it } from 'vitest'
import type { GridCameraSnapshotV2 } from '../gridGeometry.js'
import { clampGridScreenPointToDataPanes } from '../gridPointer.js'

function cameraWithPanes(panes: GridCameraSnapshotV2['panes']): GridCameraSnapshotV2 {
  return {
    axisVersionX: 0,
    axisVersionY: 0,
    bodyScrollX: 0,
    bodyScrollY: 0,
    bodyViewportHeight: 120,
    bodyViewportWidth: 200,
    bodyWorldX: 0,
    bodyWorldY: 0,
    columnAnchor: { index: 0, intraOffset: 0, offset: 0, size: 104 },
    dpr: 1,
    frozenColumnCount: 0,
    frozenHeight: 0,
    frozenRowCount: 0,
    frozenWidth: 0,
    panes,
    rowAnchor: { index: 0, intraOffset: 0, offset: 0, size: 22 },
    scrollLeft: 0,
    scrollTop: 0,
    seq: 1,
    sheetName: 'Sheet1',
    updatedAt: 0,
    velocityX: 0,
    velocityY: 0,
  }
}

describe('clampGridScreenPointToDataPanes', () => {
  it('clamps fill-handle drag points back into the grid data pane', () => {
    const camera = cameraWithPanes([
      { kind: 'corner-header', frame: { x: 0, y: 0, width: 40, height: 24 }, scrollAxes: { x: false, y: false } },
      { kind: 'column-header-body', frame: { x: 40, y: 0, width: 200, height: 24 }, scrollAxes: { x: true, y: false } },
      { kind: 'row-header-body', frame: { x: 0, y: 24, width: 40, height: 120 }, scrollAxes: { x: false, y: true } },
      { kind: 'body', frame: { x: 40, y: 24, width: 200, height: 120 }, scrollAxes: { x: true, y: true } },
    ])

    expect(clampGridScreenPointToDataPanes(camera, { x: 260, y: 180 })).toEqual({ x: 239, y: 143 })
    expect(clampGridScreenPointToDataPanes(camera, { x: 10, y: 4 })).toEqual({ x: 40, y: 24 })
  })

  it('uses the union of frozen and scrolling data panes', () => {
    const camera = cameraWithPanes([
      { kind: 'frozen-cells', frame: { x: 40, y: 24, width: 80, height: 44 }, scrollAxes: { x: false, y: false } },
      { kind: 'frozen-rows', frame: { x: 120, y: 24, width: 240, height: 44 }, scrollAxes: { x: true, y: false } },
      { kind: 'frozen-columns', frame: { x: 40, y: 68, width: 80, height: 220 }, scrollAxes: { x: false, y: true } },
      { kind: 'body', frame: { x: 120, y: 68, width: 240, height: 220 }, scrollAxes: { x: true, y: true } },
    ])

    expect(clampGridScreenPointToDataPanes(camera, { x: 999, y: 999 })).toEqual({ x: 359, y: 287 })
  })

  it('returns null when no data pane can receive cell hits', () => {
    const camera = cameraWithPanes([
      { kind: 'corner-header', frame: { x: 0, y: 0, width: 40, height: 24 }, scrollAxes: { x: false, y: false } },
    ])

    expect(clampGridScreenPointToDataPanes(camera, { x: 12, y: 12 })).toBeNull()
  })
})
