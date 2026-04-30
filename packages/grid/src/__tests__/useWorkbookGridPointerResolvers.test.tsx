// @vitest-environment jsdom
import { act, useMemo, useRef } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { createGridSelection } from '../gridSelection.js'
import { useWorkbookGridPointerResolvers } from '../useWorkbookGridPointerResolvers.js'

describe('useWorkbookGridPointerResolvers', () => {
  afterEach(() => {
    document.body.innerHTML = ''
  })

  test('treats the live camera geometry as authoritative over stale visible-region fallback math', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const gridMetrics = getGridMetrics()
    const staleVisibleRegion = {
      freezeCols: 0,
      freezeRows: 0,
      range: { x: 0, y: 0, width: 12, height: 24 },
      tx: 0,
      ty: 0,
    }
    let latestResolvers: ReturnType<typeof useWorkbookGridPointerResolvers> | null = null

    function Harness() {
      const hostRef = useRef<HTMLDivElement | null>(null)
      const liveGeometry = useMemo(
        () =>
          createGridGeometrySnapshotFromAxes({
            columns: createGridAxisWorldIndex({ axisLength: 100, defaultSize: gridMetrics.columnWidth }),
            dpr: 1,
            gridMetrics,
            hostHeight: 180,
            hostWidth: 220,
            rows: createGridAxisWorldIndex({ axisLength: 100, defaultSize: gridMetrics.rowHeight }),
            scrollLeft: 0,
            scrollTop: 0,
            sheetName: 'Sheet1',
            updatedAt: 100,
          }),
        [],
      )
      latestResolvers = useWorkbookGridPointerResolvers({
        columnWidths: {},
        getCellScreenBounds: () => undefined,
        getGeometrySnapshot: () => liveGeometry,
        getVisibleRegion: () => staleVisibleRegion,
        gridMetrics,
        gridSelection: createGridSelection(0, 0),
        hostRef,
        rowHeights: {},
        selectedCell: { col: 0, row: 0 },
      })

      return (
        <div
          ref={(node) => {
            hostRef.current = node
            if (node) {
              node.getBoundingClientRect = () =>
                ({
                  bottom: 240,
                  height: 240,
                  left: 0,
                  right: 1000,
                  top: 0,
                  width: 1000,
                  x: 0,
                  y: 0,
                  toJSON: () => ({}),
                }) as DOMRect
            }
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestResolvers?.resolveHeaderSelectionAtPointer(300, 10)).toBeNull()
    expect(latestResolvers?.resolvePointerCell(300, 80)).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })
})
