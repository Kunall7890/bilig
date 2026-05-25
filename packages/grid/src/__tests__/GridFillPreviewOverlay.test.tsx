// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, test } from 'vitest'
import { GridFillPreviewOverlay } from '../GridFillPreviewOverlay.js'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'

describe('GridFillPreviewOverlay', () => {
  afterEach(() => {
    document.body.innerHTML = ''
  })

  test('keeps preview chrome tied to live geometry while the grid scrolls', async () => {
    const scrollStore = new WorkbookGridScrollStore()
    let scrollLeft = 0
    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(
        createElement(GridFillPreviewOverlay, {
          fillPreviewRange: { height: 2, width: 1, x: 2, y: 2 },
          getGeometrySnapshot: () => createGeometry(scrollLeft),
          scrollTransformStore: scrollStore,
        }),
      )
    })

    const preview = () => rootHost.querySelector<HTMLElement>('[data-grid-fill-preview="true"]')
    expect(preview()?.style.left).toBe('246px')

    scrollLeft = 100
    await act(async () => {
      scrollStore.setSnapshot({ scrollLeft, scrollTop: 0, tx: -scrollLeft, ty: 0 })
    })

    expect(preview()?.style.left).toBe('146px')

    await act(async () => {
      root.unmount()
    })
  })
})

function createGeometry(scrollLeft: number) {
  return createGridGeometrySnapshotFromAxes({
    columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
    dpr: 2,
    freezeCols: 0,
    freezeRows: 0,
    gridMetrics: getGridMetrics(),
    hostHeight: 240,
    hostWidth: 560,
    rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
    scrollLeft,
    scrollTop: 0,
    sheetName: 'Sheet1',
    updatedAt: 100,
  })
}
