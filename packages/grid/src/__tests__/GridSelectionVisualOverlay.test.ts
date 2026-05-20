// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { createColumnSliceSelection, createRangeSelection, createGridSelection } from '../gridSelection.js'
import { buildGridSelectionVisualRects, GridSelectionVisualOverlay } from '../GridSelectionVisualOverlay.js'

describe('GridSelectionVisualOverlay', () => {
  afterEach(() => {
    document.body.innerHTML = ''
  })

  test('builds crisp DOM visual rects for range selections', () => {
    const geometry = createGeometry()
    const selection = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 3])

    const rects = buildGridSelectionVisualRects({
      geometry,
      gridSelection: selection,
      selectedCell: [1, 1],
      selectionRange: selection.current?.range ?? null,
      showFillHandle: true,
    })

    expect(rects).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'selection-fill', bounds: expect.objectContaining({ x: 247, y: 45, width: 198, height: 18 }) }),
        expect.objectContaining({ role: 'selection-fill', bounds: expect.objectContaining({ x: 147, y: 65, width: 298, height: 38 }) }),
        expect.objectContaining({ role: 'selection-border', bounds: expect.objectContaining({ x: 146, y: 44, width: 300, height: 60 }) }),
        expect.objectContaining({ role: 'active-border', bounds: expect.objectContaining({ x: 146, y: 44, width: 100, height: 20 }) }),
        expect.objectContaining({ role: 'fill-handle', bounds: expect.objectContaining({ x: 442.5, y: 100.5, width: 7, height: 7 }) }),
      ]),
    )
    expect(rects).not.toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'selection-fill', bounds: expect.objectContaining({ x: 147, y: 45, width: 98, height: 18 }) }),
      ]),
    )
  })

  test('builds visible header, body, and active-cell rects for axis selections', () => {
    const geometry = createGeometry()
    const selection = createColumnSliceSelection(1, 3, 4)

    const rects = buildGridSelectionVisualRects({
      geometry,
      gridSelection: selection,
      selectedCell: [1, 4],
      selectionRange: selection.current?.range ?? null,
      showFillHandle: false,
    })

    expect(rects).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'header-fill', bounds: expect.objectContaining({ x: 147, y: 1, width: 98, height: 22 }) }),
        expect.objectContaining({ role: 'header-fill', bounds: expect.objectContaining({ x: 247, y: 1, width: 98, height: 22 }) }),
        expect.objectContaining({ role: 'header-fill', bounds: expect.objectContaining({ x: 347, y: 1, width: 98, height: 22 }) }),
        expect.objectContaining({ role: 'selection-fill', bounds: expect.objectContaining({ x: 147, y: 25, width: 298, height: 78 }) }),
        expect.objectContaining({ role: 'selection-fill', bounds: expect.objectContaining({ x: 247, y: 105, width: 198, height: 18 }) }),
        expect.objectContaining({ role: 'selection-fill', bounds: expect.objectContaining({ x: 147, y: 125, width: 298 }) }),
        expect.objectContaining({ role: 'active-border', bounds: expect.objectContaining({ x: 146, y: 104, width: 100, height: 20 }) }),
      ]),
    )
    expect(rects).not.toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'selection-fill', bounds: expect.objectContaining({ x: 147, y: 105, width: 98, height: 18 }) }),
      ]),
    )
  })

  test('clips scrolled axis header fills to their panes', () => {
    const geometry = createFrozenScrolledGeometry()
    const selection = createColumnSliceSelection(1, 3, 2)

    const rects = buildGridSelectionVisualRects({
      geometry,
      gridSelection: selection,
      selectedCell: [1, 2],
      selectionRange: selection.current?.range ?? null,
      showFillHandle: false,
    })

    expect(rects).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'header-fill', bounds: expect.objectContaining({ x: 147, y: 1, width: 48, height: 22 }) }),
        expect.objectContaining({ role: 'header-fill', bounds: expect.objectContaining({ x: 197, y: 1, width: 98, height: 22 }) }),
        expect.objectContaining({ role: 'header-fill', bounds: expect.objectContaining({ x: 297, y: 1, width: 98, height: 22 }) }),
      ]),
    )
    expect(rects).not.toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'header-fill', bounds: expect.objectContaining({ x: 97, y: 1, width: 98, height: 22 }) }),
      ]),
    )
  })

  test('keeps semantic visual rect keys stable while clipped header bounds move', () => {
    const selection = createColumnSliceSelection(1, 3, 2)
    const firstRects = buildGridSelectionVisualRects({
      geometry: createFrozenScrolledGeometry(50),
      gridSelection: selection,
      selectedCell: [1, 2],
      selectionRange: selection.current?.range ?? null,
      showFillHandle: false,
    })
    const nextRects = buildGridSelectionVisualRects({
      geometry: createFrozenScrolledGeometry(51),
      gridSelection: selection,
      selectedCell: [1, 2],
      selectionRange: selection.current?.range ?? null,
      showFillHandle: false,
    })

    const firstHeader = firstRects.find((rect) => rect.role === 'header-fill' && rect.key === 'header-fill:column:1')
    const nextHeader = nextRects.find((rect) => rect.role === 'header-fill' && rect.key === 'header-fill:column:1')

    expect(firstHeader?.key).toBe(nextHeader?.key)
    expect(firstHeader?.bounds).not.toEqual(nextHeader?.bounds)
    expect(firstHeader?.bounds.width).not.toBe(nextHeader?.bounds.width)
  })

  test('draws a single selected cell with active-cell chrome', () => {
    const geometry = createGeometry()
    const selection = createGridSelection(2, 4)

    const rects = buildGridSelectionVisualRects({
      geometry,
      gridSelection: selection,
      selectedCell: [2, 4],
      selectionRange: selection.current?.range ?? null,
      showFillHandle: true,
    })

    expect(rects).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'active-border', bounds: expect.objectContaining({ x: 246, y: 104, width: 100, height: 20 }) }),
        expect.objectContaining({ role: 'fill-handle', bounds: expect.objectContaining({ x: 342.5, y: 120.5, width: 7, height: 7 }) }),
      ]),
    )
    expect(rects.some((rect) => rect.role === 'selection-border')).toBe(false)
    expect(rects.some((rect) => rect.role === 'selection-fill')).toBe(false)
  })

  test('does not double-paint active cell sides already covered by the selected range perimeter', () => {
    const geometry = createGeometry()
    const topLeftSelection = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 3])
    const bottomRightSelection = createRangeSelection(createGridSelection(3, 3), [3, 3], [1, 1])

    const topLeftActiveBorder = buildGridSelectionVisualRects({
      geometry,
      gridSelection: topLeftSelection,
      selectedCell: [1, 1],
      selectionRange: topLeftSelection.current?.range ?? null,
      showFillHandle: true,
    }).find((rect) => rect.role === 'active-border')
    const bottomRightActiveBorder = buildGridSelectionVisualRects({
      geometry,
      gridSelection: bottomRightSelection,
      selectedCell: [3, 3],
      selectionRange: bottomRightSelection.current?.range ?? null,
      showFillHandle: true,
    }).find((rect) => rect.role === 'active-border')

    expect(topLeftActiveBorder?.borderSides).toEqual({ bottom: true, left: false, right: true, top: false })
    expect(topLeftActiveBorder?.strokeWidth).toBe(2)
    expect(bottomRightActiveBorder?.borderSides).toEqual({ bottom: false, left: true, right: false, top: true })
    expect(bottomRightActiveBorder?.strokeWidth).toBe(2)
  })

  test('builds hover chrome in the DOM overlay without covering the selected range', () => {
    const geometry = createGeometry()
    const selection = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 3])

    expect(
      buildGridSelectionVisualRects({
        geometry,
        gridSelection: selection,
        hoverCell: [4, 5],
        selectedCell: [1, 1],
        selectionRange: selection.current?.range ?? null,
        showFillHandle: false,
      }),
    ).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ role: 'hover-fill', bounds: expect.objectContaining({ x: 447, y: 125, width: 98, height: 18 }) }),
      ]),
    )

    expect(
      buildGridSelectionVisualRects({
        geometry,
        gridSelection: selection,
        hoverCell: [2, 2],
        selectedCell: [1, 1],
        selectionRange: selection.current?.range ?? null,
        showFillHandle: false,
      }).some((rect) => rect.role === 'hover-fill'),
    ).toBe(false)
  })

  test('chrome-only mode keeps fill geometry mounted but only paints crisp selection chrome', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const geometry = createGeometry()
    const selection = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 3])
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(GridSelectionVisualOverlay, {
          geometry,
          gridSelection: selection,
          hoverCell: [4, 5],
          selectedCell: [1, 1],
          selectionChromeMode: 'chrome-only',
          selectionRange: selection.current?.range ?? null,
          showFillHandle: true,
        }),
      )
    })

    const selectionFill = queryVisualElement(host, 'selection-fill')
    const hoverFill = queryVisualElement(host, 'hover-fill')
    const selectionBorder = queryVisualElement(host, 'selection-border')
    const activeBorder = queryVisualElement(host, 'active-border')
    const fillHandle = queryVisualElement(host, 'fill-handle')

    expect(selectionFill).toBeInstanceOf(HTMLElement)
    expect(hoverFill).toBeInstanceOf(HTMLElement)
    expect(selectionBorder).toBeInstanceOf(HTMLElement)
    expect(activeBorder).toBeInstanceOf(HTMLElement)
    expect(fillHandle).toBeInstanceOf(HTMLElement)
    expect(selectionFill?.style.opacity).toBe('0')
    expect(hoverFill?.style.opacity).toBe('0')
    expect(selectionBorder?.style.opacity).toBe('')
    expect(activeBorder?.style.opacity).toBe('')
    expect(fillHandle?.style.opacity).toBe('')

    await act(async () => {
      root.unmount()
    })
  })
})

function createGeometry() {
  const metrics = getGridMetrics()
  return createGridGeometrySnapshotFromAxes({
    columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
    dpr: 2,
    freezeCols: 0,
    freezeRows: 0,
    gridMetrics: metrics,
    hostHeight: 240,
    hostWidth: 560,
    rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
    scrollLeft: 0,
    scrollTop: 0,
    sheetName: 'Sheet1',
    updatedAt: 100,
  })
}

function queryVisualElement(host: ParentNode, role: string): HTMLElement | null {
  const node = host.querySelector(`[data-grid-selection-visual-role="${role}"]`)
  return node instanceof HTMLElement ? node : null
}

function createFrozenScrolledGeometry(scrollLeft = 50) {
  const metrics = getGridMetrics()
  return createGridGeometrySnapshotFromAxes({
    columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
    dpr: 2,
    freezeCols: 1,
    freezeRows: 1,
    gridMetrics: metrics,
    hostHeight: 220,
    hostWidth: 520,
    rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
    scrollLeft,
    scrollTop: 10,
    sheetName: 'Sheet1',
    updatedAt: 100,
  })
}
