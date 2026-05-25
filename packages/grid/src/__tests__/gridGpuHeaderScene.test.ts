import { expect, test } from 'vitest'
import { CompactSelection } from '../gridTypes.js'
import {
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  getGridMetrics,
} from '../gridMetrics.js'
import { parseGpuColor } from '../gridGpuPrimitives.js'
import { buildGridGpuHeaderScene } from '../gridGpuHeaderScene.js'

const DATA_LEFT = PRODUCT_ROW_MARKER_WIDTH
const DATA_TOP = PRODUCT_HEADER_HEIGHT
const COLUMN_WIDTH = PRODUCT_COLUMN_WIDTH
const ROW_HEIGHT = PRODUCT_ROW_HEIGHT

const palette = {
  gridLineColor: parseGpuColor('#e3e9f0'),
  headerFillColor: parseGpuColor('#f8f9fa'),
  headerSelectedFillColor: parseGpuColor('#e6f4ea'),
  headerHoverFillColor: parseGpuColor('#f1f3f4'),
  headerDragAnchorFillColor: parseGpuColor('#d7eadf'),
  selectionFillColor: parseGpuColor('rgba(31, 122, 67, 0.06)'),
  resizeGuideColor: parseGpuColor('rgba(121, 105, 123, 0.82)'),
  resizeGuideGlowColor: parseGpuColor('rgba(168, 158, 169, 0.18)'),
}

function createSelection() {
  return {
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
    current: undefined,
  }
}

test('builds GPU-backed header backgrounds and selection highlights', () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: createSelection(),
    selectedCell: [2, 3],
    selectionRange: { x: 2, y: 3, width: 1, height: 1 },
    visibleItems: [[2, 3]],
    visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: null,
    resizeGuideRow: null,
    activeHeaderDrag: null,
  })

  expect(scene.fillRects).toContainEqual({
    x: 0,
    y: 0,
    width: DATA_LEFT,
    height: DATA_TOP,
    color: { r: 248 / 255, g: 249 / 255, b: 250 / 255, a: 1 },
  })
  expect(scene.fillRects).toContainEqual({
    x: DATA_LEFT,
    y: 0,
    width: COLUMN_WIDTH,
    height: DATA_TOP,
    color: { r: 230 / 255, g: 244 / 255, b: 234 / 255, a: 1 },
  })
  expect(scene.fillRects).toContainEqual({
    x: 0,
    y: DATA_TOP,
    width: DATA_LEFT,
    height: ROW_HEIGHT,
    color: { r: 230 / 255, g: 244 / 255, b: 234 / 255, a: 1 },
  })
})

test('fills selected column axis through the header/body seam without a second border', () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: {
      columns: CompactSelection.fromSingleSelection([1, 3]),
      rows: CompactSelection.empty(),
      current: undefined,
    },
    selectedCell: [1, 3],
    selectionRange: null,
    visibleItems: [
      [1, 3],
      [2, 3],
    ],
    visibleRegion: { range: { x: 1, y: 3, width: 2, height: 1 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: DATA_LEFT + COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: 100, height: ROW_HEIGHT }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: null,
    resizeGuideRow: null,
    activeHeaderDrag: null,
  })

  expect(scene.fillRects).toContainEqual({
    x: DATA_LEFT,
    y: DATA_TOP,
    width: 2 * COLUMN_WIDTH,
    height: ROW_HEIGHT,
    color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 0.06 },
  })
  expect(scene.borderRects).not.toContainEqual({
    x: DATA_LEFT,
    y: DATA_TOP - 1,
    width: COLUMN_WIDTH,
    height: 1,
    color: { r: 227 / 255, g: 233 / 255, b: 240 / 255, a: 1 },
  })
  expect(scene.borderRects).not.toContainEqual({
    x: DATA_LEFT + COLUMN_WIDTH,
    y: DATA_TOP - 1,
    width: COLUMN_WIDTH,
    height: 1,
    color: { r: 227 / 255, g: 233 / 255, b: 240 / 255, a: 1 },
  })
})

test('fills selected row axis through the row-header/body seam without a second border', () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: {
      columns: CompactSelection.empty(),
      rows: CompactSelection.fromSingleSelection([3, 6]),
      current: undefined,
    },
    selectedCell: [1, 3],
    selectionRange: null,
    visibleItems: [
      [1, 3],
      [2, 3],
    ],
    visibleRegion: { range: { x: 1, y: 3, width: 2, height: 3 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: DATA_LEFT + COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: 100, height: ROW_HEIGHT }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: null,
    resizeGuideRow: null,
    activeHeaderDrag: null,
  })

  expect(scene.fillRects).toContainEqual({
    x: DATA_LEFT,
    y: DATA_TOP,
    width: 2 * COLUMN_WIDTH,
    height: 3 * ROW_HEIGHT,
    color: { r: 31 / 255, g: 122 / 255, b: 67 / 255, a: 0.06 },
  })
  expect(scene.borderRects).not.toContainEqual({
    x: DATA_LEFT - 1,
    y: DATA_TOP,
    width: 1,
    height: ROW_HEIGHT,
    color: { r: 227 / 255, g: 233 / 255, b: 240 / 255, a: 1 },
  })
  expect(scene.borderRects).not.toContainEqual({
    x: DATA_LEFT - 1,
    y: DATA_TOP + ROW_HEIGHT,
    width: 1,
    height: ROW_HEIGHT,
    color: { r: 227 / 255, g: 233 / 255, b: 240 / 255, a: 1 },
  })
  expect(scene.borderRects).not.toContainEqual({
    x: DATA_LEFT - 1,
    y: DATA_TOP + 2 * ROW_HEIGHT,
    width: 1,
    height: ROW_HEIGHT,
    color: { r: 227 / 255, g: 233 / 255, b: 240 / 255, a: 1 },
  })
})

test('builds GPU resize guides for hovered columns', () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: createSelection(),
    selectedCell: [0, 0],
    selectionRange: null,
    visibleItems: [[2, 3]],
    visibleRegion: { range: { x: 2, y: 3, width: 1, height: 2 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: 2,
    resizeGuideRow: null,
    activeHeaderDrag: null,
  })

  expect(scene.fillRects).toContainEqual({
    x: DATA_LEFT + COLUMN_WIDTH - 2,
    y: 0,
    width: 3,
    height: DATA_TOP + 2 * ROW_HEIGHT,
    color: { r: 168 / 255, g: 158 / 255, b: 169 / 255, a: 0.18 },
  })
  expect(scene.borderRects).toContainEqual({
    x: DATA_LEFT + COLUMN_WIDTH - 1,
    y: 0,
    width: 1,
    height: DATA_TOP + 2 * ROW_HEIGHT,
    color: { r: 121 / 255, g: 105 / 255, b: 123 / 255, a: 0.82 },
  })
})

test('builds GPU drag guides for active column header drags', () => {
  const scene = buildGridGpuHeaderScene({
    palette,
    columnWidths: {},
    gridMetrics: getGridMetrics(),
    gridSelection: {
      columns: CompactSelection.fromSingleSelection([1, 3]),
      rows: CompactSelection.empty(),
      current: undefined,
    },
    selectedCell: [1, 2],
    selectionRange: null,
    visibleItems: [
      [1, 2],
      [2, 2],
    ],
    visibleRegion: { range: { x: 1, y: 2, width: 2, height: 2 }, tx: 0, ty: 0 },
    getCellBounds: () => ({ x: DATA_LEFT + COLUMN_WIDTH, y: DATA_TOP + ROW_HEIGHT + 4, width: 100, height: ROW_HEIGHT }),
    rowHeights: {},
    hoveredHeader: null,
    resizeGuideColumn: null,
    resizeGuideRow: null,
    activeHeaderDrag: { kind: 'column', index: 1 },
  })

  expect(scene.borderRects).toContainEqual({
    x: DATA_LEFT,
    y: 0,
    width: 1,
    height: DATA_TOP + 2 * ROW_HEIGHT,
    color: { r: 121 / 255, g: 105 / 255, b: 123 / 255, a: 0.82 },
  })
  expect(scene.borderRects).toContainEqual({
    x: DATA_LEFT + 2 * COLUMN_WIDTH - 1,
    y: 0,
    width: 1,
    height: DATA_TOP + 2 * ROW_HEIGHT,
    color: { r: 121 / 255, g: 105 / 255, b: 123 / 255, a: 0.82 },
  })
  expect(scene.fillRects).toContainEqual({
    x: DATA_LEFT,
    y: DATA_TOP - 3,
    width: COLUMN_WIDTH,
    height: 3,
    color: { r: 121 / 255, g: 105 / 255, b: 123 / 255, a: 0.82 },
  })
})
