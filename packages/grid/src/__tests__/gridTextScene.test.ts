import { describe, expect, test } from 'vitest'
import { MAX_COLS, MAX_ROWS, OPTIMISTIC_CELL_SNAPSHOT_FLAG, ValueTag, type CellStyleRecord } from '@bilig/protocol'
import { buildGridTextScene } from '../gridTextScene.js'
import { getResolvedCellFontFamily } from '../gridCells.js'
import type { GridEngineLike } from '../grid-engine.js'
import {
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  getGridMetrics,
} from '../gridMetrics.js'
import {
  WORKBOOK_DEFAULT_FONT_SIZE,
  WORKBOOK_HEADER_FONT_SANS,
  WORKBOOK_HEADER_FONT_WEIGHT,
  workbookFontPointSizeToCssPx,
  workbookHeaderFontPointSizeToCssPx,
  workbookThemeColors,
} from '../workbookTheme.js'

type TestCellValue =
  | { tag: ValueTag.Empty }
  | { tag: ValueTag.Number; value: number }
  | { tag: ValueTag.Boolean; value: boolean }
  | { tag: ValueTag.String; value: string; stringId?: number }
  | { tag: ValueTag.Error; code: number }

function createCellSnapshot(value: TestCellValue, styleId: string | undefined = 'style-1') {
  return {
    sheetName: 'Sheet1',
    address: 'A1',
    input: '',
    value,
    flags: 0,
    version: 0,
    ...(styleId ? { styleId } : {}),
  }
}

type TestCellSnapshot = ReturnType<typeof createCellSnapshot>
const CELL_FONT_FAMILY = getResolvedCellFontFamily()
const HEADER_FONT_FAMILY = WORKBOOK_HEADER_FONT_SANS
const CELL_FONT_SIZE = workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)
const CELL_FONT = `400 ${CELL_FONT_SIZE}px ${CELL_FONT_FAMILY}`
const HEADER_FONT_SIZE = workbookHeaderFontPointSizeToCssPx()
const HEADER_FONT = `${WORKBOOK_HEADER_FONT_WEIGHT} ${HEADER_FONT_SIZE}px ${HEADER_FONT_FAMILY}`
const HEADER_SELECTED_COLOR = workbookThemeColors.accent
const HEADER_DRAG_COLOR = workbookThemeColors.accentDark
const HEADER_HOVER_COLOR = workbookThemeColors.text
const CELL_TEXT_COLOR = workbookThemeColors.text
const DATA_LEFT = PRODUCT_ROW_MARKER_WIDTH
const DATA_TOP = PRODUCT_HEADER_HEIGHT
const COLUMN_WIDTH = PRODUCT_COLUMN_WIDTH
const ROW_HEIGHT = PRODUCT_ROW_HEIGHT

function makeEngine(
  styles: Record<string, CellStyleRecord>,
  snapshots: TestCellSnapshot | Record<string, TestCellSnapshot> = createCellSnapshot({
    tag: ValueTag.String,
    value: 'hello',
  }),
  mergedRanges: Record<string, { sheetName: string; startAddress: string; endAddress: string }> = {},
): GridEngineLike {
  return {
    getCell: (_sheetName, address) =>
      'address' in snapshots ? snapshots : (snapshots[address] ?? createCellSnapshot({ tag: ValueTag.Empty }, undefined)),
    getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
    getMergeRange: (_sheetName, address) => mergedRanges[address],
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => undefined,
    },
  }
}

describe('gridTextScene', () => {
  test('builds cell text items with resolved alignment and style', () => {
    const engine = makeEngine({
      'style-1': {
        alignment: { horizontal: 'right' },
        font: { bold: true, color: '#ff0000', italic: true, size: 14 },
      },
    })

    const scene = buildGridTextScene({
      engine,
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [[0, 0]],
      visibleRegion: { range: { x: 0, y: 0, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 100, top: 200, width: 300, height: 200 },
      getCellBounds: () => ({ x: 110, y: 220, width: 90, height: ROW_HEIGHT }),
    })

    expect(scene.items).toContainEqual({
      x: 10,
      y: 20,
      width: 90,
      height: ROW_HEIGHT,
      clipInsetTop: DATA_TOP - 20,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: DATA_LEFT - 10,
      col: 0,
      text: 'hello',
      align: 'right',
      wrap: false,
      color: '#ff0000',
      font: `italic 700 ${workbookFontPointSizeToCssPx(14)}px ${CELL_FONT_FAMILY}`,
      fontSize: workbookFontPointSizeToCssPx(14),
      row: 0,
      underline: false,
      strike: false,
    })
  })

  test('renders text once across merged cells', () => {
    const engine = makeEngine(
      {},
      {
        A1: createCellSnapshot({ tag: ValueTag.String, value: 'merged title' }),
        B1: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
      },
      {
        A1: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
        B1: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      },
    )

    const scene = buildGridTextScene({
      engine,
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [
        [0, 0],
        [1, 0],
      ],
      visibleRegion: { range: { x: 0, y: 0, width: 2, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 100, top: 200, width: 300, height: 200 },
      getCellBounds: (col) => ({ x: 110 + col * 90, y: 220, width: 90, height: ROW_HEIGHT }),
    })

    const dataItems = scene.items.filter((item) => item.text === 'merged title')
    expect(dataItems).toHaveLength(1)
    expect(dataItems[0]).toMatchObject({
      col: 0,
      row: 0,
      x: 10,
      y: 20,
      width: 180,
      height: ROW_HEIGHT,
    })
  })

  test('keeps merged ranges visible when the top-left cell is blank', () => {
    const engine = makeEngine(
      {},
      {
        A1: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        B1: createCellSnapshot({ tag: ValueTag.String, value: 'visible after merge' }),
      },
      {
        A1: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
        B1: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      },
    )

    const scene = buildGridTextScene({
      engine,
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [
        [0, 0],
        [1, 0],
      ],
      visibleRegion: { range: { x: 0, y: 0, width: 2, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 100, top: 200, width: 300, height: 200 },
      getCellBounds: (col) => ({ x: 110 + col * 90, y: 220, width: 90, height: ROW_HEIGHT }),
    })

    const dataItems = scene.items.filter((item) => item.text === 'visible after merge')
    expect(dataItems).toHaveLength(1)
    expect(dataItems[0]).toMatchObject({
      col: 0,
      row: 0,
      text: 'visible after merge',
      width: 180,
    })
  })

  test('adds column headers and row markers with selected header emphasis', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 3],
      selectionRange: { x: 2, y: 3, width: 1, height: 1 },
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items).toContainEqual({
      x: DATA_LEFT,
      y: 0,
      width: COLUMN_WIDTH,
      height: DATA_TOP,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: DATA_TOP,
      width: DATA_LEFT,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '4',
      align: 'right',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
  })

  test('adds hovered and active-drag header text emphasis', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      activeHeaderDrag: { kind: 'column', index: 2 },
      hoveredHeader: { kind: 'row', index: 4 },
      selectedCell: [2, 3],
      selectionRange: { x: 2, y: 3, width: 1, height: 1 },
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 2 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items).toContainEqual({
      x: DATA_LEFT,
      y: 0,
      width: COLUMN_WIDTH,
      height: DATA_TOP,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: HEADER_DRAG_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: DATA_TOP,
      width: DATA_LEFT,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '4',
      align: 'right',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: DATA_TOP + ROW_HEIGHT,
      width: DATA_LEFT,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '5',
      align: 'right',
      wrap: false,
      color: HEADER_HOVER_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
  })

  test('renders frozen header labels beside the scrollable pane', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      selectionRange: { x: 0, y: 0, width: 1, height: 1 },
      sheetName: 'Sheet1',
      visibleItems: [
        [0, 0],
        [2, 0],
        [0, 3],
        [2, 3],
      ],
      visibleRegion: {
        range: { x: 2, y: 3, width: 1, height: 1 },
        tx: 0,
        ty: 0,
        freezeRows: 1,
        freezeCols: 1,
      },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: (col, row) => ({
        x: col === 0 ? DATA_LEFT : DATA_LEFT + COLUMN_WIDTH,
        y: row === 0 ? DATA_TOP : DATA_TOP + ROW_HEIGHT,
        width: COLUMN_WIDTH,
        height: ROW_HEIGHT,
      }),
    })

    expect(scene.items).toContainEqual({
      x: DATA_LEFT,
      y: 0,
      width: COLUMN_WIDTH,
      height: DATA_TOP,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'A',
      align: 'center',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: DATA_LEFT + COLUMN_WIDTH,
      y: 0,
      width: COLUMN_WIDTH,
      height: DATA_TOP,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: workbookThemeColors.textMuted,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: DATA_TOP,
      width: DATA_LEFT,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '1',
      align: 'right',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: DATA_TOP + ROW_HEIGHT,
      width: DATA_LEFT,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '4',
      align: 'right',
      wrap: false,
      color: workbookThemeColors.textMuted,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
  })

  test('adds resize-guide emphasis to the active column header label', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.Empty })),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      resizeGuideColumn: 2,
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items).toContainEqual({
      x: DATA_LEFT,
      y: 0,
      width: COLUMN_WIDTH,
      height: DATA_TOP,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: HEADER_DRAG_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
  })

  test('omits the active editing cell text item while preserving headers', () => {
    const scene = buildGridTextScene({
      engine: makeEngine({}, createCellSnapshot({ tag: ValueTag.String, value: 'editing' })),
      columnWidths: {},
      editingCell: [2, 3],
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 3],
      sheetName: 'Sheet1',
      visibleItems: [[2, 3]],
      visibleRegion: { range: { x: 2, y: 3, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 3 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items.some((item) => item.text === 'editing')).toBe(false)
    expect(scene.items).toContainEqual({
      x: DATA_LEFT,
      y: 0,
      width: COLUMN_WIDTH,
      height: DATA_TOP,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'C',
      align: 'center',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: HEADER_FONT,
      fontSize: HEADER_FONT_SIZE,
      underline: false,
      strike: false,
    })
  })

  test('renders the selected cell from the authoritative snapshot when the engine cache lags', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          C5: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.String, value: 'selected' }),
        address: 'C5',
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 4 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items.find((item) => item.text === 'selected')).toEqual({
      x: DATA_LEFT + 2 * COLUMN_WIDTH,
      y: DATA_TOP + 4 * ROW_HEIGHT,
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 28,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      col: 2,
      text: 'selected',
      align: 'left',
      wrap: false,
      color: CELL_TEXT_COLOR,
      font: CELL_FONT,
      fontSize: CELL_FONT_SIZE,
      row: 4,
      spillColEnd: undefined,
      underline: false,
      strike: false,
    })
  })

  test('falls back to the engine cell when the selected snapshot address does not match', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          C5: createCellSnapshot({ tag: ValueTag.String, value: 'engine text' }),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.String, value: 'stale text' }),
        address: 'B4',
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 4 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items.at(-1)?.text).toBe('engine text')
  })

  test('keeps the engine text when the selected snapshot is temporarily empty', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          C5: createCellSnapshot({ tag: ValueTag.String, value: 'engine text' }),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        address: 'C5',
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 4 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items.at(-1)?.text).toBe('engine text')
  })

  test('uses a newer empty selected snapshot to suppress stale engine text', () => {
    const scene = buildGridTextScene({
      contentMode: 'data',
      engine: makeEngine(
        {},
        {
          C5: {
            ...createCellSnapshot({ tag: ValueTag.String, value: 'engine text' }),
            address: 'C5',
            version: 3,
          },
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        address: 'C5',
        version: 4,
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 4 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items.some((item) => item.text === 'engine text')).toBe(false)
  })

  test('uses an optimistic empty selected snapshot to suppress stale engine text while clear confirmation races', () => {
    const scene = buildGridTextScene({
      contentMode: 'data',
      engine: makeEngine(
        {},
        {
          C5: {
            ...createCellSnapshot({ tag: ValueTag.String, value: 'engine text' }),
            address: 'C5',
            version: 8,
          },
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [2, 4],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        address: 'C5',
        flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
        version: 8,
      },
      sheetName: 'Sheet1',
      visibleItems: [[2, 4]],
      visibleRegion: { range: { x: 2, y: 4, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT + 2 * COLUMN_WIDTH, y: DATA_TOP + 4 * ROW_HEIGHT, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items.some((item) => item.text === 'engine text')).toBe(false)
  })

  test('scales header text with the selected font size for full-sheet selections', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {
          'style-selected': {
            id: 'style-selected',
            font: { size: 20 },
          },
        },
        {
          A1: {
            ...createCellSnapshot({ tag: ValueTag.Number, value: 1200 }, 'style-selected'),
            address: 'A1',
          },
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      selectedCellSnapshot: {
        ...createCellSnapshot({ tag: ValueTag.Number, value: 1200 }, 'style-selected'),
        address: 'A1',
      },
      selectionRange: { x: 0, y: 0, width: MAX_COLS, height: MAX_ROWS },
      sheetName: 'Sheet1',
      visibleItems: [[0, 0]],
      visibleRegion: { range: { x: 0, y: 0, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: () => ({ x: DATA_LEFT, y: DATA_TOP, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items).toContainEqual({
      x: DATA_LEFT,
      y: 0,
      width: COLUMN_WIDTH,
      height: DATA_TOP,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: 'A',
      align: 'center',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: `${WORKBOOK_HEADER_FONT_WEIGHT} ${workbookFontPointSizeToCssPx(20)}px ${HEADER_FONT_FAMILY}`,
      fontSize: workbookFontPointSizeToCssPx(20),
      underline: false,
      strike: false,
    })
    expect(scene.items).toContainEqual({
      x: 0,
      y: DATA_TOP,
      width: DATA_LEFT,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      text: '1',
      align: 'right',
      wrap: false,
      color: HEADER_SELECTED_COLOR,
      font: `${WORKBOOK_HEADER_FONT_WEIGHT} ${workbookFontPointSizeToCssPx(20)}px ${HEADER_FONT_FAMILY}`,
      fontSize: workbookFontPointSizeToCssPx(20),
      underline: false,
      strike: false,
    })
  })

  test('spills left-aligned string text across contiguous empty cells', () => {
    const scene = buildGridTextScene({
      engine: makeEngine(
        {},
        {
          B12: createCellSnapshot({ tag: ValueTag.String, value: 'spill text' }),
          C12: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
          D12: createCellSnapshot({ tag: ValueTag.Empty }, undefined),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [0, 0],
      sheetName: 'Sheet1',
      visibleItems: [
        [1, 11],
        [2, 11],
        [3, 11],
      ],
      visibleRegion: { range: { x: 1, y: 11, width: 3, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: 320, height: 240 },
      getCellBounds: (col) => ({
        x: DATA_LEFT + col * COLUMN_WIDTH,
        y: 266,
        width: COLUMN_WIDTH,
        height: ROW_HEIGHT,
      }),
    })

    expect(scene.items).toContainEqual({
      x: DATA_LEFT + COLUMN_WIDTH,
      y: 266,
      width: 312,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 132,
      clipInsetBottom: 48,
      clipInsetLeft: 0,
      col: 1,
      spillColEnd: 3,
      text: 'spill text',
      align: 'left',
      wrap: false,
      color: CELL_TEXT_COLOR,
      font: CELL_FONT,
      fontSize: CELL_FONT_SIZE,
      row: 11,
      underline: false,
      strike: false,
    })
  })

  test('does not clip the first visible cell in data pane mode', () => {
    const scene = buildGridTextScene({
      contentMode: 'data',
      engine: makeEngine(
        {},
        {
          B2: createCellSnapshot({ tag: ValueTag.Number, value: 123 }),
        },
      ),
      columnWidths: {},
      gridMetrics: getGridMetrics(),
      selectedCell: [1, 1],
      sheetName: 'Sheet1',
      visibleItems: [[1, 1]],
      visibleRegion: { range: { x: 1, y: 1, width: 1, height: 1 }, tx: 0, ty: 0 },
      hostBounds: { left: 0, top: 0, width: COLUMN_WIDTH, height: ROW_HEIGHT },
      getCellBounds: () => ({ x: 0, y: 0, width: COLUMN_WIDTH, height: ROW_HEIGHT }),
    })

    expect(scene.items).toContainEqual({
      x: 0,
      y: 0,
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT,
      clipInsetTop: 0,
      clipInsetRight: 0,
      clipInsetBottom: 0,
      clipInsetLeft: 0,
      col: 1,
      text: '123',
      align: 'right',
      wrap: false,
      color: CELL_TEXT_COLOR,
      font: CELL_FONT,
      fontSize: CELL_FONT_SIZE,
      row: 1,
      underline: false,
      strike: false,
    })
  })
})
