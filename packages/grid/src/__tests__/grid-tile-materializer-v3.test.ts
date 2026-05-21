import { describe, expect, test } from 'vitest'
import { ValueTag, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type CellSnapshot, type CellStyleRecord } from '@bilig/protocol'
import { indexToColumn } from '@bilig/formula'
import { getGridMetrics } from '../gridMetrics.js'
import type { GridEngineLike } from '../grid-engine.js'
import { buildLocalFixedRenderTiles } from '../renderer-v3/local-render-tile-materializer.js'
import { materializeGridRenderTileV3 } from '../renderer-v3/grid-tile-materializer.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import { GRID_TILE_PACKET_V3_MAGIC } from '../renderer-v3/tile-packet-v3.js'

function createCellSnapshot(address: string, value: CellSnapshot['value'], styleId: string | undefined = 'style-1'): CellSnapshot {
  return {
    address,
    flags: 0,
    input: '',
    sheetName: 'Sheet1',
    value,
    version: 0,
    ...(styleId ? { styleId } : {}),
  }
}

function makeEngine(
  cells: Record<string, CellSnapshot>,
  styles: Record<string, CellStyleRecord> = {},
  mergedRanges: Record<string, { sheetName: string; startAddress: string; endAddress: string }> = {},
): GridEngineLike {
  return {
    getCell: (_sheetName, address) => cells[address] ?? createCellSnapshot(address, { tag: ValueTag.Empty }, undefined),
    getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
    getMergeRange: (_sheetName, address) => mergedRanges[address],
    subscribeCells: () => () => undefined,
    workbook: {
      getSheet: () => undefined,
    },
  }
}

function hasHorizontalBorderAtY(tile: { readonly rectCount: number; readonly rectInstances: Float32Array }, y: number): boolean {
  return hasBorderMatching(tile, (offset) => {
    const rectY = tile.rectInstances[offset + 1]
    const width = tile.rectInstances[offset + 2]
    const height = tile.rectInstances[offset + 3]
    return Math.abs(rectY - y) < 0.01 && width > 1 && height <= 1.5
  })
}

function hasVerticalBorderAtX(tile: { readonly rectCount: number; readonly rectInstances: Float32Array }, x: number): boolean {
  return hasBorderMatching(tile, (offset) => {
    const rectX = tile.rectInstances[offset + 0]
    const width = tile.rectInstances[offset + 2]
    const height = tile.rectInstances[offset + 3]
    return Math.abs(rectX - x) < 0.01 && width <= 1.5 && height > 1
  })
}

function hasBorderMatching(
  tile: { readonly rectCount: number; readonly rectInstances: Float32Array },
  matches: (offset: number) => boolean,
): boolean {
  for (let index = 0; index < tile.rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const borderAlpha = tile.rectInstances[offset + 11]
    const instanceKind = tile.rectInstances[offset + 13]
    if (instanceKind === 1 && borderAlpha > 0 && matches(offset)) {
      return true
    }
  }
  return false
}

describe('renderer-v3 grid tile materializer', () => {
  test('materializes fixed content tiles with native v3 packets instead of v2 scene packets', () => {
    const gridMetrics = getGridMetrics()
    const dirtyLocalRows = new Uint32Array([0, 0])
    const dirtyLocalCols = new Uint32Array([0, 0])
    const dirtyMasks = new Uint32Array([5])
    const tile = materializeGridRenderTileV3({
      axisSeqX: 5,
      axisSeqY: 6,
      cameraSeq: 7,
      columnWidths: {},
      dprBucket: 2,
      engine: makeEngine(
        {
          A1: createCellSnapshot('A1', { tag: ValueTag.String, value: 'alpha' }),
        },
        {
          'style-1': {
            fill: { backgroundColor: '#ff0000' },
          },
        },
      ),
      freezeSeq: 8,
      glyphAtlasSeq: 9,
      gridMetrics,
      materializedAtSeq: 10,
      packetSeq: 11,
      rectSeq: 12,
      rowHeights: {},
      sheetId: 3,
      sheetOrdinal: 1,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 13,
      textSeq: 14,
      valueSeq: 15,
      dirtyLocalCols,
      dirtyLocalRows,
      dirtyMasks,
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
        rowStart: 0,
      },
    })

    expect(tile.packet?.magic).toBe(GRID_TILE_PACKET_V3_MAGIC)
    expect(tile.tileId).toBe(tile.packet?.tileKey)
    expect(tile.coord).toMatchObject({ colTile: 0, dprBucket: 2, rowTile: 0, sheetId: 3, sheetOrdinal: 1 })
    expect(tile.version).toEqual({ axisX: 5, axisY: 6, freeze: 8, styles: 13, text: 14, values: 15 })
    expect(tile.bounds).toEqual({ colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1, colStart: 0, rowEnd: VIEWPORT_TILE_ROW_COUNT - 1, rowStart: 0 })
    expect(tile.textRuns).toContainEqual(expect.objectContaining({ text: 'alpha', x: 0, y: 0 }))
    expect(tile.packet?.dirtyLocalRows).toBe(dirtyLocalRows)
    expect(tile.packet?.dirtyLocalCols).toBe(dirtyLocalCols)
    expect(tile.packet?.dirtyMasks).toBe(dirtyMasks)
    expect(tile.dirtyLocalRows).toBe(dirtyLocalRows)
    expect(tile.rectCount).toBeGreaterThan(0)
    expect(tile.rectInstances.length).toBeGreaterThanOrEqual(tile.rectCount * 20)
    expect(tile.packet && 'key' in tile.packet).toBe(false)
  })

  test('local fixed tile generation returns one v3 packet per fixed protocol tile', () => {
    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 4,
      columnWidths: {},
      dprBucket: 1,
      engine: makeEngine({
        A1: createCellSnapshot('A1', { tag: ValueTag.String, value: 'visible' }),
      }),
      freezeSeq: 1,
      generation: 21,
      gridMetrics: getGridMetrics(),
      rowHeights: {},
      sheetId: 2,
      sheetOrdinal: 0,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT,
        rowStart: 0,
      },
    })

    expect(tiles).toHaveLength(4)
    expect(tiles.map((tile) => tile.packet?.magic)).toEqual([
      GRID_TILE_PACKET_V3_MAGIC,
      GRID_TILE_PACKET_V3_MAGIC,
      GRID_TILE_PACKET_V3_MAGIC,
      GRID_TILE_PACKET_V3_MAGIC,
    ])
  })

  test('fixed data tiles omit leading gridlines so adjacent tile boundaries do not double paint seams', () => {
    const gridMetrics = getGridMetrics()
    const tile = materializeGridRenderTileV3({
      axisSeqX: 5,
      axisSeqY: 6,
      cameraSeq: 7,
      columnWidths: {},
      dprBucket: 1,
      engine: makeEngine({}),
      freezeSeq: 8,
      glyphAtlasSeq: 9,
      gridMetrics,
      materializedAtSeq: 10,
      packetSeq: 11,
      rectSeq: 12,
      rowHeights: {},
      sheetId: 3,
      sheetOrdinal: 1,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 13,
      textSeq: 14,
      valueSeq: 15,
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT * 2 - 1,
        colStart: VIEWPORT_TILE_COLUMN_COUNT,
        rowEnd: VIEWPORT_TILE_ROW_COUNT * 2 - 1,
        rowStart: VIEWPORT_TILE_ROW_COUNT,
      },
    })

    expect(hasHorizontalBorderAtY(tile, 0)).toBe(false)
    expect(hasVerticalBorderAtX(tile, 0)).toBe(false)
    expect(hasHorizontalBorderAtY(tile, gridMetrics.rowHeight - 1)).toBe(true)
    expect(hasVerticalBorderAtX(tile, gridMetrics.columnWidth - 1)).toBe(true)
  })

  test('local fixed tile generation paints the selected-cell snapshot override', () => {
    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 4,
      columnWidths: {},
      dprBucket: 1,
      engine: makeEngine({}),
      freezeSeq: 1,
      generation: 21,
      gridMetrics: getGridMetrics(),
      rowHeights: {},
      selectedCell: [3, 52],
      selectedCellSnapshot: createCellSnapshot('D53', { tag: ValueTag.String, value: 'Month 1' }),
      sheetId: 2,
      sheetOrdinal: 0,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT * 2 - 1,
        rowStart: VIEWPORT_TILE_ROW_COUNT,
      },
    })

    expect(tiles.flatMap((tile) => tile.textRuns.map((run) => run.text))).toContain('Month 1')
  })

  test('local fixed tile generation suppresses the active editor cell text', () => {
    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 4,
      columnWidths: {},
      dprBucket: 1,
      editingCell: [3, 52],
      engine: makeEngine({}),
      freezeSeq: 1,
      generation: 21,
      gridMetrics: getGridMetrics(),
      rowHeights: {},
      selectedCell: [3, 52],
      selectedCellSnapshot: createCellSnapshot('D53', { tag: ValueTag.String, value: 'Month 1' }),
      sheetId: 2,
      sheetOrdinal: 0,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT * 2 - 1,
        rowStart: VIEWPORT_TILE_ROW_COUNT,
      },
    })

    expect(tiles.flatMap((tile) => tile.textRuns.map((run) => run.text))).not.toContain('Month 1')
  })

  test('local fixed tile generation preserves runtime freeze sequence', () => {
    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 4,
      columnWidths: {},
      dprBucket: 1,
      engine: makeEngine({}),
      freezeSeq: 37,
      generation: 21,
      gridMetrics: getGridMetrics(),
      rowHeights: {},
      sheetId: 2,
      sheetOrdinal: 0,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
        rowStart: 0,
      },
    })

    expect(tiles).toHaveLength(1)
    expect(tiles[0]?.version.freeze).toBe(37)
  })

  test('materializes visible text for merged cells whose anchor is blank', () => {
    const tile = materializeGridRenderTileV3({
      axisSeqX: 5,
      axisSeqY: 6,
      cameraSeq: 7,
      columnWidths: {},
      dprBucket: 1,
      engine: makeEngine(
        {
          A1: createCellSnapshot('A1', { tag: ValueTag.Empty }, undefined),
          B1: createCellSnapshot('B1', { tag: ValueTag.String, value: 'merged value' }, undefined),
        },
        {},
        {
          A1: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
          B1: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
        },
      ),
      freezeSeq: 8,
      glyphAtlasSeq: 9,
      gridMetrics: getGridMetrics(),
      materializedAtSeq: 10,
      packetSeq: 11,
      rectSeq: 12,
      rowHeights: {},
      sheetId: 3,
      sheetOrdinal: 1,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 13,
      textSeq: 14,
      valueSeq: 15,
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
        rowStart: 0,
      },
    })

    expect(tile.textRuns).toContainEqual(expect.objectContaining({ col: 0, row: 0, text: 'merged value' }))
    expect(tile.textRuns).not.toContainEqual(expect.objectContaining({ col: 1, row: 0, text: 'merged value' }))
    expect(tile.textRuns.find((run) => run.text === 'merged value')?.width).toBeGreaterThan(getGridMetrics().columnWidth)
  })

  test('materializes inbound left-aligned text spill runs across fixed tile boundaries', () => {
    const sourceCol = VIEWPORT_TILE_COLUMN_COUNT - 1
    const targetCol = VIEWPORT_TILE_COLUMN_COUNT
    const sourceAddress = `${indexToColumn(sourceCol)}1`
    const blockerAddress = `${indexToColumn(targetCol)}1`
    const baseInput = {
      axisSeqX: 5,
      axisSeqY: 6,
      cameraSeq: 7,
      columnWidths: {},
      dprBucket: 1,
      freezeSeq: 8,
      glyphAtlasSeq: 9,
      gridMetrics: getGridMetrics(),
      materializedAtSeq: 10,
      packetSeq: 11,
      rectSeq: 12,
      rowHeights: {},
      sheetId: 3,
      sheetOrdinal: 1,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 13,
      textSeq: 14,
      valueSeq: 15,
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT * 2 - 1,
        colStart: VIEWPORT_TILE_COLUMN_COUNT,
        rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
        rowStart: 0,
      },
    }
    const spillText = 'spills into next tile with enough width to cross boundary'

    const tile = materializeGridRenderTileV3({
      ...baseInput,
      engine: makeEngine({
        [sourceAddress]: createCellSnapshot(sourceAddress, { tag: ValueTag.String, value: spillText }),
      }),
    })
    const inboundRun = tile.textRuns.find((run) => run.col === sourceCol && run.row === 0)

    expect(inboundRun).toMatchObject({
      clipX: 0,
      row: 0,
      text: spillText,
    })
    expect(inboundRun?.clipWidth).toBeGreaterThan(0)
    expect(inboundRun?.x).toBeLessThan(0)
    expect(inboundRun?.spillColEnd).toBeGreaterThanOrEqual(targetCol)

    const blockedTile = materializeGridRenderTileV3({
      ...baseInput,
      engine: makeEngine({
        [blockerAddress]: createCellSnapshot(blockerAddress, { tag: ValueTag.Number, value: 1 }),
        [sourceAddress]: createCellSnapshot(sourceAddress, { tag: ValueTag.String, value: spillText }),
      }),
    })

    expect(blockedTile.textRuns.find((run) => run.col === sourceCol && run.row === 0)).toBeUndefined()
  })
})
