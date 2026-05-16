import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { GridVisibleTextRefreshCache } from '../runtime/gridVisibleTextRefreshCache.js'

describe('GridVisibleTextRefreshCache', () => {
  it('rejects unanchored remote text runs instead of letting ghost text render', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 1,
      textRuns: [
        {
          ...createTextRun({ text: 'ghost text' }),
          col: undefined,
          row: undefined,
        },
      ],
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine() }))).toBe(true)
  })

  it('rejects inconsistent remote text run counts', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 2,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' })],
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({ A1: 'A1' }) }))).toBe(true)
  })

  it('accepts remote text only when the visible cell text matches the workbook state', () => {
    const cache = new GridVisibleTextRefreshCache()
    const tile = createTile({
      textCount: 1,
      textRuns: [createTextRun({ col: 0, row: 0, text: 'A1' })],
    })

    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({ A1: 'A1' }) }))).toBe(false)
    expect(cache.needsLocalRefresh(tile.tileId, tile, createInput({ engine: createEngine({ A1: 'fresh A1' }) }))).toBe(true)
  })
})

function createInput(overrides: Partial<Parameters<GridVisibleTextRefreshCache['needsLocalRefresh']>[2]> = {}) {
  return {
    engine: createEngine(),
    sceneRevision: 1,
    sheetName: 'Sheet1',
    visibleViewport: { colEnd: 2, colStart: 0, rowEnd: 2, rowStart: 0 },
    ...overrides,
  }
}

function createEngine(values: Record<string, string> = {}): GridEngineLike {
  return {
    getCell: (_sheetName, address) => createCellSnapshot(address, values[address] ?? ''),
    getCellStyle: () => undefined,
    getRenderRevisionSnapshot: () => ({
      authoritativeRevision: 1,
      localRevision: 1,
      projectedRevision: 1,
      tileSceneCameraSeq: 1,
      tileSceneRevision: 1,
    }),
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => undefined,
    },
  }
}

function createCellSnapshot(address: string, value: string): CellSnapshot {
  if (value.length === 0) {
    return {
      address,
      flags: 0,
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Empty },
      version: 1,
    }
  }
  return {
    address,
    flags: 0,
    input: value,
    sheetName: 'Sheet1',
    value: { tag: ValueTag.String, value, stringId: 1 },
    version: 1,
  }
}

function createTile(overrides: Partial<GridRenderTile> = {}): GridRenderTile {
  return {
    bounds: { colEnd: 2, colStart: 0, rowEnd: 2, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 1,
      sheetOrdinal: 1,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 0,
    rectInstances: new Float32Array(),
    textCount: 0,
    textMetrics: new Float32Array(),
    textRuns: [],
    tileId: 1,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 1,
      styles: 1,
      text: 1,
      values: 1,
    },
    ...overrides,
  }
}

function createTextRun(overrides: Partial<GridRenderTile['textRuns'][number]> = {}): GridRenderTile['textRuns'][number] {
  return {
    clipHeight: 20,
    clipWidth: 100,
    clipX: 0,
    clipY: 0,
    col: 0,
    color: '#111827',
    font: '400 12px Arial',
    fontSize: 12,
    height: 20,
    row: 0,
    strike: false,
    text: '',
    underline: false,
    width: 100,
    x: 0,
    y: 0,
    ...overrides,
  }
}
