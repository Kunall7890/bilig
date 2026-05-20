import { describe, expect, test } from 'vitest'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'
import { createGlyphAtlas } from '../renderer-v3/typegpu-atlas-manager.js'
import { syncTypeGpuTilePaneResourcesV3, TypeGpuTileResourceCacheV3 } from '../renderer-v3/typegpu-tile-buffer-pool.js'
import type { TypeGpuRendererArtifacts } from '../renderer-v3/typegpu-primitives.js'

const TEXT_INSTANCE_BYTE_COUNT = 16 * Float32Array.BYTES_PER_ELEMENT

interface RecordedBufferWrite {
  readonly bytes: number
  readonly endOffset?: number | undefined
  readonly startOffset?: number | undefined
}

describe('typegpu v3 tile text patches', () => {
  test('skips text uploads for authoritative full-span dirties with equivalent run payloads', () => {
    const writes: RecordedBufferWrite[] = []
    const artifacts = createRecordedArtifacts(writes)
    const tileResources = new TypeGpuTileResourceCacheV3(artifacts)
    const atlas = createGlyphAtlas()
    const baseRuns = createTextRuns(Array.from({ length: 10 }, (_, index) => `${index}`))
    const baseTile = createTile({
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })

    syncTypeGpuTilePaneResourcesV3({
      atlas,
      panes: [createPane(baseTile)],
      tileResources,
    })
    writes.length = 0

    const unchangedTile = createTile({
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 0, length: baseRuns.length }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
      lastBatchId: 2,
      textCount: baseRuns.length,
      textRuns: baseRuns.map((run) => createTextRun(run)),
      version: {
        ...baseTile.version,
        text: 2,
        values: 2,
      },
    })

    syncTypeGpuTilePaneResourcesV3({
      atlas,
      panes: [createPane(unchangedTile)],
      tileResources,
    })

    expect(writes).toHaveLength(0)
  })

  test('patches changed text runs instead of full-writing authoritative full-span text dirties', () => {
    const writes: RecordedBufferWrite[] = []
    const artifacts = createRecordedArtifacts(writes)
    const tileResources = new TypeGpuTileResourceCacheV3(artifacts)
    const atlas = createGlyphAtlas()
    const baseRuns = createTextRuns(Array.from({ length: 10 }, (_, index) => `${index}`))
    const baseTile = createTile({
      textCount: baseRuns.length,
      textRuns: baseRuns,
    })

    syncTypeGpuTilePaneResourcesV3({
      atlas,
      panes: [createPane(baseTile)],
      tileResources,
    })
    writes.length = 0

    const changedRuns = baseRuns.slice()
    changedRuns[4] = createTextRun({ ...baseRuns[4], text: '4444' })
    const changedTile = createTile({
      dirty: {
        glyphSpans: [],
        rectSpans: [],
        textSpans: [{ offset: 0, length: baseRuns.length }],
      },
      dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
      lastBatchId: 2,
      textCount: changedRuns.length,
      textRuns: changedRuns,
      version: {
        ...baseTile.version,
        text: 2,
        values: 2,
      },
    })

    syncTypeGpuTilePaneResourcesV3({
      atlas,
      panes: [createPane(changedTile)],
      tileResources,
    })

    const fullPayloadBytes = changedRuns.reduce((byteCount, run) => byteCount + run.text.length * TEXT_INSTANCE_BYTE_COUNT, 0)
    const totalWriteBytes = writes.reduce((byteCount, write) => byteCount + write.bytes, 0)

    expect(writes).toHaveLength(2)
    expect(writes.every((write) => write.startOffset !== undefined)).toBe(true)
    expect(writes.some((write) => write.bytes === fullPayloadBytes && write.startOffset === undefined)).toBe(false)
    expect(totalWriteBytes).toBeLessThan(fullPayloadBytes)
  })
})

function createTile(overrides: Partial<GridRenderTile> = {}): GridRenderTile {
  const version = {
    axisX: 1,
    axisY: 1,
    freeze: 0,
    styles: 1,
    text: 1,
    values: 1,
    ...overrides.version,
  }
  return {
    bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 7,
      sheetOrdinal: 0,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 0,
    rectInstances: new Float32Array(),
    textCount: 0,
    textMetrics: new Float32Array(),
    textRuns: [],
    tileId: 101,
    version,
    ...overrides,
  }
}

function createPane(tile: GridRenderTile): WorkbookRenderTilePaneState {
  return {
    contentOffset: { x: 0, y: 0 },
    drawVisible: true,
    frame: { height: 704, width: 1024, x: 0, y: 0 },
    generation: 1,
    paneId: 'body',
    scrollAxes: { x: true, y: true },
    surfaceSize: { height: 704, width: 1024 },
    tile,
    viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
  }
}

function createTextRuns(values: readonly string[]): GridRenderTile['textRuns'] {
  return values.map((text, index) =>
    createTextRun({
      col: index,
      text,
      x: index * 104,
      clipX: index * 104,
    }),
  )
}

function createTextRun(overrides: Partial<GridRenderTile['textRuns'][number]> = {}): GridRenderTile['textRuns'][number] {
  return {
    align: 'left',
    clipHeight: 22,
    clipWidth: 104,
    clipX: 0,
    clipY: 0,
    color: '#111111',
    col: 0,
    font: '400 11px sans-serif',
    fontSize: 11,
    height: 22,
    row: 0,
    strike: false,
    text: 'A',
    underline: false,
    width: 104,
    wrap: false,
    x: 0,
    y: 0,
    ...overrides,
  }
}

function createRecordedArtifacts(writes: RecordedBufferWrite[]): TypeGpuRendererArtifacts {
  const root: TypeGpuRendererArtifacts['root'] = Object.create(null)
  root.createBuffer = () => createRecordedBuffer(writes)
  return {
    atlasHeight: 0,
    atlasTexture: null,
    atlasVersion: 0,
    atlasWidth: 0,
    context: Object.create(null),
    device: Object.create(null),
    format: 'rgba8unorm',
    quadBuffer: createRecordedBuffer(writes),
    rectPipeline: Object.create(null),
    root,
    sampler: Object.create(null),
    textPipeline: Object.create(null),
  }
}

function createRecordedBuffer(writes: RecordedBufferWrite[]): TypeGpuRendererArtifacts['quadBuffer'] {
  const buffer = Object.create(null)
  buffer.$usage = () => buffer
  buffer.destroy = () => undefined
  buffer.write = (source: ArrayBuffer, options?: { readonly startOffset?: number; readonly endOffset?: number }) => {
    writes.push({
      bytes: source.byteLength,
      endOffset: options?.endOffset,
      startOffset: options?.startOffset,
    })
  }
  return buffer
}
