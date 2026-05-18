import { describe, expect, test } from 'vitest'
import type { GpuBufferHandleV3 } from '../renderer-v3/gpu-buffer-arena.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { TypeGpuTileResourceCacheV3, resolveWorkbookTileContentBufferKeyV3 } from '../renderer-v3/typegpu-tile-buffer-pool.js'
import { hasCompleteTypeGpuBodyTileContentV3, hasDrawableTypeGpuBodyPaneFramesV3 } from '../renderer-v3/typegpu-tile-render-pass.js'

function createRenderTile(
  tileId = 101,
  overrides: {
    readonly rectCount?: number | undefined
    readonly textCount?: number | undefined
  } = {},
): GridRenderTile {
  return {
    bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
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
    rectCount: overrides.rectCount ?? 0,
    rectInstances: new Float32Array(0),
    textCount: overrides.textCount ?? 0,
    textMetrics: new Float32Array(0),
    textRuns: [],
    tileId,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 0,
      styles: 1,
      text: 1,
      values: 1,
    },
  }
}

function createPane(paneId: string, tile = createRenderTile()): WorkbookRenderTilePaneState {
  return {
    contentOffset: { x: 0, y: 0 },
    frame: { height: 240, width: 320, x: 0, y: 0 },
    generation: 1,
    paneId,
    scrollAxes: { x: true, y: true },
    surfaceSize: { height: 240, width: 320 },
    tile,
    viewport: tile.bounds,
  }
}

function createHandle(layout: GpuBufferHandleV3['layout']): GpuBufferHandleV3 {
  return {
    buffer: {},
    capacityBytes: 256,
    classId: 8,
    layout,
    usedBytes: 64,
  }
}

describe('typegpu tile render pass readiness', () => {
  test('blocks drawing when no tile panes are available', () => {
    expect(
      hasCompleteTypeGpuBodyTileContentV3({
        tilePanes: [],
        tileResources: new TypeGpuTileResourceCacheV3(),
      }),
    ).toBe(false)
  })

  test('allows drawing when there are no body panes', () => {
    const tileResources = new TypeGpuTileResourceCacheV3()

    expect(
      hasCompleteTypeGpuBodyTileContentV3({
        tilePanes: [createPane('top-body')],
        tileResources,
      }),
    ).toBe(true)
  })

  test('blocks drawing when body content is missing', () => {
    expect(
      hasCompleteTypeGpuBodyTileContentV3({
        tilePanes: [createPane('body')],
        tileResources: new TypeGpuTileResourceCacheV3(),
      }),
    ).toBe(false)
  })

  test('blocks drawing when required body buffers are not ready', () => {
    const tileResources = new TypeGpuTileResourceCacheV3()
    const pane = createPane('body', createRenderTile(101, { rectCount: 1, textCount: 1 }))
    const content = tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(pane))
    content.rectCount = 1
    content.textCount = 1

    expect(hasCompleteTypeGpuBodyTileContentV3({ tilePanes: [pane], tileResources })).toBe(false)

    Reflect.set(content, 'rectHandle', createHandle('rectInstances'))
    expect(hasCompleteTypeGpuBodyTileContentV3({ tilePanes: [pane], tileResources })).toBe(false)

    Reflect.set(content, 'textHandle', createHandle('textRuns'))
    expect(hasCompleteTypeGpuBodyTileContentV3({ tilePanes: [pane], tileResources })).toBe(true)
  })

  test('blocks drawing when cached content is stale for a text-bearing body tile', () => {
    const tileResources = new TypeGpuTileResourceCacheV3()
    const pane = createPane('body', createRenderTile(101, { textCount: 1 }))
    const content = tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(pane))
    content.textCount = 0

    expect(hasCompleteTypeGpuBodyTileContentV3({ tilePanes: [pane], tileResources })).toBe(false)
  })

  test('allows native text layer frames to draw while GPU text resources are absent', () => {
    const tileResources = new TypeGpuTileResourceCacheV3()
    const pane = createPane('body', createRenderTile(101, { rectCount: 1, textCount: 1 }))
    const content = tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(pane))
    content.rectCount = 1
    Reflect.set(content, 'rectHandle', createHandle('rectInstances'))

    expect(hasCompleteTypeGpuBodyTileContentV3({ drawText: false, tilePanes: [pane], tileResources })).toBe(true)
  })

  test('blocks drawing body panes whose content entry has no drawable instances', () => {
    const tileResources = new TypeGpuTileResourceCacheV3()
    const pane = createPane('body:0:1')
    tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(pane))

    expect(hasCompleteTypeGpuBodyTileContentV3({ tilePanes: [pane], tileResources })).toBe(false)
  })

  test('ignores hidden resident body panes when deciding whether the visible TypeGPU frame can draw', () => {
    const tileResources = new TypeGpuTileResourceCacheV3()
    const visiblePane = createPane('body', createRenderTile(101, { rectCount: 1 }))
    const hiddenPane = {
      ...createPane('body:0:1', createRenderTile(102)),
      drawVisible: false,
    }
    const visibleContent = tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(visiblePane))
    visibleContent.rectCount = 1
    Reflect.set(visibleContent, 'rectHandle', createHandle('rectInstances'))
    tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(hiddenPane))

    expect(hasCompleteTypeGpuBodyTileContentV3({ tilePanes: [visiblePane, hiddenPane], tileResources })).toBe(true)
  })

  test('ignores offscreen resident body panes when the visible frame has complete content', () => {
    const tileResources = new TypeGpuTileResourceCacheV3()
    const visiblePane = createPane('body', createRenderTile(101, { rectCount: 1 }))
    const offscreenPane = {
      ...createPane('body:offscreen', createRenderTile(102, { rectCount: 1 })),
      frame: { height: 240, width: 320, x: 800, y: 0 },
    }
    const surface = { dpr: 1, height: 360, pixelHeight: 360, pixelWidth: 640, width: 640 }
    const visibleContent = tileResources.getContent(resolveWorkbookTileContentBufferKeyV3(visiblePane))
    visibleContent.rectCount = 1
    Reflect.set(visibleContent, 'rectHandle', createHandle('rectInstances'))

    expect(hasDrawableTypeGpuBodyPaneFramesV3({ surface, tilePanes: [visiblePane, offscreenPane] })).toBe(true)
    expect(hasCompleteTypeGpuBodyTileContentV3({ surface, tilePanes: [visiblePane, offscreenPane], tileResources })).toBe(true)
  })

  test('blocks drawing when every visible body pane is outside the surface', () => {
    const surface = { dpr: 1, height: 360, pixelHeight: 360, pixelWidth: 640, width: 640 }
    const offscreenPane = {
      ...createPane('body:offscreen', createRenderTile(102, { rectCount: 1 })),
      frame: { height: 240, width: 320, x: 800, y: 0 },
    }

    expect(hasDrawableTypeGpuBodyPaneFramesV3({ surface, tilePanes: [offscreenPane] })).toBe(false)
  })
})
