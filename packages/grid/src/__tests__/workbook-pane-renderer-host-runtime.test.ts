// @vitest-environment jsdom
import { afterEach, describe, expect, test, vi } from 'vitest'
import { WorkbookPaneRendererHostRuntimeV3 } from '../renderer-v3/workbook-pane-renderer-host-runtime.js'
import { WorkbookPaneRendererRuntimeV3, type WorkbookPaneFrameDrawerV3 } from '../renderer-v3/workbook-pane-renderer-runtime.js'
import { WorkbookPaneSurfaceRuntimeV3 } from '../renderer-v3/workbook-pane-surface-runtime.js'

function createHost(width: number, height: number): HTMLDivElement {
  const host = document.createElement('div')
  Object.defineProperty(host, 'clientWidth', { configurable: true, value: width })
  Object.defineProperty(host, 'clientHeight', { configurable: true, value: height })
  return host
}

function installManualAnimationFrames(): { flushNextFrame: () => void; restore: () => void } {
  const callbacks = new Map<number, FrameRequestCallback>()
  let nextHandle = 1
  const requestFrame = vi.spyOn(window, 'requestAnimationFrame').mockImplementation((callback) => {
    const handle = nextHandle
    nextHandle += 1
    callbacks.set(handle, callback)
    return handle
  })
  const cancelFrame = vi.spyOn(window, 'cancelAnimationFrame').mockImplementation((handle) => {
    callbacks.delete(handle)
  })
  return {
    flushNextFrame: () => {
      const next = callbacks.entries().next()
      if (next.done) {
        throw new Error('no animation frame is scheduled')
      }
      const [handle, callback] = next.value
      callbacks.delete(handle)
      callback(performance.now())
    },
    restore: () => {
      requestFrame.mockRestore()
      cancelFrame.mockRestore()
    },
  }
}

describe('WorkbookPaneRendererHostRuntimeV3', () => {
  afterEach(() => {
    vi.restoreAllMocks()
  })

  test('owns the surface-to-renderer handoff outside React state', async () => {
    const animationFrames = installManualAnimationFrames()
    const backend = {}
    const createBackend = vi.fn(async () => backend)
    const destroyBackend = vi.fn()
    const syncSurface = vi.fn()
    const drawFrame = vi.fn<WorkbookPaneFrameDrawerV3>()
    const runtime = new WorkbookPaneRendererHostRuntimeV3({
      rendererRuntime: new WorkbookPaneRendererRuntimeV3(drawFrame),
      surfaceRuntime: new WorkbookPaneSurfaceRuntimeV3({
        createBackend,
        createResizeObserver: () => null,
        destroyBackend,
        getDevicePixelRatio: () => 2,
        syncSurface,
      }),
    })
    const host = createHost(640, 360)
    const canvas = document.createElement('canvas')

    runtime.updateProps({
      active: true,
      cameraStore: null,
      geometry: null,
      headerPanes: [],
      host,
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [],
      scrollTransformStore: null,
      tilePanes: [],
    })
    runtime.setCanvas(canvas)
    await Promise.resolve()
    animationFrames.flushNextFrame()

    expect(createBackend).toHaveBeenCalledWith(canvas)
    expect(syncSurface).toHaveBeenCalledWith({
      backend,
      canvas,
      size: {
        dpr: 2,
        height: 360,
        pixelHeight: 720,
        pixelWidth: 1280,
        width: 640,
      },
    })
    expect(drawFrame).toHaveBeenCalled()
    expect(drawFrame.mock.calls.at(-1)?.[0]).toMatchObject({
      backend,
      surface: {
        dpr: 2,
        height: 360,
        pixelHeight: 720,
        pixelWidth: 1280,
        width: 640,
      },
      tilePanes: [],
    })

    runtime.dispose()
    animationFrames.restore()

    expect(destroyBackend).toHaveBeenCalledWith(backend)
    expect(canvas.width).toBe(0)
    expect(canvas.height).toBe(0)
  })

  test('detaches the WebGPU surface when the pane becomes inactive', async () => {
    const backend = {}
    const destroyBackend = vi.fn()
    const runtime = new WorkbookPaneRendererHostRuntimeV3({
      rendererRuntime: new WorkbookPaneRendererRuntimeV3(vi.fn<WorkbookPaneFrameDrawerV3>()),
      surfaceRuntime: new WorkbookPaneSurfaceRuntimeV3({
        createBackend: vi.fn(async () => backend),
        createResizeObserver: () => null,
        destroyBackend,
        syncSurface: vi.fn(),
      }),
    })

    runtime.updateProps({
      active: true,
      cameraStore: null,
      geometry: null,
      headerPanes: [],
      host: createHost(640, 360),
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [],
      scrollTransformStore: null,
      tilePanes: [],
    })
    runtime.setCanvas(document.createElement('canvas'))
    await Promise.resolve()

    runtime.updateProps({
      active: false,
      cameraStore: null,
      geometry: null,
      headerPanes: [],
      host: createHost(640, 360),
      overlay: null,
      overlayBuilder: null,
      preloadTilePanes: [],
      scrollTransformStore: null,
      tilePanes: [],
    })

    expect(destroyBackend).toHaveBeenCalledWith(backend)
    runtime.dispose()
  })
})
