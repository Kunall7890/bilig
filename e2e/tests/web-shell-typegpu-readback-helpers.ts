import type { Page } from '@playwright/test'

export async function installTypeGpuCellReadbackHarness(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globalWindow = window as Window & {
      __biligCellReadbackHarnessInstalled?: boolean
      __biligCellReadbackState?: {
        bgra: Uint8Array
        bytesPerRow: number
        height: number
        ready: boolean
        sequence: number
        width: number
      }
      __biligCellReadbackInspector?: {
        readonly countBluePixels: (region: { readonly x0: number; readonly y0: number; readonly x1: number; readonly y1: number }) => number
        readonly countGreenPixels: (region: {
          readonly x0: number
          readonly y0: number
          readonly x1: number
          readonly y1: number
        }) => number
        readonly getSequence: () => number
        readonly isReady: () => boolean
      }
    }
    const readbackState = globalWindow.__biligCellReadbackState ?? {
      bgra: new Uint8Array(0),
      bytesPerRow: 0,
      height: 0,
      ready: false,
      sequence: 0,
      width: 0,
    }
    globalWindow.__biligCellReadbackState = readbackState
    readbackState.bgra = new Uint8Array(0)
    readbackState.bytesPerRow = 0
    readbackState.height = 0
    readbackState.ready = false
    readbackState.sequence = 0
    readbackState.width = 0

    const countPixels = (
      region: { readonly x0: number; readonly y0: number; readonly x1: number; readonly y1: number },
      targetColor: 'blue' | 'green',
    ) => {
      if (!readbackState.ready) {
        return 0
      }
      let count = 0
      const x0 = Math.max(0, Math.floor(region.x0))
      const y0 = Math.max(0, Math.floor(region.y0))
      const x1 = Math.min(readbackState.width, Math.ceil(region.x1))
      const y1 = Math.min(readbackState.height, Math.ceil(region.y1))
      for (let y = y0; y < y1; y += 1) {
        for (let x = x0; x < x1; x += 1) {
          const offset = y * readbackState.bytesPerRow + x * 4
          const blue = readbackState.bgra[offset + 0] ?? 255
          const green = readbackState.bgra[offset + 1] ?? 0
          const red = readbackState.bgra[offset + 2] ?? 255
          const alpha = readbackState.bgra[offset + 3] ?? 0
          if (targetColor === 'green' && alpha > 220 && red < 110 && green > 135 && blue < 120) {
            count += 1
          }
          if (targetColor === 'blue' && alpha > 220 && red < 120 && green < 120 && blue > 135) {
            count += 1
          }
        }
      }
      return count
    }

    globalWindow.__biligCellReadbackInspector = {
      countBluePixels(region) {
        return countPixels(region, 'blue')
      },
      countGreenPixels(region) {
        return countPixels(region, 'green')
      },
      getSequence() {
        return readbackState.sequence
      },
      isReady() {
        return readbackState.ready
      },
    }

    if (globalWindow.__biligCellReadbackHarnessInstalled || !navigator.gpu) {
      return
    }
    globalWindow.__biligCellReadbackHarnessInstalled = true

    const functionKind = 'function'
    const isCanvasContextConfigure = (value: unknown): value is (this: GPUCanvasContext, descriptor: GPUCanvasConfiguration) => void =>
      typeof value === functionKind
    const isCanvasContextGetCurrentTexture = (value: unknown): value is (this: GPUCanvasContext) => GPUTexture =>
      typeof value === functionKind
    const originalConfigure = Object.getOwnPropertyDescriptor(GPUCanvasContext.prototype, 'configure')?.value
    if (!isCanvasContextConfigure(originalConfigure)) {
      return
    }
    GPUCanvasContext.prototype.configure = function configureCellReadback(descriptor: GPUCanvasConfiguration) {
      return originalConfigure.call(this, {
        ...descriptor,
        usage: (descriptor.usage ?? GPUTextureUsage.RENDER_ATTACHMENT) | GPUTextureUsage.COPY_SRC,
      })
    }

    const originalRequestAdapter = navigator.gpu.requestAdapter.bind(navigator.gpu)
    navigator.gpu.requestAdapter = async (...adapterArgs) => {
      const adapter = await originalRequestAdapter(...adapterArgs)
      if (!adapter) {
        return adapter
      }
      const originalRequestDevice = adapter.requestDevice.bind(adapter)
      adapter.requestDevice = async (...deviceArgs) => {
        const device = await originalRequestDevice(...deviceArgs)
        const originalGetCurrentTexture = Object.getOwnPropertyDescriptor(GPUCanvasContext.prototype, 'getCurrentTexture')?.value
        if (!isCanvasContextGetCurrentTexture(originalGetCurrentTexture)) {
          return device
        }
        let lastTexture: GPUTexture | null = null
        let lastWidth = 0
        let lastHeight = 0
        GPUCanvasContext.prototype.getCurrentTexture = function recordCellReadbackTexture() {
          const texture = originalGetCurrentTexture.call(this)
          if (this.canvas instanceof HTMLCanvasElement && this.canvas.getAttribute('data-testid') === 'grid-pane-renderer') {
            lastTexture = texture
            lastWidth = this.canvas.width
            lastHeight = this.canvas.height
          }
          return texture
        }

        const originalSubmit = device.queue.submit.bind(device.queue)
        device.queue.submit = (buffers: Iterable<GPUCommandBuffer>) => {
          const commandBuffers = Array.from(buffers)
          if (lastTexture && lastWidth > 0 && lastHeight > 0) {
            const bytesPerRow = Math.ceil((lastWidth * 4) / 256) * 256
            const buffer = device.createBuffer({
              size: bytesPerRow * lastHeight,
              usage: GPUBufferUsage.COPY_DST | GPUBufferUsage.MAP_READ,
            })
            const encoder = device.createCommandEncoder()
            encoder.copyTextureToBuffer(
              { texture: lastTexture },
              { buffer, bytesPerRow, rowsPerImage: lastHeight },
              { width: lastWidth, height: lastHeight, depthOrArrayLayers: 1 },
            )
            const result = originalSubmit([...commandBuffers, encoder.finish()])
            void buffer
              .mapAsync(GPUMapMode.READ)
              .then(() => {
                readbackState.bgra = new Uint8Array(buffer.getMappedRange()).slice()
                readbackState.bytesPerRow = bytesPerRow
                readbackState.height = lastHeight
                readbackState.ready = true
                readbackState.sequence += 1
                readbackState.width = lastWidth
                return readbackState.sequence
              })
              .finally(() => {
                try {
                  buffer.unmap()
                } catch {
                  // The browser may already have released the mapped range on teardown.
                }
                buffer.destroy()
              })
            return result
          }
          return originalSubmit(commandBuffers)
        }
        return device
      }
      return adapter
    }
  })
}

export async function countGreenFillReadbackPixelsInCell(page: Page, columnIndex: number, rowIndex: number): Promise<number> {
  return await countFillReadbackPixelsInCell(page, columnIndex, rowIndex, 'green')
}

export async function countBlueFillReadbackPixelsInCell(page: Page, columnIndex: number, rowIndex: number): Promise<number> {
  return await countFillReadbackPixelsInCell(page, columnIndex, rowIndex, 'blue')
}

async function countFillReadbackPixelsInCell(page: Page, columnIndex: number, rowIndex: number, color: 'blue' | 'green'): Promise<number> {
  const [columnLeft, columnWidth, rowTop, rowHeight, scroll, canvas] = await Promise.all([
    getProductColumnLeft(page, columnIndex),
    getProductColumnWidth(page, columnIndex),
    getProductRowTop(page, rowIndex),
    getProductRowHeight(page, rowIndex),
    page.getByTestId('grid-scroll-viewport').evaluate((node) => ({
      scrollLeft: node.scrollLeft,
      scrollTop: node.scrollTop,
    })),
    page.getByTestId('grid-pane-renderer').evaluate((node) => {
      if (!(node instanceof HTMLCanvasElement)) {
        throw new Error('TypeGPU renderer is not a canvas')
      }
      return {
        clientHeight: node.clientHeight,
        clientWidth: node.clientWidth,
        height: node.height,
        width: node.width,
      }
    }),
  ])
  const scaleX = canvas.clientWidth > 0 ? canvas.width / canvas.clientWidth : 1
  const scaleY = canvas.clientHeight > 0 ? canvas.height / canvas.clientHeight : 1
  return await page.evaluate(
    ({ region, targetColor }) => {
      const inspector = (
        window as Window & {
          __biligCellReadbackInspector?: {
            readonly countBluePixels: (region: {
              readonly x0: number
              readonly y0: number
              readonly x1: number
              readonly y1: number
            }) => number
            readonly countGreenPixels: (region: {
              readonly x0: number
              readonly y0: number
              readonly x1: number
              readonly y1: number
            }) => number
            readonly isReady: () => boolean
          }
        }
      ).__biligCellReadbackInspector
      if (!inspector?.isReady()) {
        return 0
      }
      return targetColor === 'green' ? inspector.countGreenPixels(region) : inspector.countBluePixels(region)
    },
    {
      region: {
        x0: (columnLeft - scroll.scrollLeft + 8) * scaleX,
        x1: (columnLeft - scroll.scrollLeft + columnWidth - 8) * scaleX,
        y0: (PRODUCT_HEADER_HEIGHT + rowTop - scroll.scrollTop + 5) * scaleY,
        y1: (PRODUCT_HEADER_HEIGHT + rowTop - scroll.scrollTop + rowHeight - 5) * scaleY,
      },
      targetColor: color,
    },
  )
}
