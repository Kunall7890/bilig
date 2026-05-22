import type { Page } from '@playwright/test'

export async function waitForBenchmarkCorpus(page: Page, timeoutMs = 60_000) {
  await page.waitForFunction(
    () => {
      const collector = (window as Window & { __biligScrollPerf?: { getBenchmarkState?: () => { state: string; error: string | null } } })
        .__biligScrollPerf
      const state = collector?.getBenchmarkState?.()
      return state?.state === 'ready' || state?.state === 'error'
    },
    undefined,
    { timeout: timeoutMs },
  )

  const benchmarkState = await page.evaluate(() => {
    const collector = (
      window as Window & {
        __biligScrollPerf?: {
          getBenchmarkState?: () => {
            state: string
            error: string | null
            fixture: { id: string; materializedCellCount: number; sheetName: string } | null
          }
        }
      }
    ).__biligScrollPerf
    return collector?.getBenchmarkState?.() ?? null
  })

  if (!benchmarkState) {
    throw new Error('benchmark corpus state was not available')
  }
  if (benchmarkState.state === 'error') {
    throw new Error(benchmarkState.error ?? 'benchmark corpus installation failed')
  }
  return benchmarkState
}

export async function startWorkbookScrollPerf(
  page: Page,
  workload: string,
  options: {
    readonly primeRenderer?: boolean
  } = {},
) {
  await page.bringToFront()
  await settleWorkbookScrollPerf(page, 2)
  if (options.primeRenderer ?? true) {
    await primeWorkbookGridScrollRenderer(page)
  }
  await page.evaluate((nextWorkload) => {
    ;(window as Window & { __biligScrollPerf?: { startSampling?: (workload: string) => void } }).__biligScrollPerf?.startSampling?.(
      nextWorkload,
    )
  }, workload)
}

async function primeWorkbookGridScrollRenderer(page: Page) {
  await page.getByTestId('grid-scroll-viewport').evaluate(async (element) => {
    if (!(element instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    const startLeft = element.scrollLeft
    const startTop = element.scrollTop
    element.scrollLeft = startLeft + 1
    element.scrollTop = startTop + 1
    element.dispatchEvent(new Event('scroll'))
    await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
    element.scrollLeft = startLeft
    element.scrollTop = startTop
    element.dispatchEvent(new Event('scroll'))
    await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
  })
}

export async function warmStartWorkbookScrollPerf(page: Page, workload: string, warmupFrames = 12, maxAttempts = 8) {
  const quietFrames = Math.max(warmupFrames, 96)
  const runWarmup = async (attempt: number): Promise<void> => {
    await startWorkbookScrollPerf(page, `${workload}:warmup:${String(attempt)}`, { primeRenderer: attempt === 1 })
    await settleWorkbookScrollPerf(page, warmupFrames + quietFrames + 2)
    const warmupReport = await stopWorkbookScrollPerf(page)
    if (!warmupReport) {
      throw new Error('warmup performance report was not available')
    }
    const surfaceCommits = Object.values(warmupReport.counters.surfaceCommits ?? {})
    const hasSurfaceCommitNoise = surfaceCommits.some((count) => count > 0)
    const hasTypeGpuWarmupNoise =
      warmupReport.counters.typeGpuAtlasUploadBytes > 0 ||
      warmupReport.counters.typeGpuBufferAllocations > 0 ||
      warmupReport.counters.typeGpuConfigures > 0 ||
      warmupReport.counters.typeGpuSurfaceResizes > 0 ||
      warmupReport.counters.typeGpuVertexUploadBytes > 0
    const hasRenderNoise =
      warmupReport.counters.viewportSubscriptions > 0 ||
      warmupReport.counters.fullPatches > 0 ||
      warmupReport.counters.headerPaneBuilds > 0 ||
      warmupReport.counters.reactCommits > 0 ||
      warmupReport.counters.canvasSurfaceMounts > 0 ||
      warmupReport.counters.domSurfaceMounts > 0 ||
      hasSurfaceCommitNoise ||
      hasTypeGpuWarmupNoise
    if (!hasRenderNoise) {
      return
    }
    if (attempt >= maxAttempts) {
      throw new Error(`scroll performance never reached a steady state for ${workload}`)
    }
    await runWarmup(attempt + 1)
  }
  await runWarmup(1)
  await startWorkbookScrollPerf(page, workload, { primeRenderer: false })
}

export async function resetGridScroll(page: Page, input: { left?: number; top?: number } = {}) {
  await page.getByTestId('grid-scroll-viewport').evaluate(async (element, position) => {
    if (!(element instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    element.scrollLeft = position.left ?? 0
    element.scrollTop = position.top ?? 0
    element.dispatchEvent(new Event('scroll'))
    await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
  }, input)
}

export async function settleWorkbookScrollPerf(page: Page, frames = 4) {
  await page.evaluate(async (frameCount) => {
    await Array.from({ length: frameCount }).reduce<Promise<void>>(async (previous) => {
      await previous
      await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
    }, Promise.resolve())
  }, frames)
}

export async function stopWorkbookScrollPerf(page: Page) {
  return await page.evaluate(() => {
    return (
      (
        window as Window & {
          __biligScrollPerf?: {
            stopSampling?: () => {
              workload: string
              fixture: { id: string; materializedCellCount: number; sheetName: string } | null
              samples: { frameMs: number[]; inputToDrawMs: number[]; longTasksMs: number[] }
              summary: {
                frameMs: { min: number; median: number; p95: number; p99: number; max: number }
                inputToDrawMs: { min: number; median: number; p95: number; p99: number; max: number }
                longTasksMs: { min: number; median: number; p95: number; p99: number; max: number }
              }
              counters: {
                viewportSubscriptions: number
                fullPatches: number
                damagePatches: number
                damageCells: number
                rendererTileInterestBatches: number
                rendererTileExactHits: number
                rendererTileStaleHits: number
                rendererTileMisses: number
                rendererVisibleDirtyTiles: number
                rendererWarmDirtyTiles: number
                visibleWindowChanges: number
                headerPaneBuilds: number
                reactCommits: number
                canvasSurfaceMounts: number
                domSurfaceMounts: number
                canvasPaints: Record<string, number>
                surfaceCommits: Record<string, number>
                typeGpuAtlasUploadBytes: number
                typeGpuAtlasDirtyPages: number
                typeGpuAtlasDirtyPageUploadBytes: number
                typeGpuBufferAllocationBytes: number
                typeGpuBufferAllocations: number
                typeGpuConfigures: number
                typeGpuDrawCalls: number
                typeGpuPaneDraws: number
                typeGpuSubmits: number
                typeGpuSurfaceResizes: number
                typeGpuTileMisses: number
                typeGpuUniformWriteBytes: number
                typeGpuVertexUploadBytes: number
              }
            } | null
          }
        }
      ).__biligScrollPerf?.stopSampling?.() ?? null
    )
  })
}

export async function performHorizontalGridBrowse(page: Page, input: { distancePx: number; steps?: number }) {
  await performGridBrowse(page, { deltaX: input.distancePx, deltaY: 0, ...(input.steps ? { steps: input.steps } : {}) })
}

export async function performVerticalGridBrowse(page: Page, input: { distancePx: number; steps?: number }) {
  await performGridBrowse(page, { deltaX: 0, deltaY: input.distancePx, ...(input.steps ? { steps: input.steps } : {}) })
}

export async function performDiagonalGridBrowse(page: Page, input: { deltaX: number; deltaY: number; steps?: number }) {
  await performGridBrowse(page, input)
}

async function performGridBrowse(page: Page, input: { deltaX: number; deltaY: number; steps?: number }) {
  await page.getByTestId('grid-scroll-viewport').evaluate(
    async (element, options) => {
      if (!(element instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not an HTMLDivElement')
      }
      const viewport = element
      const steps = Math.max(1, options.steps ?? 120)
      const startLeft = viewport.scrollLeft
      const startTop = viewport.scrollTop
      const advance = async (step: number): Promise<void> => {
        if (step > steps) {
          return
        }
        viewport.scrollLeft = startLeft + (options.deltaX * step) / steps
        viewport.scrollTop = startTop + (options.deltaY * step) / steps
        viewport.dispatchEvent(new Event('scroll'))
        await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()))
        await advance(step + 1)
      }
      await advance(1)
    },
    { deltaX: input.deltaX, deltaY: input.deltaY, ...(input.steps ? { steps: input.steps } : {}) },
  )
}
