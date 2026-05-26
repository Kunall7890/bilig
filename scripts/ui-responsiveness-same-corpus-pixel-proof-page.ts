import type { Page } from '@playwright/test'

import { biligRenderedSurfaceReadiness } from './ui-responsiveness-same-corpus-surface.ts'
import { readBiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface-page.ts'
import type { SameCorpusProductPixelGridProof } from './ui-responsiveness-same-corpus-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-proof.ts'

interface ScreenshotBuffer {
  toString(encoding: 'base64'): string
}

interface CapturedProductScreenshot {
  readonly buffer: ScreenshotBuffer | null
  readonly captured: boolean
}

interface ScreenshotGridPixelAnalysis {
  readonly height: number
  readonly horizontalLineCoverageBands: number
  readonly horizontalLineRuns: number
  readonly largestHorizontalLineGapPx: number
  readonly largestVerticalLineGapPx: number
  readonly nonBlankPixels: number
  readonly verticalLineCoverageBands: number
  readonly verticalLineRuns: number
  readonly visibleGridLinePixels: number
  readonly width: number
}

export async function captureProductScreenshot(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  path: string,
): Promise<CapturedProductScreenshot> {
  const selector =
    product === 'bilig' ? '[data-testid="sheet-grid"]' : product === 'google-sheets' ? '.grid-scrollable-wrapper' : '.ewr-grdcontarea-grid'
  if (product === 'microsoft-excel-web') {
    const candidates = await Promise.all(
      page.frames().map(async (frame) => {
        const locator = frame.locator(selector).first()
        return { locator, count: await locator.count().catch(() => 0) }
      }),
    )
    const candidate = candidates.find((entry) => entry.count > 0)
    if (candidate) {
      return {
        buffer: await candidate.locator.screenshot({ path }),
        captured: true,
      }
    }
    return { buffer: null, captured: false }
  }
  const locator = page.locator(selector).first()
  if ((await locator.count().catch(() => 0)) === 0) {
    return { buffer: null, captured: false }
  }
  return {
    buffer: await locator.screenshot({ path }),
    captured: true,
  }
}

export async function readProductPixelGridProof(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  screenshotBuffer: ScreenshotBuffer | null,
  contractVersion: string,
): Promise<SameCorpusProductPixelGridProof> {
  if (product === 'bilig') {
    return await readBiligPixelGridProof(page, screenshotBuffer, contractVersion)
  }
  if (!screenshotBuffer) {
    return emptyPixelGridProof(product, product === 'google-sheets' ? 'google-sheets-visible-grid' : 'excel-web-visible-grid')
  }
  if (product === 'google-sheets') {
    return await readScreenshotPixelGridProof(page, product, 'google-sheets-visible-grid', screenshotBuffer)
  }
  return await readScreenshotPixelGridProof(page, product, 'excel-web-visible-grid', screenshotBuffer)
}

async function readBiligPixelGridProof(
  page: Page,
  screenshotBuffer: ScreenshotBuffer | null,
  contractVersion: string,
): Promise<SameCorpusProductPixelGridProof> {
  const renderedSurface = await readBiligRenderedSurfaceState(page)
  const readiness = biligRenderedSurfaceReadiness(renderedSurface)
  const pixelAnalysis = screenshotBuffer ? await analyzeScreenshotGridPixels(page, screenshotBuffer) : null
  const captured = readiness.ready && Boolean(pixelAnalysis) && isScreenshotGridPixelAnalysisComplete(pixelAnalysis)
  return {
    product: 'bilig',
    captured,
    method: 'typegpu-visible-canvas',
    viewportPixelWidth: pixelAnalysis?.width ?? renderedSurface?.typeGpu?.pixelWidth ?? 0,
    viewportPixelHeight: pixelAnalysis?.height ?? renderedSurface?.typeGpu?.pixelHeight ?? 0,
    evidence: [`contractVersion=${contractVersion}`, ...readiness.evidence, ...pixelGridEvidence(pixelAnalysis)],
  }
}

async function readScreenshotPixelGridProof(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  method: SameCorpusProductPixelGridProof['method'],
  screenshotBuffer: ScreenshotBuffer,
): Promise<SameCorpusProductPixelGridProof> {
  const analysis = await analyzeScreenshotGridPixels(page, screenshotBuffer)
  return {
    product,
    captured: isScreenshotGridPixelAnalysisComplete(analysis),
    method,
    viewportPixelWidth: analysis.width,
    viewportPixelHeight: analysis.height,
    evidence: pixelGridEvidence(analysis),
  }
}

function emptyPixelGridProof(
  product: UiResponsivenessSameCorpusProduct,
  method: SameCorpusProductPixelGridProof['method'],
): SameCorpusProductPixelGridProof {
  return {
    product,
    captured: false,
    method,
    viewportPixelWidth: 0,
    viewportPixelHeight: 0,
    evidence: ['visible workbook grid surface was not found'],
  }
}

async function analyzeScreenshotGridPixels(page: Page, screenshotBuffer: ScreenshotBuffer): Promise<ScreenshotGridPixelAnalysis> {
  return await page.evaluate(
    async ({ dataUrl }) => {
      const image = await new Promise<HTMLImageElement>((resolveImage, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolveImage(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode same-corpus grid proof screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d', { willReadFrequently: true })
      if (!context) {
        throw new Error('Missing 2d context for same-corpus grid proof screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixelData = context.getImageData(0, 0, canvas.width, canvas.height).data
      const columnScores = Array.from({ length: canvas.width }, () => 0)
      const rowScores = Array.from({ length: canvas.height }, () => 0)
      const minimumAlpha = 180
      const maximumThinLineRunWidth = 4
      const isGridLinePixel = (red: number, green: number, blue: number, alpha: number): boolean => {
        if (alpha <= minimumAlpha) {
          return false
        }
        const max = Math.max(red, green, blue)
        const min = Math.min(red, green, blue)
        const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue
        return luminance >= 125 && luminance <= 246 && max - min <= 58
      }
      const countThinLineRuns = (scores: readonly number[], threshold: number): { readonly pixels: number; readonly runs: number } => {
        let linePixels = 0
        let runs = 0
        let runLength = 0
        let runPixels = 0
        const flush = () => {
          if (runLength > 0 && runLength <= maximumThinLineRunWidth) {
            runs += 1
            linePixels += runPixels
          }
          runLength = 0
          runPixels = 0
        }
        for (const score of scores) {
          if (score >= threshold) {
            runLength += 1
            runPixels += score
          } else {
            flush()
          }
        }
        flush()
        return { pixels: linePixels, runs }
      }
      // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright evaluates this helper inside the browser context.
      const countCoverageBands = (lineFlags: readonly boolean[], bandCount: number): number => {
        let coveredBands = 0
        for (let bandIndex = 0; bandIndex < bandCount; bandIndex += 1) {
          const start = Math.floor((lineFlags.length * bandIndex) / bandCount)
          const end = Math.floor((lineFlags.length * (bandIndex + 1)) / bandCount)
          if (lineFlags.slice(start, Math.max(start + 1, end)).some(Boolean)) {
            coveredBands += 1
          }
        }
        return coveredBands
      }
      // oxlint-disable-next-line eslint-plugin-unicorn(consistent-function-scoping) -- Playwright evaluates this helper inside the browser context.
      const largestGap = (lineFlags: readonly boolean[]): number => {
        let largest = 0
        let current = 0
        for (const hasLine of lineFlags) {
          if (hasLine) {
            largest = Math.max(largest, current)
            current = 0
          } else {
            current += 1
          }
        }
        return Math.max(largest, current)
      }
      let nonBlankPixels = 0
      for (let y = 0; y < canvas.height; y += 1) {
        for (let x = 0; x < canvas.width; x += 1) {
          const index = (y * canvas.width + x) * 4
          const red = pixelData[index] ?? 255
          const green = pixelData[index + 1] ?? 255
          const blue = pixelData[index + 2] ?? 255
          const alpha = pixelData[index + 3] ?? 0
          if (alpha > minimumAlpha && (red < 248 || green < 248 || blue < 248)) {
            nonBlankPixels += 1
          }
          if (isGridLinePixel(red, green, blue, alpha)) {
            columnScores[x] = (columnScores[x] ?? 0) + 1
            rowScores[y] = (rowScores[y] ?? 0) + 1
          }
        }
      }
      const verticalThreshold = Math.max(8, Math.floor(canvas.height * 0.18))
      const horizontalThreshold = Math.max(8, Math.floor(canvas.width * 0.18))
      const verticalLineColumns = columnScores.map((score) => score >= verticalThreshold)
      const horizontalLineRows = rowScores.map((score) => score >= horizontalThreshold)
      const vertical = countThinLineRuns(columnScores, verticalThreshold)
      const horizontal = countThinLineRuns(rowScores, horizontalThreshold)
      return {
        height: canvas.height,
        horizontalLineCoverageBands: countCoverageBands(horizontalLineRows, 6),
        horizontalLineRuns: horizontal.runs,
        largestHorizontalLineGapPx: largestGap(horizontalLineRows),
        largestVerticalLineGapPx: largestGap(verticalLineColumns),
        nonBlankPixels,
        verticalLineCoverageBands: countCoverageBands(verticalLineColumns, 6),
        verticalLineRuns: vertical.runs,
        visibleGridLinePixels: vertical.pixels + horizontal.pixels,
        width: canvas.width,
      }
    },
    { dataUrl: `data:image/png;base64,${screenshotBuffer.toString('base64')}` },
  )
}

function isScreenshotGridPixelAnalysisComplete(analysis: ScreenshotGridPixelAnalysis): boolean {
  return (
    analysis.width > 128 &&
    analysis.height > 128 &&
    analysis.nonBlankPixels > 128 &&
    analysis.visibleGridLinePixels >= 24 &&
    analysis.verticalLineRuns >= 3 &&
    analysis.horizontalLineRuns >= 3 &&
    analysis.verticalLineCoverageBands >= 5 &&
    analysis.horizontalLineCoverageBands >= 5 &&
    analysis.largestVerticalLineGapPx <= 384 &&
    analysis.largestHorizontalLineGapPx <= 160
  )
}

function pixelGridEvidence(analysis: ScreenshotGridPixelAnalysis | null): string[] {
  if (!analysis) {
    return ['pixelGridProofVersion=grid-pixels-v1', 'pixelSampleSource=screenshot', 'gap=missing grid proof screenshot']
  }
  return [
    'pixelGridProofVersion=grid-pixels-v1',
    'pixelSampleSource=screenshot',
    `screenshotPixelWidth=${String(analysis.width)}`,
    `screenshotPixelHeight=${String(analysis.height)}`,
    `nonBlankPixels=${String(analysis.nonBlankPixels)}`,
    `visibleGridLinePixels=${String(analysis.visibleGridLinePixels)}`,
    `verticalLineRuns=${String(analysis.verticalLineRuns)}`,
    `horizontalLineRuns=${String(analysis.horizontalLineRuns)}`,
    `verticalLineCoverageBands=${String(analysis.verticalLineCoverageBands)}`,
    `horizontalLineCoverageBands=${String(analysis.horizontalLineCoverageBands)}`,
    `largestVerticalLineGapPx=${String(analysis.largestVerticalLineGapPx)}`,
    `largestHorizontalLineGapPx=${String(analysis.largestHorizontalLineGapPx)}`,
  ]
}
