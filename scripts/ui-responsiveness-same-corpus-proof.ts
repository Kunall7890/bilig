import { mkdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'

import type { Page } from '@playwright/test'

import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import { biligRenderedSurfaceReadiness } from './ui-responsiveness-same-corpus-surface.ts'
import { readBiligRenderedSurfaceState } from './ui-responsiveness-same-corpus-surface-page.ts'
import type {
  SameCorpusCaptureMeasurement,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
export const sameCorpusUiRenderProofContractVersion = 'same-corpus-ui-v2'

export interface SameCorpusScenarioProof {
  readonly biligMeanMs: number
  readonly biligP95Ms: number
  readonly googleMeanMs: number
  readonly googleP95Ms: number
  readonly microsoftExcelWebMeanMs?: number
  readonly microsoftExcelWebP95Ms?: number
  readonly meanRatio: number
  readonly p95Ratio: number
  readonly microsoftExcelWebMeanRatio?: number
  readonly microsoftExcelWebP95Ratio?: number
  readonly screenshotProof: SameCorpusScreenshotProof
  readonly pixelGridProof: SameCorpusPixelGridProof
}

export interface SameCorpusScreenshotProof {
  readonly captured: boolean
  readonly requiredProducts: readonly UiResponsivenessSameCorpusProduct[]
  readonly artifactPaths: readonly string[]
  readonly missingProducts: readonly UiResponsivenessSameCorpusProduct[]
}

export interface SameCorpusPixelGridProof {
  readonly captured: boolean
  readonly requiredProducts: readonly UiResponsivenessSameCorpusProduct[]
  readonly products: readonly SameCorpusProductPixelGridProof[]
  readonly missingProducts: readonly UiResponsivenessSameCorpusProduct[]
}

export interface SameCorpusProductVisualProof {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly screenshotPath: string | null
  readonly screenshotCaptured: boolean
  readonly pixelGridProof: SameCorpusProductPixelGridProof
}

export interface SameCorpusProductPixelGridProof {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly captured: boolean
  readonly method: 'typegpu-visible-canvas' | 'google-sheets-visible-grid' | 'excel-web-visible-grid'
  readonly viewportPixelWidth: number
  readonly viewportPixelHeight: number
  readonly evidence: readonly string[]
}

export function isSameCorpusProductPixelGridProofComplete(proof: SameCorpusProductPixelGridProof): boolean {
  if (!proof.captured || proof.viewportPixelWidth <= 0 || proof.viewportPixelHeight <= 0) {
    return false
  }
  const evidence = sameCorpusEvidenceMap(proof.evidence)
  if (!hasStrictScreenshotPixelGridEvidence(evidence)) {
    return false
  }
  if (proof.product === 'google-sheets') {
    return proof.method === 'google-sheets-visible-grid'
  }
  if (proof.product === 'microsoft-excel-web') {
    return proof.method === 'excel-web-visible-grid'
  }
  if (proof.method !== 'typegpu-visible-canvas' || proof.evidence.some((entry) => entry.startsWith('gap='))) {
    return false
  }
  const gridProjectedRevision = evidence.get('gridProjectedRevision') ?? ''
  const tileSceneRevision = evidence.get('tileSceneRevision') ?? ''
  const tilePaneCount = numericEvidence(evidence, 'tilePaneCount')
  const headerPaneCount = numericEvidence(evidence, 'headerPaneCount')
  const presentedTilePaneCount = numericEvidence(evidence, 'presentedTilePaneCount')
  const presentedHeaderPaneCount = numericEvidence(evidence, 'presentedHeaderPaneCount')
  const expectedPixelWidth = numericEvidence(evidence, 'expectedPixelWidth')
  const expectedPixelHeight = numericEvidence(evidence, 'expectedPixelHeight')
  const canvasPixelWidth = numericEvidence(evidence, 'canvasPixelWidth')
  const canvasPixelHeight = numericEvidence(evidence, 'canvasPixelHeight')
  return (
    evidence.get('mode') === 'typegpu-v3' &&
    evidence.get('contractVersion') === sameCorpusUiRenderProofContractVersion &&
    evidence.get('backendStatus') === 'ready' &&
    evidence.get('frameProofStatus') === 'presented' &&
    evidence.get('hasPresentedVisibleFrame') === 'true' &&
    isPositiveNumber(tilePaneCount) &&
    isPositiveNumber(headerPaneCount) &&
    presentedTilePaneCount === tilePaneCount &&
    presentedHeaderPaneCount === headerPaneCount &&
    evidence.get('canvasCoversViewport') === 'true' &&
    isPositiveNumber(expectedPixelWidth) &&
    isPositiveNumber(expectedPixelHeight) &&
    isPositiveNumber(canvasPixelWidth) &&
    isPositiveNumber(canvasPixelHeight) &&
    canvasPixelWidth >= expectedPixelWidth - 2 &&
    canvasPixelHeight >= expectedPixelHeight - 2 &&
    gridProjectedRevision.length > 0 &&
    evidence.get('typeGpuProjectedRevision') === gridProjectedRevision &&
    evidence.get('visibleProjectedRevision') === gridProjectedRevision &&
    tileSceneRevision.length > 0 &&
    evidence.get('visibleRenderRevision') === tileSceneRevision
  )
}

export function buildCaptureScenarioProof(args: {
  readonly bilig: SameCorpusCaptureMeasurement
  readonly googleSheets: SameCorpusCaptureMeasurement
  readonly microsoftExcelWeb?: SameCorpusCaptureMeasurement | undefined
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
}): SameCorpusScenarioProof {
  return buildScenarioProof({
    biligTiming: summarizeNumbers(primaryCaptureTimingSamples(args.bilig)),
    googleSheetsTiming: summarizeNumbers(primaryCaptureTimingSamples(args.googleSheets)),
    microsoftExcelWebTiming: args.microsoftExcelWeb ? summarizeNumbers(primaryCaptureTimingSamples(args.microsoftExcelWeb)) : null,
    visualProofs: args.visualProofs,
  })
}

export function buildScorecardScenarioProof(args: {
  readonly bilig: UiResponsivenessSameCorpusMeasurement
  readonly googleSheets: UiResponsivenessSameCorpusMeasurement
  readonly microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement | undefined
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
}): SameCorpusScenarioProof {
  return buildScenarioProof({
    biligTiming: primaryScorecardTiming(args.bilig),
    googleSheetsTiming: primaryScorecardTiming(args.googleSheets),
    microsoftExcelWebTiming: args.microsoftExcelWeb ? primaryScorecardTiming(args.microsoftExcelWeb) : null,
    visualProofs: args.visualProofs,
  })
}

export async function captureSameCorpusProductVisualProof(args: {
  readonly caseId: string
  readonly outputPath: string
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
}): Promise<SameCorpusProductVisualProof> {
  const screenshotPath = screenshotArtifactPath(args.outputPath, args.caseId, args.product, args.sampleIndex)
  mkdirSync(dirname(screenshotPath), { recursive: true })
  const screenshot = await captureProductScreenshot(args.page, args.product, screenshotPath)
  const pixelGridProof = await readProductPixelGridProof(args.page, args.product, screenshot.buffer)
  return {
    product: args.product,
    screenshotPath: screenshot.captured ? repoRelativePath(screenshotPath) : null,
    screenshotCaptured: screenshot.captured,
    pixelGridProof,
  }
}

export function validateSameCorpusScenarioProof(
  proof: SameCorpusScenarioProof,
  caseId: string,
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): void {
  const expected = buildScorecardScenarioProof({
    bilig,
    googleSheets,
    microsoftExcelWeb,
    visualProofs: proof.pixelGridProof.products.map((entry) => ({
      product: entry.product,
      screenshotPath: proof.screenshotProof.artifactPaths.find((artifact) => artifact.includes(`${entry.product}-`)) ?? null,
      screenshotCaptured: !proof.screenshotProof.missingProducts.includes(entry.product),
      pixelGridProof: entry,
    })),
  })
  if (
    proof.biligMeanMs !== expected.biligMeanMs ||
    proof.biligP95Ms !== expected.biligP95Ms ||
    proof.googleMeanMs !== expected.googleMeanMs ||
    proof.googleP95Ms !== expected.googleP95Ms ||
    proof.microsoftExcelWebMeanMs !== expected.microsoftExcelWebMeanMs ||
    proof.microsoftExcelWebP95Ms !== expected.microsoftExcelWebP95Ms ||
    proof.meanRatio !== expected.meanRatio ||
    proof.p95Ratio !== expected.p95Ratio ||
    proof.microsoftExcelWebMeanRatio !== expected.microsoftExcelWebMeanRatio ||
    proof.microsoftExcelWebP95Ratio !== expected.microsoftExcelWebP95Ratio
  ) {
    throw new Error(`UI responsiveness same-corpus scenario proof timing is stale: ${caseId}`)
  }
  if (
    proof.screenshotProof.captured !== expected.screenshotProof.captured ||
    JSON.stringify(proof.screenshotProof.requiredProducts) !== JSON.stringify(expected.screenshotProof.requiredProducts) ||
    JSON.stringify(proof.screenshotProof.missingProducts) !== JSON.stringify(expected.screenshotProof.missingProducts)
  ) {
    throw new Error(`UI responsiveness same-corpus screenshot proof is stale: ${caseId}`)
  }
  if (
    proof.pixelGridProof.captured !== expected.pixelGridProof.captured ||
    JSON.stringify(proof.pixelGridProof.requiredProducts) !== JSON.stringify(expected.pixelGridProof.requiredProducts) ||
    (proof.pixelGridProof.captured &&
      JSON.stringify(proof.pixelGridProof.missingProducts) !== JSON.stringify(expected.pixelGridProof.missingProducts))
  ) {
    throw new Error(`UI responsiveness same-corpus pixel grid proof is stale: ${caseId}`)
  }
  if (!proof.screenshotProof.captured) {
    throw new Error(`UI responsiveness same-corpus scenario proof is missing screenshot proof: ${caseId}`)
  }
}

function buildScenarioProof(args: {
  readonly biligTiming: NumericSummary
  readonly googleSheetsTiming: NumericSummary
  readonly microsoftExcelWebTiming: NumericSummary | null
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
}): SameCorpusScenarioProof {
  const requiredProducts = ['bilig', 'google-sheets'] as const satisfies readonly UiResponsivenessSameCorpusProduct[]
  const screenshotProducts = new Set(
    args.visualProofs.filter((entry) => entry.screenshotCaptured && entry.screenshotPath).map((entry) => entry.product),
  )
  const pixelProducts = new Set(
    args.visualProofs.filter((entry) => isSameCorpusProductPixelGridProofComplete(entry.pixelGridProof)).map((entry) => entry.product),
  )
  return {
    biligMeanMs: args.biligTiming.mean,
    biligP95Ms: args.biligTiming.p95,
    googleMeanMs: args.googleSheetsTiming.mean,
    googleP95Ms: args.googleSheetsTiming.p95,
    ...(args.microsoftExcelWebTiming
      ? {
          microsoftExcelWebMeanMs: args.microsoftExcelWebTiming.mean,
          microsoftExcelWebP95Ms: args.microsoftExcelWebTiming.p95,
          microsoftExcelWebMeanRatio: ratio(args.biligTiming.mean, args.microsoftExcelWebTiming.mean),
          microsoftExcelWebP95Ratio: ratio(args.biligTiming.p95, args.microsoftExcelWebTiming.p95),
        }
      : {}),
    meanRatio: ratio(args.biligTiming.mean, args.googleSheetsTiming.mean),
    p95Ratio: ratio(args.biligTiming.p95, args.googleSheetsTiming.p95),
    screenshotProof: {
      captured: requiredProducts.every((product) => screenshotProducts.has(product)),
      requiredProducts,
      artifactPaths: args.visualProofs.flatMap((entry) => (entry.screenshotPath ? [entry.screenshotPath] : [])),
      missingProducts: requiredProducts.filter((product) => !screenshotProducts.has(product)),
    },
    pixelGridProof: {
      captured: requiredProducts.every((product) => pixelProducts.has(product)),
      requiredProducts,
      products: args.visualProofs.map((entry) => entry.pixelGridProof),
      missingProducts: requiredProducts.filter((product) => !pixelProducts.has(product)),
    },
  }
}

function sameCorpusEvidenceMap(evidence: readonly string[]): ReadonlyMap<string, string> {
  return new Map(
    evidence.flatMap((entry) => {
      const separatorIndex = entry.indexOf('=')
      return separatorIndex > 0 ? [[entry.slice(0, separatorIndex), entry.slice(separatorIndex + 1)]] : []
    }),
  )
}

function numericEvidence(evidence: ReadonlyMap<string, string>, key: string): number | null {
  const value = evidence.get(key)
  if (value === undefined) {
    return null
  }
  const numericValue = Number(value)
  return Number.isFinite(numericValue) ? numericValue : null
}

function isPositiveNumber(value: number | null): value is number {
  return value !== null && value > 0
}

function minimumEvidenceNumber(evidence: ReadonlyMap<string, string>, key: string, minimum: number): boolean {
  const value = numericEvidence(evidence, key)
  return value !== null && value >= minimum
}

function hasStrictScreenshotPixelGridEvidence(evidence: ReadonlyMap<string, string>): boolean {
  return (
    evidence.get('pixelGridProofVersion') === 'grid-pixels-v1' &&
    evidence.get('pixelSampleSource') === 'screenshot' &&
    minimumEvidenceNumber(evidence, 'visibleGridLinePixels', 24) &&
    minimumEvidenceNumber(evidence, 'verticalLineRuns', 3) &&
    minimumEvidenceNumber(evidence, 'horizontalLineRuns', 3)
  )
}

function primaryCaptureTimingSamples(measurement: SameCorpusCaptureMeasurement): readonly number[] {
  return measurement.scrollEventResponseMsSamples ?? measurement.operationResponseMsSamples
}

function primaryScorecardTiming(measurement: UiResponsivenessSameCorpusMeasurement): NumericSummary {
  return measurement.scrollEventResponseMs ?? measurement.operationResponseMs
}

function ratio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}

interface ScreenshotBuffer {
  toString(encoding: 'base64'): string
}

interface CapturedProductScreenshot {
  readonly buffer: ScreenshotBuffer | null
  readonly captured: boolean
}

interface ScreenshotGridPixelAnalysis {
  readonly height: number
  readonly horizontalLineRuns: number
  readonly nonBlankPixels: number
  readonly verticalLineRuns: number
  readonly visibleGridLinePixels: number
  readonly width: number
}

function screenshotArtifactPath(
  outputPath: string,
  caseId: string,
  product: UiResponsivenessSameCorpusProduct,
  sampleIndex: number,
): string {
  return resolve(`${outputPath}.proof`, caseId, `${product}-sample-${String(sampleIndex + 1)}.png`)
}

function repoRelativePath(path: string): string {
  return relative(rootDir, path)
}

async function captureProductScreenshot(
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

async function readProductPixelGridProof(
  page: Page,
  product: UiResponsivenessSameCorpusProduct,
  screenshotBuffer: ScreenshotBuffer | null,
): Promise<SameCorpusProductPixelGridProof> {
  if (product === 'bilig') {
    return await readBiligPixelGridProof(page, screenshotBuffer)
  }
  if (!screenshotBuffer) {
    return emptyPixelGridProof(product, product === 'google-sheets' ? 'google-sheets-visible-grid' : 'excel-web-visible-grid')
  }
  if (product === 'google-sheets') {
    return await readScreenshotPixelGridProof(page, product, 'google-sheets-visible-grid', screenshotBuffer)
  }
  return await readScreenshotPixelGridProof(page, product, 'excel-web-visible-grid', screenshotBuffer)
}

async function readBiligPixelGridProof(page: Page, screenshotBuffer: ScreenshotBuffer | null): Promise<SameCorpusProductPixelGridProof> {
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
    evidence: [`contractVersion=${sameCorpusUiRenderProofContractVersion}`, ...readiness.evidence, ...pixelGridEvidence(pixelAnalysis)],
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
      const vertical = countThinLineRuns(columnScores, Math.max(8, Math.floor(canvas.height * 0.18)))
      const horizontal = countThinLineRuns(rowScores, Math.max(8, Math.floor(canvas.width * 0.18)))
      return {
        height: canvas.height,
        horizontalLineRuns: horizontal.runs,
        nonBlankPixels,
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
    analysis.horizontalLineRuns >= 3
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
  ]
}
