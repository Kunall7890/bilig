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
export const sameCorpusUiRenderProofContractVersion = 'same-corpus-ui-v3'

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
  readonly productVerdicts: readonly SameCorpusProductPixelGridProofVerdict[]
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

export type SameCorpusProductPixelGridProofEvidenceStatus = 'current-contract' | 'legacy-insufficient' | 'missing' | 'invalid'

export interface SameCorpusProductPixelGridProofVerdict {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly evidenceStatus: SameCorpusProductPixelGridProofEvidenceStatus
  readonly acceptedForCurrentScorecard: boolean
  readonly contractVersion: string | null
  readonly requiredContractVersion: typeof sameCorpusUiRenderProofContractVersion
  readonly invalidReasons: readonly string[]
}

export function isSameCorpusProductPixelGridProofComplete(proof: SameCorpusProductPixelGridProof): boolean {
  return validateSameCorpusProductPixelGridProof(proof).acceptedForCurrentScorecard
}

export function validateSameCorpusProductPixelGridProof(proof: SameCorpusProductPixelGridProof): SameCorpusProductPixelGridProofVerdict {
  const evidence = sameCorpusEvidenceMap(proof.evidence)
  const invalidReasons = sameCorpusProductPixelGridInvalidReasons(proof, evidence)
  const acceptedForCurrentScorecard = invalidReasons.length === 0
  return {
    product: proof.product,
    evidenceStatus: acceptedForCurrentScorecard
      ? 'current-contract'
      : isLegacyInsufficientPixelProof(proof, evidence)
        ? 'legacy-insufficient'
        : proof.captured
          ? 'invalid'
          : 'missing',
    acceptedForCurrentScorecard,
    contractVersion: evidence.get('contractVersion') ?? null,
    requiredContractVersion: sameCorpusUiRenderProofContractVersion,
    invalidReasons,
  }
}

export function validateBiligVisibleFrameProof(proof: SameCorpusProductPixelGridProof): SameCorpusProductPixelGridProofVerdict {
  if (proof.product !== 'bilig') {
    return {
      product: proof.product,
      evidenceStatus: 'invalid',
      acceptedForCurrentScorecard: false,
      contractVersion: sameCorpusEvidenceMap(proof.evidence).get('contractVersion') ?? null,
      requiredContractVersion: sameCorpusUiRenderProofContractVersion,
      invalidReasons: ['Bilig visible-frame proof validator received a non-Bilig product proof'],
    }
  }
  return validateSameCorpusProductPixelGridProof(proof)
}

function sameCorpusProductPixelGridInvalidReasons(proof: SameCorpusProductPixelGridProof, evidence: ReadonlyMap<string, string>): string[] {
  const invalidReasons: string[] = []
  if (!proof.captured) {
    invalidReasons.push('pixel proof is not marked captured')
  }
  if (proof.viewportPixelWidth <= 0 || proof.viewportPixelHeight <= 0) {
    invalidReasons.push('viewport pixel dimensions are missing')
  }
  if (!hasStrictScreenshotPixelGridEvidence(evidence)) {
    invalidReasons.push('missing strict screenshot grid-pixel proof')
  }
  if (proof.product === 'google-sheets') {
    if (proof.method !== 'google-sheets-visible-grid') {
      invalidReasons.push('Google Sheets proof method is not google-sheets-visible-grid')
    }
    return invalidReasons
  }
  if (proof.product === 'microsoft-excel-web') {
    if (proof.method !== 'excel-web-visible-grid') {
      invalidReasons.push('Excel Web proof method is not excel-web-visible-grid')
    }
    return invalidReasons
  }
  invalidReasons.push(...biligVisibleFrameInvalidReasons(proof, evidence))
  return invalidReasons
}

function biligVisibleFrameInvalidReasons(proof: SameCorpusProductPixelGridProof, evidence: ReadonlyMap<string, string>): string[] {
  const invalidReasons: string[] = []
  if (proof.method !== 'typegpu-visible-canvas') {
    invalidReasons.push('Bilig proof method is not typegpu-visible-canvas')
  }
  for (const entry of proof.evidence) {
    if (entry.startsWith('gap=')) {
      invalidReasons.push(entry.slice('gap='.length))
    }
  }
  const gridProjectedRevision = evidence.get('gridProjectedRevision') ?? ''
  const gridAuthoritativeRevision = evidence.get('gridAuthoritativeRevision') ?? ''
  const gridLocalRevision = evidence.get('gridLocalRevision') ?? ''
  const tileSceneRevision = evidence.get('tileSceneRevision') ?? ''
  const frameProofSignature = evidence.get('frameProofSignature') ?? ''
  const presentedFrameProofSignature = evidence.get('presentedFrameProofSignature') ?? ''
  const currentSceneOwnershipSignature = evidence.get('currentSceneOwnershipSignature') ?? ''
  const presentedSceneOwnershipSignature = evidence.get('presentedSceneOwnershipSignature') ?? ''
  const currentContentSignature = evidence.get('currentContentSignature') ?? ''
  const presentedContentSignature = evidence.get('presentedContentSignature') ?? ''
  const currentTextSignature = evidence.get('currentTextSignature') ?? ''
  const presentedTextSignature = evidence.get('presentedTextSignature') ?? ''
  const currentRectSignature = evidence.get('currentRectSignature') ?? ''
  const presentedRectSignature = evidence.get('presentedRectSignature') ?? ''
  const tilePaneCount = numericEvidence(evidence, 'tilePaneCount')
  const headerPaneCount = numericEvidence(evidence, 'headerPaneCount')
  const presentedTilePaneCount = numericEvidence(evidence, 'presentedTilePaneCount')
  const presentedHeaderPaneCount = numericEvidence(evidence, 'presentedHeaderPaneCount')
  const currentTextRunCount = numericEvidence(evidence, 'currentTextRunCount')
  const presentedTextRunCount = numericEvidence(evidence, 'presentedTextRunCount')
  const currentRectCount = numericEvidence(evidence, 'currentRectCount')
  const presentedRectCount = numericEvidence(evidence, 'presentedRectCount')
  const expectedPixelWidth = numericEvidence(evidence, 'expectedPixelWidth')
  const expectedPixelHeight = numericEvidence(evidence, 'expectedPixelHeight')
  const canvasPixelWidth = numericEvidence(evidence, 'canvasPixelWidth')
  const canvasPixelHeight = numericEvidence(evidence, 'canvasPixelHeight')
  if (evidence.get('mode') !== 'typegpu-v3') {
    invalidReasons.push(`renderer mode is ${evidence.get('mode') ?? 'missing'}`)
  }
  if (evidence.get('contractVersion') !== sameCorpusUiRenderProofContractVersion) {
    invalidReasons.push(`missing contractVersion=${sameCorpusUiRenderProofContractVersion}`)
  }
  if (evidence.get('backendStatus') !== 'ready') {
    invalidReasons.push(`TypeGPU backend is ${evidence.get('backendStatus') ?? 'missing'}`)
  }
  if (evidence.get('frameProofStatus') !== 'presented') {
    invalidReasons.push(`frame proof is ${evidence.get('frameProofStatus') ?? 'missing'}`)
  }
  if (frameProofSignature.length === 0) {
    invalidReasons.push('frame proof signature is missing')
  }
  if (presentedFrameProofSignature.length === 0) {
    invalidReasons.push('presented frame proof signature is missing')
  }
  if (evidence.get('hasPresentedFrame') !== 'true') {
    invalidReasons.push('current frame signature has not been presented')
  }
  if (frameProofSignature.length > 0 && presentedFrameProofSignature.length > 0 && presentedFrameProofSignature !== frameProofSignature) {
    invalidReasons.push('presented frame proof signature does not match current frame')
  }
  if (currentSceneOwnershipSignature.length === 0) {
    invalidReasons.push('current visible-scene ownership signature is missing')
  }
  if (presentedSceneOwnershipSignature.length === 0) {
    invalidReasons.push('presented visible-scene ownership signature is missing')
  }
  if (
    currentSceneOwnershipSignature.length > 0 &&
    presentedSceneOwnershipSignature.length > 0 &&
    presentedSceneOwnershipSignature !== currentSceneOwnershipSignature
  ) {
    invalidReasons.push('presented visible-scene ownership does not match current scene')
  }
  if (evidence.get('hasPresentedVisibleFrame') !== 'true') {
    invalidReasons.push('visible frame has not been presented')
  }
  if (!isPositiveNumber(tilePaneCount) || !isPositiveNumber(headerPaneCount)) {
    invalidReasons.push('current tile/header pane counts are empty')
  }
  if (!isPositiveNumber(presentedTilePaneCount) || !isPositiveNumber(presentedHeaderPaneCount)) {
    invalidReasons.push('presented tile/header pane counts are empty')
  }
  if (
    isPositiveNumber(tilePaneCount) &&
    isPositiveNumber(headerPaneCount) &&
    isPositiveNumber(presentedTilePaneCount) &&
    isPositiveNumber(presentedHeaderPaneCount) &&
    (presentedTilePaneCount !== tilePaneCount || presentedHeaderPaneCount !== headerPaneCount)
  ) {
    invalidReasons.push('presented tile/header pane counts do not cover the current visible panes')
  }
  if (evidence.get('canvasCoversViewport') !== 'true') {
    invalidReasons.push('TypeGPU canvas backing pixels do not cover the viewport')
  }
  if (currentContentSignature.length === 0) {
    invalidReasons.push('current visible content signature is missing')
  }
  if (presentedContentSignature.length === 0) {
    invalidReasons.push('presented visible content signature is missing')
  }
  if (currentContentSignature.length > 0 && presentedContentSignature.length > 0 && presentedContentSignature !== currentContentSignature) {
    invalidReasons.push('presented visible content signature does not match current tiles')
  }
  if (currentTextSignature.length === 0) {
    invalidReasons.push('current visible text signature is missing')
  }
  if (presentedTextSignature.length === 0) {
    invalidReasons.push('presented visible text signature is missing')
  }
  if (currentTextSignature.length > 0 && presentedTextSignature.length > 0 && presentedTextSignature !== currentTextSignature) {
    invalidReasons.push('presented visible text signature does not match current tiles')
  }
  if (currentRectSignature.length === 0) {
    invalidReasons.push('current visible rect signature is missing')
  }
  if (presentedRectSignature.length === 0) {
    invalidReasons.push('presented visible rect signature is missing')
  }
  if (currentRectSignature.length > 0 && presentedRectSignature.length > 0 && presentedRectSignature !== currentRectSignature) {
    invalidReasons.push('presented visible rect signature does not match current tiles')
  }
  if (!isNonNegativeNumber(currentTextRunCount) || !isNonNegativeNumber(presentedTextRunCount)) {
    invalidReasons.push('visible text run payload counts are missing')
  }
  if (
    isNonNegativeNumber(currentTextRunCount) &&
    isNonNegativeNumber(presentedTextRunCount) &&
    currentTextRunCount !== presentedTextRunCount
  ) {
    invalidReasons.push('presented visible text run count does not match current tiles')
  }
  if (!isPositiveNumber(currentRectCount) || !isPositiveNumber(presentedRectCount)) {
    invalidReasons.push('visible rect payload counts are empty')
  }
  if (isPositiveNumber(currentRectCount) && isPositiveNumber(presentedRectCount) && currentRectCount !== presentedRectCount) {
    invalidReasons.push('presented visible rect count does not match current tiles')
  }
  if (!isPositiveNumber(expectedPixelWidth) || !isPositiveNumber(expectedPixelHeight)) {
    invalidReasons.push('expected viewport pixel dimensions are missing')
  }
  if (!isPositiveNumber(canvasPixelWidth) || !isPositiveNumber(canvasPixelHeight)) {
    invalidReasons.push('TypeGPU canvas pixel dimensions are missing')
  }
  if (
    isPositiveNumber(expectedPixelWidth) &&
    isPositiveNumber(expectedPixelHeight) &&
    isPositiveNumber(canvasPixelWidth) &&
    isPositiveNumber(canvasPixelHeight) &&
    (canvasPixelWidth < expectedPixelWidth - 2 || canvasPixelHeight < expectedPixelHeight - 2)
  ) {
    invalidReasons.push('TypeGPU canvas backing pixels do not cover the viewport')
  }
  if (gridAuthoritativeRevision.length === 0) {
    invalidReasons.push('grid authoritative render revision is missing')
  }
  if (evidence.get('typeGpuAuthoritativeRevision') !== gridAuthoritativeRevision) {
    invalidReasons.push('TypeGPU authoritative render revision does not match the grid revision')
  }
  if (evidence.get('visibleAuthoritativeRevision') !== gridAuthoritativeRevision) {
    invalidReasons.push('visible authoritative render revision does not match the grid revision')
  }
  if (gridLocalRevision.length === 0) {
    invalidReasons.push('grid local render revision is missing')
  }
  if (evidence.get('typeGpuLocalRevision') !== gridLocalRevision) {
    invalidReasons.push('TypeGPU local render revision does not match the grid revision')
  }
  if (evidence.get('visibleLocalRevision') !== gridLocalRevision) {
    invalidReasons.push('visible local render revision does not match the grid revision')
  }
  if (gridProjectedRevision.length === 0) {
    invalidReasons.push('grid projected render revision is missing')
  }
  if (evidence.get('typeGpuProjectedRevision') !== gridProjectedRevision) {
    invalidReasons.push('TypeGPU projected render revision does not match the grid revision')
  }
  if (evidence.get('visibleProjectedRevision') !== gridProjectedRevision) {
    invalidReasons.push('visible projected render revision does not match the grid revision')
  }
  if (tileSceneRevision.length === 0) {
    invalidReasons.push('tile scene revision is missing')
  }
  if (evidence.get('visibleRenderRevision') !== tileSceneRevision) {
    invalidReasons.push('visible render revision does not match the tile scene revision')
  }
  return invalidReasons
}

function isLegacyInsufficientPixelProof(proof: SameCorpusProductPixelGridProof, evidence: ReadonlyMap<string, string>): boolean {
  if (proof.product === 'bilig') {
    return (
      evidence.get('mode') === 'typegpu-v3' ||
      evidence.has('tilePaneCount') ||
      evidence.has('headerPaneCount') ||
      proof.evidence.some((entry) => entry.startsWith('expectedPixel'))
    )
  }
  return proof.evidence.length > 0 && !proof.evidence.some((entry) => entry === 'visible workbook grid surface was not found')
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
    JSON.stringify(proof.pixelGridProof.missingProducts) !== JSON.stringify(expected.pixelGridProof.missingProducts) ||
    JSON.stringify(proof.pixelGridProof.productVerdicts) !== JSON.stringify(expected.pixelGridProof.productVerdicts)
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
  const productVerdicts = args.visualProofs.map((entry) => validateSameCorpusProductPixelGridProof(entry.pixelGridProof))
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
      productVerdicts,
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

function isNonNegativeNumber(value: number | null): value is number {
  return value !== null && value >= 0
}

function minimumEvidenceNumber(evidence: ReadonlyMap<string, string>, key: string, minimum: number): boolean {
  const value = numericEvidence(evidence, key)
  return value !== null && value >= minimum
}

function maximumEvidenceNumber(evidence: ReadonlyMap<string, string>, key: string, maximum: number): boolean {
  const value = numericEvidence(evidence, key)
  return value !== null && value <= maximum
}

function hasStrictScreenshotPixelGridEvidence(evidence: ReadonlyMap<string, string>): boolean {
  return (
    evidence.get('pixelGridProofVersion') === 'grid-pixels-v1' &&
    evidence.get('pixelSampleSource') === 'screenshot' &&
    minimumEvidenceNumber(evidence, 'visibleGridLinePixels', 24) &&
    minimumEvidenceNumber(evidence, 'verticalLineRuns', 3) &&
    minimumEvidenceNumber(evidence, 'horizontalLineRuns', 3) &&
    minimumEvidenceNumber(evidence, 'verticalLineCoverageBands', 5) &&
    minimumEvidenceNumber(evidence, 'horizontalLineCoverageBands', 5) &&
    maximumEvidenceNumber(evidence, 'largestVerticalLineGapPx', 384) &&
    maximumEvidenceNumber(evidence, 'largestHorizontalLineGapPx', 160)
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
