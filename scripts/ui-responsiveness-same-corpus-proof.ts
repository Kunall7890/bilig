import { mkdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'

import type { Page } from '@playwright/test'

import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import { captureProductScreenshot, readProductPixelGridProof } from './ui-responsiveness-same-corpus-pixel-proof-page.ts'
import {
  missingSemanticUiProof,
  readProductSemanticUiProof,
  validateSameCorpusProductSemanticUiProof,
  type SameCorpusProductSemanticUiProof,
  type SameCorpusSemanticUiProof,
  type SameCorpusMutationTargetProof,
  type SameCorpusMutationTargetReadback,
  type SameCorpusMutationTargetScreenshotProof,
  type SameCorpusMutationTargetScreenshotProofSet,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type {
  SameCorpusCaptureMeasurement,
  SameCorpusCaptureCorpusVerification,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'
import { uiSameCorpusWorkloadMutatesWorkbook, type UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
export const sameCorpusUiRenderProofContractVersion = 'same-corpus-ui-v6'

export {
  validateSameCorpusProductSemanticUiProof,
  type SameCorpusProductSemanticUiProof,
  type SameCorpusProductSemanticUiProofVerdict,
  type SameCorpusSemanticUiProof,
  type SameCorpusMutationTargetProof,
  type SameCorpusMutationTargetReadback,
  type SameCorpusMutationTargetScreenshotProof,
  type SameCorpusMutationTargetScreenshotProofSet,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'

export interface SameCorpusScenarioProof {
  readonly biligMeanMs: number
  readonly biligP95Ms: number
  readonly googleMeanMs: number
  readonly googleP95Ms: number
  readonly microsoftExcelWebMeanMs?: number
  readonly microsoftExcelWebP95Ms?: number
  /** Google Sheets time divided by Bilig time. Values >= 10 mean Bilig is at least 10x faster. */
  readonly meanRatio: number
  /** Google Sheets p95 divided by Bilig p95. Values >= 10 mean Bilig is at least 10x faster at p95. */
  readonly p95Ratio: number
  /** Microsoft Excel Web time divided by Bilig time. Values >= 10 mean Bilig is at least 10x faster. */
  readonly microsoftExcelWebMeanRatio?: number
  /** Microsoft Excel Web p95 divided by Bilig p95. Values >= 10 mean Bilig is at least 10x faster at p95. */
  readonly microsoftExcelWebP95Ratio?: number
  readonly screenshotProof: SameCorpusScreenshotProof
  readonly pixelGridProof: SameCorpusPixelGridProof
  readonly semanticUiProof: SameCorpusSemanticUiProof
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
  readonly semanticUiProof?: SameCorpusProductSemanticUiProof | undefined
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

const sameCorpusScreenshotProducts = [
  'bilig',
  'google-sheets',
  'microsoft-excel-web',
] as const satisfies readonly UiResponsivenessSameCorpusProduct[]

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
  const currentSceneEpochSignature = evidence.get('currentSceneEpochSignature') ?? ''
  const presentedSceneEpochSignature = evidence.get('presentedSceneEpochSignature') ?? ''
  const currentSceneOwnershipSignature = evidence.get('currentSceneOwnershipSignature') ?? ''
  const presentedSceneOwnershipSignature = evidence.get('presentedSceneOwnershipSignature') ?? ''
  const currentWorkbookRevision = evidence.get('currentWorkbookRevision') ?? ''
  const presentedWorkbookRevision = evidence.get('presentedWorkbookRevision') ?? ''
  const currentSemanticMutationRevision = evidence.get('currentSemanticMutationRevision') ?? ''
  const presentedSemanticMutationRevision = evidence.get('presentedSemanticMutationRevision') ?? ''
  const currentViewportRevision = evidence.get('currentViewportRevision') ?? ''
  const presentedViewportRevision = evidence.get('presentedViewportRevision') ?? ''
  const currentSelectionRevision = evidence.get('currentSelectionRevision') ?? ''
  const presentedSelectionRevision = evidence.get('presentedSelectionRevision') ?? ''
  const currentFillHandleRevision = evidence.get('currentFillHandleRevision') ?? ''
  const presentedFillHandleRevision = evidence.get('presentedFillHandleRevision') ?? ''
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
  if (currentSceneEpochSignature.length === 0) {
    invalidReasons.push('current visible-scene epoch signature is missing')
  }
  if (presentedSceneOwnershipSignature.length === 0) {
    invalidReasons.push('presented visible-scene ownership signature is missing')
  }
  if (presentedSceneEpochSignature.length === 0) {
    invalidReasons.push('presented visible-scene epoch signature is missing')
  }
  if (
    currentSceneEpochSignature.length > 0 &&
    presentedSceneEpochSignature.length > 0 &&
    presentedSceneEpochSignature !== currentSceneEpochSignature
  ) {
    invalidReasons.push('presented visible-scene epoch does not match current authoritative scene')
  }
  if (
    currentSceneOwnershipSignature.length > 0 &&
    presentedSceneOwnershipSignature.length > 0 &&
    presentedSceneOwnershipSignature !== currentSceneOwnershipSignature
  ) {
    invalidReasons.push('presented visible-scene ownership does not match current scene')
  }
  comparePresentedOwnershipRevision(invalidReasons, currentWorkbookRevision, presentedWorkbookRevision, 'workbook revision')
  comparePresentedOwnershipRevision(
    invalidReasons,
    currentSemanticMutationRevision,
    presentedSemanticMutationRevision,
    'semantic mutation revision',
  )
  comparePresentedOwnershipRevision(invalidReasons, currentViewportRevision, presentedViewportRevision, 'viewport revision')
  comparePresentedOwnershipRevision(invalidReasons, currentSelectionRevision, presentedSelectionRevision, 'selection revision')
  comparePresentedOwnershipRevision(invalidReasons, currentFillHandleRevision, presentedFillHandleRevision, 'fill-handle revision')
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
  readonly workload?: UiResponsivenessSameCorpusWorkload
}): SameCorpusScenarioProof {
  return buildScenarioProof({
    biligTiming: summarizeNumbers(primaryCaptureTimingSamples(args.bilig, args.workload)),
    googleSheetsTiming: summarizeNumbers(primaryCaptureTimingSamples(args.googleSheets, args.workload)),
    microsoftExcelWebTiming: args.microsoftExcelWeb
      ? summarizeNumbers(primaryCaptureTimingSamples(args.microsoftExcelWeb, args.workload))
      : null,
    visualProofs: args.visualProofs,
    workload: args.workload ?? 'open-workbook',
    sampleCount: Math.min(
      args.bilig.operationResponseMsSamples.length,
      args.googleSheets.operationResponseMsSamples.length,
      ...(args.microsoftExcelWeb ? [args.microsoftExcelWeb.operationResponseMsSamples.length] : []),
    ),
  })
}

export function buildScorecardScenarioProof(args: {
  readonly bilig: UiResponsivenessSameCorpusMeasurement
  readonly googleSheets: UiResponsivenessSameCorpusMeasurement
  readonly microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement | undefined
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
  readonly workload?: UiResponsivenessSameCorpusWorkload
  readonly sampleCount?: number
}): SameCorpusScenarioProof {
  return buildScenarioProof({
    biligTiming: primaryScorecardTiming(args.bilig, args.workload),
    googleSheetsTiming: primaryScorecardTiming(args.googleSheets, args.workload),
    microsoftExcelWebTiming: args.microsoftExcelWeb ? primaryScorecardTiming(args.microsoftExcelWeb, args.workload) : null,
    visualProofs: args.visualProofs,
    workload: args.workload ?? 'open-workbook',
    sampleCount:
      args.sampleCount ??
      Math.min(
        args.bilig.operationResponseMs.samples.length,
        args.googleSheets.operationResponseMs.samples.length,
        ...(args.microsoftExcelWeb ? [args.microsoftExcelWeb.operationResponseMs.samples.length] : []),
      ),
  })
}

export async function captureSameCorpusProductVisualProof(args: {
  readonly caseId: string
  readonly corpusVerification: SameCorpusCaptureCorpusVerification
  readonly outputPath: string
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly sampleCount: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly mutationTargetProofs?: readonly SameCorpusMutationTargetProof[]
}): Promise<SameCorpusProductVisualProof> {
  const screenshotPath = screenshotArtifactPath(args.outputPath, args.caseId, args.product, args.sampleIndex)
  mkdirSync(dirname(screenshotPath), { recursive: true })
  const screenshot = await captureProductScreenshot(args.page, args.product, screenshotPath)
  const pixelGridProof = await readProductPixelGridProof(args.page, args.product, screenshot.buffer, sameCorpusUiRenderProofContractVersion)
  const semanticUiProof = await readProductSemanticUiProof({
    corpusVerification: args.corpusVerification,
    page: args.page,
    product: args.product,
    screenshot,
    pixelGridProof,
    workload: args.workload,
    sampleCount: args.sampleCount,
    mutationTargetProofs: args.mutationTargetProofs,
  })
  return {
    product: args.product,
    screenshotPath: screenshot.captured ? repoRelativePath(screenshotPath) : null,
    screenshotCaptured: screenshot.captured,
    pixelGridProof,
    semanticUiProof,
  }
}

export function validateSameCorpusScenarioProof(
  proof: SameCorpusScenarioProof,
  caseId: string,
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  workload: UiResponsivenessSameCorpusWorkload = 'open-workbook',
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): void {
  validateSameCorpusScenarioScreenshotArtifacts(proof.screenshotProof, caseId)
  const expected = buildScorecardScenarioProof({
    bilig,
    googleSheets,
    workload,
    sampleCount: Math.min(
      bilig.operationResponseMs.samples.length,
      googleSheets.operationResponseMs.samples.length,
      ...(microsoftExcelWeb ? [microsoftExcelWeb.operationResponseMs.samples.length] : []),
    ),
    microsoftExcelWeb,
    visualProofs: proof.pixelGridProof.products.map((entry) => ({
      product: entry.product,
      screenshotPath:
        proof.screenshotProof.artifactPaths.find((artifact) => sameCorpusScreenshotArtifactProduct(artifact, caseId) === entry.product) ??
        null,
      screenshotCaptured: !proof.screenshotProof.missingProducts.includes(entry.product),
      pixelGridProof: entry,
      semanticUiProof: proof.semanticUiProof.products.find((semanticProof) => semanticProof.product === entry.product),
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
  if (
    proof.semanticUiProof.captured !== expected.semanticUiProof.captured ||
    JSON.stringify(proof.semanticUiProof.requiredProducts) !== JSON.stringify(expected.semanticUiProof.requiredProducts) ||
    JSON.stringify(proof.semanticUiProof.products) !== JSON.stringify(expected.semanticUiProof.products) ||
    JSON.stringify(proof.semanticUiProof.missingProducts) !== JSON.stringify(expected.semanticUiProof.missingProducts) ||
    JSON.stringify(proof.semanticUiProof.productVerdicts) !== JSON.stringify(expected.semanticUiProof.productVerdicts)
  ) {
    throw new Error(`UI responsiveness same-corpus semantic UI proof is stale: ${caseId}`)
  }
  if (!proof.screenshotProof.captured) {
    throw new Error(`UI responsiveness same-corpus scenario proof is missing screenshot proof: ${caseId}`)
  }
}

function validateSameCorpusScenarioScreenshotArtifacts(proof: SameCorpusScreenshotProof, caseId: string): void {
  const seenPaths = new Set<string>()
  const artifactProducts = new Map<UiResponsivenessSameCorpusProduct, string[]>()
  for (const artifactPath of proof.artifactPaths) {
    const normalizedPath = normalizeArtifactPath(artifactPath)
    if (seenPaths.has(normalizedPath)) {
      throw new Error(`UI responsiveness same-corpus screenshot proof has duplicate artifact paths: ${caseId}`)
    }
    seenPaths.add(normalizedPath)
    const product = sameCorpusScreenshotArtifactProduct(normalizedPath, caseId)
    if (!product) {
      throw new Error(`UI responsiveness same-corpus screenshot artifact path is not tied to scenario: ${caseId}`)
    }
    artifactProducts.set(product, [...(artifactProducts.get(product) ?? []), normalizedPath])
  }

  if (proof.captured && proof.missingProducts.length > 0) {
    throw new Error(`UI responsiveness same-corpus screenshot proof is internally inconsistent: ${caseId}`)
  }
  for (const product of proof.requiredProducts) {
    const paths = artifactProducts.get(product) ?? []
    if (proof.missingProducts.includes(product)) {
      if (paths.length > 0) {
        throw new Error(`UI responsiveness same-corpus screenshot proof has an artifact for missing product ${product}: ${caseId}`)
      }
      continue
    }
    if (paths.length !== 1) {
      throw new Error(`UI responsiveness same-corpus screenshot proof must include exactly one ${product} artifact: ${caseId}`)
    }
  }
}

function sameCorpusScreenshotArtifactProduct(artifactPath: string, caseId: string): UiResponsivenessSameCorpusProduct | null {
  const segments = normalizeArtifactPath(artifactPath).split('/').filter(Boolean)
  if (segments.length < 2 || segments.at(-2) !== caseId) {
    return null
  }
  const filename = segments.at(-1) ?? ''
  return sameCorpusScreenshotProducts.find((product) => filename === `${product}-sample-1.png`) ?? null
}

function normalizeArtifactPath(artifactPath: string): string {
  return artifactPath.replaceAll('\\', '/')
}

function buildScenarioProof(args: {
  readonly biligTiming: NumericSummary
  readonly googleSheetsTiming: NumericSummary
  readonly microsoftExcelWebTiming: NumericSummary | null
  readonly visualProofs: readonly SameCorpusProductVisualProof[]
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleCount: number
}): SameCorpusScenarioProof {
  const requiredProducts = ['bilig', 'google-sheets'] as const satisfies readonly UiResponsivenessSameCorpusProduct[]
  const screenshotProducts = new Set(
    args.visualProofs.filter((entry) => entry.screenshotCaptured && entry.screenshotPath).map((entry) => entry.product),
  )
  const pixelProducts = new Set(
    args.visualProofs.filter((entry) => isSameCorpusProductPixelGridProofComplete(entry.pixelGridProof)).map((entry) => entry.product),
  )
  const productVerdicts = args.visualProofs.map((entry) => validateSameCorpusProductPixelGridProof(entry.pixelGridProof))
  const semanticProducts = args.visualProofs.map((entry) => entry.semanticUiProof ?? missingSemanticUiProof(entry.product))
  const semanticProductSet = new Set(
    semanticProducts
      .filter(
        (entry) =>
          validateSameCorpusProductSemanticUiProof(entry, { workload: args.workload, sampleCount: args.sampleCount })
            .acceptedForCurrentScorecard,
      )
      .map((entry) => entry.product),
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
          microsoftExcelWebMeanRatio: ratio(args.microsoftExcelWebTiming.mean, args.biligTiming.mean),
          microsoftExcelWebP95Ratio: ratio(args.microsoftExcelWebTiming.p95, args.biligTiming.p95),
        }
      : {}),
    meanRatio: ratio(args.googleSheetsTiming.mean, args.biligTiming.mean),
    p95Ratio: ratio(args.googleSheetsTiming.p95, args.biligTiming.p95),
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
    semanticUiProof: {
      captured: requiredProducts.every((product) => semanticProductSet.has(product)),
      requiredProducts,
      products: semanticProducts,
      productVerdicts: semanticProducts.map((entry) =>
        validateSameCorpusProductSemanticUiProof(entry, { workload: args.workload, sampleCount: args.sampleCount }),
      ),
      missingProducts: requiredProducts.filter((product) => !semanticProductSet.has(product)),
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

function comparePresentedOwnershipRevision(invalidReasons: string[], current: string, presented: string, label: string): void {
  if (current.length === 0) {
    invalidReasons.push(`current visible-scene ${label} is missing`)
    return
  }
  if (presented.length === 0) {
    invalidReasons.push(`presented visible-scene ${label} is missing`)
    return
  }
  if (presented !== current) {
    invalidReasons.push(`presented visible-scene ${label} does not match current scene`)
  }
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

function primaryCaptureTimingSamples(
  measurement: SameCorpusCaptureMeasurement,
  workload: UiResponsivenessSameCorpusWorkload | undefined,
): readonly number[] {
  if (workload && uiSameCorpusWorkloadMutatesWorkbook(workload) && measurement.committedTargetProofMsSamples) {
    return measurement.committedTargetProofMsSamples
  }
  return measurement.scrollEventResponseMsSamples ?? measurement.operationResponseMsSamples
}

function primaryScorecardTiming(
  measurement: UiResponsivenessSameCorpusMeasurement,
  workload: UiResponsivenessSameCorpusWorkload | undefined,
): NumericSummary {
  if (workload && uiSameCorpusWorkloadMutatesWorkbook(workload) && measurement.committedTargetProofMs) {
    return measurement.committedTargetProofMs
  }
  return measurement.scrollEventResponseMs ?? measurement.operationResponseMs
}

function ratio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
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
