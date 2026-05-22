import { createHash } from 'node:crypto'

import { summarizeNumbers } from '../packages/benchmarks/src/stats.js'
import {
  buildWorkbookBenchmarkCorpus,
  isWorkbookBenchmarkCorpusId,
  type WorkbookBenchmarkCorpusId,
} from '../packages/benchmarks/src/workbook-corpus.js'
import {
  buildSameCorpusFingerprint,
  sameCorpusFingerprintVersion,
  type SameCorpusCaptureCorpusFingerprint,
} from './ui-responsiveness-same-corpus-fingerprint.ts'
import type {
  SameCorpusCapture,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  SameCorpusCaptureRunManifest,
  SameCorpusProductSourceWorkbookFingerprint,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
  UiResponsivenessSameCorpusProof,
  UiResponsivenessSameCorpusRunManifest,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { sameCorpusUiCaptureToolVersion } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type { SameCorpusProductVisualProof, SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import {
  buildScorecardScenarioProof,
  sameCorpusUiRenderProofContractVersion,
  validateSameCorpusScenarioProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'
import { cloneSameCorpusVerification, isSha256Hex, validateSummary } from './ui-responsiveness-same-corpus-validation-helpers.ts'

export { sameCorpusUiCaptureToolVersion } from './ui-responsiveness-same-corpus-scorecard-types.ts'
export type {
  SameCorpusCapture,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  SameCorpusCaptureRunManifest,
  SameCorpusCaptureVerifiedCell,
  SameCorpusProductSourceWorkbookFingerprint,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
  UiResponsivenessSameCorpusProof,
  UiResponsivenessSameCorpusRunManifest,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'

const requiredSameCorpusWorkloads = requiredUiResponsivenessSameCorpusWorkloads
const sameCorpusSampleCount = 3
const strictRenderedGridProofLimitation =
  'Some same-corpus cases retain timing evidence but do not satisfy strict rendered-grid proof, so they cannot count toward Google Sheets 10x UI claims.'
const expectedCorpusFingerprintCache = new Map<WorkbookBenchmarkCorpusId, SameCorpusCaptureCorpusFingerprint>()

export function buildMissingSameCorpusProof(): UiResponsivenessSameCorpusProof {
  return {
    captured: false,
    evidenceKind: 'not-captured',
    requiredProductCount: 2,
    requiredCaseCount: requiredSameCorpusWorkloads.length,
    tenXMeanAndP95CaseCount: 0,
    coveredCorpusCaseIds: [],
    runManifest: buildSameCorpusRunManifest([]),
    limitations: ['Same-corpus live browser timing against Bilig and Google Sheets has not been captured yet.'],
    cases: [],
  }
}

export function buildSameCorpusProof(capture: SameCorpusCapture): UiResponsivenessSameCorpusProof {
  validateSameCorpusCapture(capture)
  const cases = capture.cases.map(buildSameCorpusCase)
  const proof: UiResponsivenessSameCorpusProof = {
    captured: true,
    evidenceKind: 'same-corpus-browser-capture',
    requiredProductCount: 2,
    requiredCaseCount: requiredSameCorpusWorkloads.length,
    tenXMeanAndP95CaseCount: cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length,
    coveredCorpusCaseIds: [...new Set(cases.map((entry) => entry.corpusCaseId))].toSorted(),
    runManifest: buildSameCorpusRunManifest(cases),
    limitations: sameCorpusProofLimitations(capture.limitations, cases),
    cases,
  }
  validateSameCorpusProof(proof)
  return proof
}

export function validateSameCorpusProof(proof: UiResponsivenessSameCorpusProof): void {
  if (proof.requiredProductCount !== 2) {
    throw new Error('UI responsiveness same-corpus Google Sheets proof must compare Bilig and Google Sheets')
  }
  if (proof.requiredCaseCount !== requiredSameCorpusWorkloads.length) {
    throw new Error('UI responsiveness same-corpus proof required case count is stale')
  }
  if (!proof.captured) {
    if (proof.evidenceKind !== 'not-captured' || proof.cases.length !== 0) {
      throw new Error('UI responsiveness same-corpus proof has stale not-captured metadata')
    }
    if (proof.limitations.length === 0) {
      throw new Error('UI responsiveness same-corpus proof must disclose that capture is missing')
    }
    validateSameCorpusRunManifest(proof)
    return
  }
  if (proof.evidenceKind !== 'same-corpus-browser-capture') {
    throw new Error('UI responsiveness same-corpus proof has stale capture metadata')
  }
  for (const workload of requiredSameCorpusWorkloads) {
    if (!proof.cases.some((entry) => entry.workload === workload)) {
      throw new Error(`UI responsiveness same-corpus proof is missing required workload: ${workload}`)
    }
  }
  if (proof.cases.length !== proof.requiredCaseCount) {
    throw new Error('UI responsiveness same-corpus proof must include every required captured case')
  }
  const tenXCaseCount = proof.cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length
  if (proof.tenXMeanAndP95CaseCount !== tenXCaseCount) {
    throw new Error('UI responsiveness same-corpus proof 10x case count is stale')
  }
  const coveredCorpusCaseIds = [...new Set(proof.cases.map((entry) => entry.corpusCaseId))].toSorted()
  if (JSON.stringify(proof.coveredCorpusCaseIds) !== JSON.stringify(coveredCorpusCaseIds)) {
    throw new Error('UI responsiveness same-corpus proof covered corpus IDs are stale')
  }
  const hasMissingStrictGridProof = proof.cases.some((entry) => !entry.scenarioProof.pixelGridProof.captured)
  if (hasMissingStrictGridProof && !proof.limitations.includes(strictRenderedGridProofLimitation)) {
    throw new Error('UI responsiveness same-corpus proof must disclose missing strict rendered-grid proof')
  }
  for (const entry of proof.cases) {
    validateSameCorpusCase(entry)
  }
  validateSameCorpusRunManifest(proof)
}

function buildSameCorpusRunManifest(cases: readonly UiResponsivenessSameCorpusCase[]): UiResponsivenessSameCorpusRunManifest {
  const capturedWorkloads = cases.map((entry) => entry.workload)
  const corpusCaseIds = [...new Set(cases.map((entry) => entry.corpusCaseId))].toSorted()
  const corpusFingerprints = uniqueCorpusFingerprints(cases)
  const productSourceWorkbookFingerprints = uniqueProductSourceWorkbookFingerprints(cases)
  const materializedCellCounts = [...new Set(cases.map((entry) => entry.materializedCells))].toSorted((left, right) => left - right)
  const strictRenderedGridProofCaseCount = cases.filter((entry) => entry.scenarioProof.pixelGridProof.captured).length
  const legacyInsufficientRenderedGridProofCaseCount = cases.filter((entry) =>
    entry.scenarioProof.pixelGridProof.productVerdicts.some((verdict) => verdict.evidenceStatus === 'legacy-insufficient'),
  ).length
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length
  const invalidReasons = sameCorpusRunManifestInvalidReasons({
    capturedWorkloads,
    caseCount: cases.length,
    corpusCaseIds,
    corpusFingerprints,
    productSourceWorkbookFingerprints,
    legacyInsufficientRenderedGridProofCaseCount,
    materializedCellCounts,
    strictRenderedGridProofCaseCount,
    tenXMeanAndP95CaseCount,
  })
  return {
    artifactGenerator: 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts',
    contractVersion: sameCorpusUiRenderProofContractVersion,
    requiredProducts: ['bilig', 'google-sheets'],
    requiredWorkloads: requiredSameCorpusWorkloads,
    capturedWorkloads,
    corpusCaseIds,
    corpusFingerprints,
    productSourceWorkbookFingerprints,
    materializedCellCounts,
    sampleCount: manifestSampleCount(cases),
    caseCount: cases.length,
    strictRenderedGridProofCaseCount,
    legacyInsufficientRenderedGridProofCaseCount,
    tenXMeanAndP95CaseCount,
    currentContractEvidenceComplete: !invalidReasons.some(
      (reason) => reason !== 'not every required workload is 10x against Google Sheets',
    ),
    googleSheetsTenXRequirementSatisfied: invalidReasons.length === 0,
    invalidReasons,
  }
}

export function buildSameCorpusCaptureRunManifest(
  cases: readonly SameCorpusCaptureCase[],
  sampleCount: number,
): SameCorpusCaptureRunManifest {
  const capturedWorkloads = cases.map((entry) => entry.workload)
  const corpusCaseIds = [...new Set(cases.map((entry) => entry.corpusCaseId))].toSorted()
  const corpusFingerprints = uniqueCaptureCorpusFingerprints(cases)
  const productSourceWorkbookFingerprints = uniqueCaptureProductSourceWorkbookFingerprints(cases)
  const materializedCellCounts = [...new Set(cases.map((entry) => entry.materializedCells))].toSorted((left, right) => left - right)
  const strictRenderedGridProofCaseCount = cases.filter((entry) => entry.scenarioProof.pixelGridProof.captured).length
  const legacyInsufficientRenderedGridProofCaseCount = cases.filter((entry) =>
    entry.scenarioProof.pixelGridProof.productVerdicts.some((verdict) => verdict.evidenceStatus === 'legacy-insufficient'),
  ).length
  const invalidReasons = sameCorpusCaptureRunManifestInvalidReasons({
    capturedWorkloads,
    caseCount: cases.length,
    corpusCaseIds,
    corpusFingerprints,
    productSourceWorkbookFingerprints,
    legacyInsufficientRenderedGridProofCaseCount,
    materializedCellCounts,
    strictRenderedGridProofCaseCount,
  })
  return {
    artifactGenerator: 'scripts/capture-ui-responsiveness-same-corpus.ts',
    captureToolVersion: sameCorpusUiCaptureToolVersion,
    contractVersion: sameCorpusUiRenderProofContractVersion,
    requiredProducts: ['bilig', 'google-sheets'],
    requiredWorkloads: requiredSameCorpusWorkloads,
    capturedWorkloads,
    corpusCaseIds,
    corpusFingerprints,
    productSourceWorkbookFingerprints,
    materializedCellCounts,
    sampleCount,
    caseCount: cases.length,
    strictRenderedGridProofCaseCount,
    legacyInsufficientRenderedGridProofCaseCount,
    currentContractEvidenceComplete: invalidReasons.length === 0,
    captureRunSignature: sameCorpusCaptureRunSignature(cases),
    invalidReasons,
  }
}

function sameCorpusRunManifestInvalidReasons(args: {
  readonly capturedWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly caseCount: number
  readonly corpusCaseIds: readonly string[]
  readonly corpusFingerprints: readonly SameCorpusCaptureCorpusFingerprint[]
  readonly productSourceWorkbookFingerprints: readonly SameCorpusProductSourceWorkbookFingerprint[]
  readonly legacyInsufficientRenderedGridProofCaseCount: number
  readonly materializedCellCounts: readonly number[]
  readonly strictRenderedGridProofCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
}): string[] {
  const invalidReasons: string[] = []
  if (args.caseCount !== requiredSameCorpusWorkloads.length) {
    invalidReasons.push('required workload count is incomplete')
  }
  const missingWorkloads = requiredSameCorpusWorkloads.filter((workload) => !args.capturedWorkloads.includes(workload))
  if (missingWorkloads.length > 0) {
    invalidReasons.push(`missing required workloads: ${missingWorkloads.join(', ')}`)
  }
  if (new Set(args.capturedWorkloads).size !== args.capturedWorkloads.length) {
    invalidReasons.push('duplicate workload evidence is present')
  }
  if (args.corpusCaseIds.length !== 1) {
    invalidReasons.push('same-corpus evidence must use exactly one corpus case')
  }
  if (args.corpusFingerprints.length !== 1) {
    invalidReasons.push('same-corpus evidence must use exactly one benchmark workbook fingerprint')
  }
  if (!requiredProductSourceWorkbookFingerprintsPresent(args.productSourceWorkbookFingerprints)) {
    invalidReasons.push('source workbook fingerprint must be stable for every required product')
  }
  if (args.materializedCellCounts.length > 1) {
    invalidReasons.push('same-corpus evidence has mixed materialized cell counts')
  }
  if (args.strictRenderedGridProofCaseCount !== requiredSameCorpusWorkloads.length) {
    invalidReasons.push(
      `strict rendered-grid proof covers ${String(args.strictRenderedGridProofCaseCount)}/${String(requiredSameCorpusWorkloads.length)} cases`,
    )
  }
  if (args.legacyInsufficientRenderedGridProofCaseCount > 0) {
    invalidReasons.push(
      `legacy-insufficient rendered-grid proof covers ${String(args.legacyInsufficientRenderedGridProofCaseCount)}/${String(
        requiredSameCorpusWorkloads.length,
      )} cases`,
    )
  }
  if (args.tenXMeanAndP95CaseCount !== requiredSameCorpusWorkloads.length) {
    invalidReasons.push('not every required workload is 10x against Google Sheets')
  }
  return invalidReasons
}

function sameCorpusCaptureRunManifestInvalidReasons(args: {
  readonly capturedWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly caseCount: number
  readonly corpusCaseIds: readonly string[]
  readonly corpusFingerprints: readonly SameCorpusCaptureCorpusFingerprint[]
  readonly productSourceWorkbookFingerprints: readonly SameCorpusProductSourceWorkbookFingerprint[]
  readonly legacyInsufficientRenderedGridProofCaseCount: number
  readonly materializedCellCounts: readonly number[]
  readonly strictRenderedGridProofCaseCount: number
}): string[] {
  const invalidReasons: string[] = []
  if (args.caseCount !== requiredSameCorpusWorkloads.length) {
    invalidReasons.push('required workload count is incomplete')
  }
  const missingWorkloads = requiredSameCorpusWorkloads.filter((workload) => !args.capturedWorkloads.includes(workload))
  if (missingWorkloads.length > 0) {
    invalidReasons.push(`missing required workloads: ${missingWorkloads.join(', ')}`)
  }
  if (new Set(args.capturedWorkloads).size !== args.capturedWorkloads.length) {
    invalidReasons.push('duplicate workload evidence is present')
  }
  if (args.corpusCaseIds.length !== 1) {
    invalidReasons.push('same-corpus evidence must use exactly one corpus case')
  }
  if (args.corpusFingerprints.length !== 1) {
    invalidReasons.push('same-corpus evidence must use exactly one benchmark workbook fingerprint')
  }
  if (!requiredProductSourceWorkbookFingerprintsPresent(args.productSourceWorkbookFingerprints)) {
    invalidReasons.push('source workbook fingerprint must be stable for every required product')
  }
  if (args.materializedCellCounts.length > 1) {
    invalidReasons.push('same-corpus evidence has mixed materialized cell counts')
  }
  if (args.strictRenderedGridProofCaseCount !== requiredSameCorpusWorkloads.length) {
    invalidReasons.push(
      `strict rendered-grid proof covers ${String(args.strictRenderedGridProofCaseCount)}/${String(requiredSameCorpusWorkloads.length)} cases`,
    )
  }
  if (args.legacyInsufficientRenderedGridProofCaseCount > 0) {
    invalidReasons.push(
      `legacy-insufficient rendered-grid proof covers ${String(args.legacyInsufficientRenderedGridProofCaseCount)}/${String(
        requiredSameCorpusWorkloads.length,
      )} cases`,
    )
  }
  return invalidReasons
}

function uniqueCorpusFingerprints(cases: readonly UiResponsivenessSameCorpusCase[]): readonly SameCorpusCaptureCorpusFingerprint[] {
  return uniqueByStableJson(
    cases.flatMap((entry) => sameCorpusCaseMeasurements(entry).map((measurement) => measurement.corpusVerification.corpusFingerprint)),
  )
}

function uniqueCaptureCorpusFingerprints(cases: readonly SameCorpusCaptureCase[]): readonly SameCorpusCaptureCorpusFingerprint[] {
  return uniqueByStableJson(
    cases.flatMap((entry) =>
      sameCorpusCaptureCaseMeasurements(entry).map((measurement) => measurement.corpusVerification.corpusFingerprint),
    ),
  )
}

function uniqueProductSourceWorkbookFingerprints(
  cases: readonly UiResponsivenessSameCorpusCase[],
): readonly SameCorpusProductSourceWorkbookFingerprint[] {
  return uniqueByStableJson(
    cases.flatMap((entry) =>
      sameCorpusCaseMeasurements(entry).map((measurement) => ({
        product: measurement.product,
        method: measurement.corpusVerification.method,
        source: measurement.source,
        sourceWorkbookSha256: measurement.corpusVerification.sourceWorkbookSha256,
      })),
    ),
  ).toSorted((left, right) => `${left.product}:${left.source}`.localeCompare(`${right.product}:${right.source}`))
}

function uniqueCaptureProductSourceWorkbookFingerprints(
  cases: readonly SameCorpusCaptureCase[],
): readonly SameCorpusProductSourceWorkbookFingerprint[] {
  return uniqueByStableJson(
    cases.flatMap((entry) =>
      sameCorpusCaptureCaseMeasurements(entry).map((measurement) => ({
        product: measurement.product,
        method: measurement.corpusVerification.method,
        source: measurement.source,
        sourceWorkbookSha256: measurement.corpusVerification.sourceWorkbookSha256,
      })),
    ),
  ).toSorted((left, right) => `${left.product}:${left.source}`.localeCompare(`${right.product}:${right.source}`))
}

function sameCorpusCaseMeasurements(
  entry: Pick<UiResponsivenessSameCorpusCase, 'bilig' | 'googleSheets' | 'microsoftExcelWeb'>,
): readonly UiResponsivenessSameCorpusMeasurement[] {
  return [entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])]
}

function sameCorpusCaptureCaseMeasurements(entry: SameCorpusCaptureCase): readonly SameCorpusCaptureMeasurement[] {
  return [entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])]
}

function uniqueByStableJson<T>(values: readonly T[]): T[] {
  const byKey = new Map<string, T>()
  for (const value of values) {
    byKey.set(JSON.stringify(value), value)
  }
  return [...byKey.values()]
}

function requiredProductSourceWorkbookFingerprintsPresent(fingerprints: readonly SameCorpusProductSourceWorkbookFingerprint[]): boolean {
  return (['bilig', 'google-sheets'] as const satisfies readonly UiResponsivenessSameCorpusProduct[]).every(
    (product) => fingerprints.filter((entry) => entry.product === product && entry.sourceWorkbookSha256 !== null).length === 1,
  )
}

function manifestSampleCount(cases: readonly UiResponsivenessSameCorpusCase[]): number {
  return cases.length === 0 ? 0 : Math.min(...cases.map((entry) => entry.sampleCount))
}

function validateSameCorpusRunManifest(proof: UiResponsivenessSameCorpusProof): void {
  if (!proof.runManifest) {
    throw new Error('UI responsiveness same-corpus proof is missing run manifest')
  }
  const expected = buildSameCorpusRunManifest(proof.cases)
  if (JSON.stringify(proof.runManifest) !== JSON.stringify(expected)) {
    throw new Error('UI responsiveness same-corpus run manifest is stale')
  }
}

function sameCorpusProofLimitations(captureLimitations: readonly string[], cases: readonly UiResponsivenessSameCorpusCase[]): string[] {
  const limitations = [...captureLimitations]
  if (cases.some((entry) => !entry.scenarioProof.pixelGridProof.captured) && !limitations.includes(strictRenderedGridProofLimitation)) {
    limitations.push(strictRenderedGridProofLimitation)
  }
  return limitations
}

function validateSameCorpusCapture(capture: SameCorpusCapture): void {
  if (capture.sampleCount < sameCorpusSampleCount) {
    throw new Error('UI responsiveness same-corpus capture must contain at least 3 samples per product')
  }
  if (capture.cases.length === 0) {
    throw new Error('UI responsiveness same-corpus capture must include at least one case')
  }
  validateSameCorpusCaptureRunManifest(capture)
  for (const entry of capture.cases) {
    const measurements = [entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])]
    const hasAnyScrollEventSamples = measurements.some(
      (measurement) => measurement.scrollEventResponseMsSamples !== undefined || measurement.scrollMovementPxSamples !== undefined,
    )
    const requiresScrollEventSamples = uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload) || hasAnyScrollEventSamples
    for (const measurement of measurements) {
      if (
        measurement.operationResponseMsSamples.length < capture.sampleCount ||
        measurement.postOperationFrameMsSamples.length < capture.sampleCount
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few samples for ${entry.id}`)
      }
      if (
        requiresScrollEventSamples &&
        ((measurement.scrollEventResponseMsSamples?.length ?? 0) < capture.sampleCount ||
          (measurement.scrollMovementPxSamples?.length ?? 0) < capture.sampleCount)
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few scroll-event samples for ${entry.id}`)
      }
      validateSameCorpusCaptureVerification(
        measurement.corpusVerification,
        measurement.product,
        entry.materializedCells,
        entry.corpusCaseId,
        entry.id,
      )
    }
    validateSameCorpusScenarioProof(
      entry.scenarioProof,
      entry.id,
      buildSameCorpusMeasurement(entry.bilig),
      buildSameCorpusMeasurement(entry.googleSheets),
      entry.microsoftExcelWeb ? buildSameCorpusMeasurement(entry.microsoftExcelWeb) : undefined,
    )
  }
}

function validateSameCorpusCaptureRunManifest(capture: SameCorpusCapture): void {
  const expected = buildSameCorpusCaptureRunManifest(capture.cases, capture.sampleCount)
  if (JSON.stringify(capture.runManifest) !== JSON.stringify(expected)) {
    throw new Error('UI responsiveness same-corpus capture run manifest is stale')
  }
}

function sameCorpusCaptureRunSignature(cases: readonly SameCorpusCaptureCase[]): string {
  return createHash('sha256').update(JSON.stringify(cases)).digest('hex')
}

function buildSameCorpusCase(captureCase: SameCorpusCaptureCase): UiResponsivenessSameCorpusCase {
  const bilig = buildSameCorpusMeasurement(captureCase.bilig)
  const googleSheets = buildSameCorpusMeasurement(captureCase.googleSheets)
  const microsoftExcelWeb = captureCase.microsoftExcelWeb ? buildSameCorpusMeasurement(captureCase.microsoftExcelWeb) : undefined
  const biligToGoogleSheetsMeanRatio = ratio(bilig.operationResponseMs.mean, googleSheets.operationResponseMs.mean)
  const biligToGoogleSheetsP95Ratio = ratio(bilig.operationResponseMs.p95, googleSheets.operationResponseMs.p95)
  const biligToMicrosoftExcelWebMeanRatio = microsoftExcelWeb
    ? ratio(bilig.operationResponseMs.mean, microsoftExcelWeb.operationResponseMs.mean)
    : undefined
  const biligToMicrosoftExcelWebP95Ratio = microsoftExcelWeb
    ? ratio(bilig.operationResponseMs.p95, microsoftExcelWeb.operationResponseMs.p95)
    : undefined
  const scrollEventMetrics = sameCorpusScrollEventMetrics(bilig, googleSheets, microsoftExcelWeb)
  const comparedProducts = [bilig, googleSheets, ...(microsoftExcelWeb ? [microsoftExcelWeb] : [])]
  const postOperationFrameGuardrailPassed = comparedProducts.every(
    (entry) => entry.postOperationFrameMs.p95 > 0 && entry.postOperationFrameMs.p95 <= 50,
  )
  const sourceWorkbookFingerprintGuardrailPassed = comparedProducts.every(hasCapturedSourceWorkbookFingerprint)
  const scrollMovementGuardrailPassed =
    scrollEventMetrics !== null && comparedProducts.every((entry) => (entry.scrollMovementPx?.min ?? 0) >= 1)
  const requiresScrollEventMetric = uiSameCorpusWorkloadRequiresScrollEventEvidence(captureCase.workload)
  const timingMetricPassedAgainstGoogleSheets = requiresScrollEventMetric
    ? scrollEventMetrics !== null &&
      scrollEventMetrics.biligToGoogleSheetsMeanRatio <= 0.1 &&
      scrollEventMetrics.biligToGoogleSheetsP95Ratio <= 0.1 &&
      scrollMovementGuardrailPassed
    : biligToGoogleSheetsMeanRatio <= 0.1 && biligToGoogleSheetsP95Ratio <= 0.1
  const timingMetricPassedAgainstMicrosoftExcelWeb = microsoftExcelWeb
    ? requiresScrollEventMetric
      ? scrollEventMetrics !== null &&
        scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio <= 0.1 &&
        scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio <= 0.1 &&
        scrollMovementGuardrailPassed
      : (biligToMicrosoftExcelWebMeanRatio ?? Number.POSITIVE_INFINITY) <= 0.1 &&
        (biligToMicrosoftExcelWebP95Ratio ?? Number.POSITIVE_INFINITY) <= 0.1
    : undefined
  const scenarioProof = buildScorecardScenarioProof({
    bilig,
    googleSheets,
    microsoftExcelWeb,
    visualProofs: scenarioProofVisualProofs(captureCase.scenarioProof),
  })
  const visualProofGuardrailPassed = scenarioProof.screenshotProof.captured && scenarioProof.pixelGridProof.captured
  const tenXMeanAndP95AgainstGoogleSheets =
    timingMetricPassedAgainstGoogleSheets &&
    postOperationFrameGuardrailPassed &&
    visualProofGuardrailPassed &&
    sourceWorkbookFingerprintGuardrailPassed
  const tenXMeanAndP95AgainstMicrosoftExcelWeb =
    timingMetricPassedAgainstMicrosoftExcelWeb === undefined
      ? undefined
      : timingMetricPassedAgainstMicrosoftExcelWeb &&
        postOperationFrameGuardrailPassed &&
        visualProofGuardrailPassed &&
        sourceWorkbookFingerprintGuardrailPassed
  return {
    id: captureCase.id,
    corpusCaseId: captureCase.corpusCaseId,
    materializedCells: captureCase.materializedCells,
    workload: captureCase.workload,
    sampleCount: Math.min(
      bilig.operationResponseMs.samples.length,
      googleSheets.operationResponseMs.samples.length,
      ...(microsoftExcelWeb ? [microsoftExcelWeb.operationResponseMs.samples.length] : []),
    ),
    bilig,
    googleSheets,
    ...(microsoftExcelWeb ? { microsoftExcelWeb } : {}),
    biligToGoogleSheetsMeanRatio,
    biligToGoogleSheetsP95Ratio,
    ...(biligToMicrosoftExcelWebMeanRatio !== undefined ? { biligToMicrosoftExcelWebMeanRatio } : {}),
    ...(biligToMicrosoftExcelWebP95Ratio !== undefined ? { biligToMicrosoftExcelWebP95Ratio } : {}),
    ...(requiresScrollEventMetric && scrollEventMetrics
      ? {
          biligToGoogleSheetsScrollEventMeanRatio: scrollEventMetrics.biligToGoogleSheetsMeanRatio,
          biligToGoogleSheetsScrollEventP95Ratio: scrollEventMetrics.biligToGoogleSheetsP95Ratio,
          ...(microsoftExcelWeb
            ? {
                biligToMicrosoftExcelWebScrollEventMeanRatio: scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio,
                biligToMicrosoftExcelWebScrollEventP95Ratio: scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio,
              }
            : {}),
          tenXMeanAndP95Metric: 'scrollEventResponseMs' as const,
          scrollMovementGuardrailPassed,
        }
      : { tenXMeanAndP95Metric: 'operationResponseMs' as const }),
    scenarioProof,
    postOperationFrameGuardrailPassed,
    sourceWorkbookFingerprintGuardrailPassed,
    tenXMeanAndP95AgainstGoogleSheets,
    ...(tenXMeanAndP95AgainstMicrosoftExcelWeb !== undefined ? { tenXMeanAndP95AgainstMicrosoftExcelWeb } : {}),
    passed: tenXMeanAndP95AgainstGoogleSheets,
  }
}

function scenarioProofVisualProofs(proof: SameCorpusScenarioProof): SameCorpusProductVisualProof[] {
  return proof.pixelGridProof.products.map((entry) => ({
    product: entry.product,
    screenshotPath: proof.screenshotProof.artifactPaths.find((artifact) => artifact.includes(`${entry.product}-`)) ?? null,
    screenshotCaptured: !proof.screenshotProof.missingProducts.includes(entry.product),
    pixelGridProof: entry,
  }))
}

function buildSameCorpusMeasurement(capture: SameCorpusCaptureMeasurement): UiResponsivenessSameCorpusMeasurement {
  return {
    product: capture.product,
    source: capture.source,
    operationResponseMs: summarizeNumbers(capture.operationResponseMsSamples),
    postOperationFrameMs: summarizeNumbers(capture.postOperationFrameMsSamples),
    ...(capture.scrollEventResponseMsSamples ? { scrollEventResponseMs: summarizeNumbers(capture.scrollEventResponseMsSamples) } : {}),
    ...(capture.scrollMovementPxSamples ? { scrollMovementPx: summarizeNumbers(capture.scrollMovementPxSamples) } : {}),
    corpusVerification: cloneSameCorpusVerification(capture.corpusVerification),
    limitations: [...capture.limitations],
  }
}

function sameCorpusScrollEventMetrics(
  bilig: UiResponsivenessSameCorpusMeasurement,
  googleSheets: UiResponsivenessSameCorpusMeasurement,
  microsoftExcelWeb?: UiResponsivenessSameCorpusMeasurement,
): {
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly biligToMicrosoftExcelWebMeanRatio: number
  readonly biligToMicrosoftExcelWebP95Ratio: number
} | null {
  if (
    !bilig.scrollEventResponseMs ||
    !googleSheets.scrollEventResponseMs ||
    (microsoftExcelWeb && !microsoftExcelWeb.scrollEventResponseMs)
  ) {
    return null
  }
  return {
    biligToGoogleSheetsMeanRatio: ratio(bilig.scrollEventResponseMs.mean, googleSheets.scrollEventResponseMs.mean),
    biligToGoogleSheetsP95Ratio: ratio(bilig.scrollEventResponseMs.p95, googleSheets.scrollEventResponseMs.p95),
    biligToMicrosoftExcelWebMeanRatio: microsoftExcelWeb
      ? ratio(bilig.scrollEventResponseMs.mean, microsoftExcelWeb.scrollEventResponseMs!.mean)
      : Number.POSITIVE_INFINITY,
    biligToMicrosoftExcelWebP95Ratio: microsoftExcelWeb
      ? ratio(bilig.scrollEventResponseMs.p95, microsoftExcelWeb.scrollEventResponseMs!.p95)
      : Number.POSITIVE_INFINITY,
  }
}

function ratio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}

function validateSameCorpusCase(entry: UiResponsivenessSameCorpusCase): void {
  if (entry.materializedCells <= 0 || !Number.isInteger(entry.materializedCells)) {
    throw new Error(`UI responsiveness same-corpus case has invalid materialized cell count: ${entry.id}`)
  }
  validateSameCorpusMeasurement(entry.bilig, 'bilig', entry.id, entry.corpusCaseId, entry.materializedCells)
  validateSameCorpusMeasurement(entry.googleSheets, 'google-sheets', entry.id, entry.corpusCaseId, entry.materializedCells)
  if (entry.microsoftExcelWeb) {
    validateSameCorpusMeasurement(entry.microsoftExcelWeb, 'microsoft-excel-web', entry.id, entry.corpusCaseId, entry.materializedCells)
  }
  if (
    uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload) &&
    ![entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])].every((measurement) =>
      hasSameCorpusScrollEvidence(measurement),
    )
  ) {
    throw new Error(`UI responsiveness same-corpus proof is missing scroll-event evidence for ${entry.id}`)
  }
  const comparableSampleCount = Math.min(
    entry.bilig.operationResponseMs.samples.length,
    entry.googleSheets.operationResponseMs.samples.length,
    ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb.operationResponseMs.samples.length] : []),
  )
  if (entry.sampleCount !== comparableSampleCount || comparableSampleCount < sameCorpusSampleCount) {
    throw new Error(`UI responsiveness same-corpus case has too few comparable samples: ${entry.id}`)
  }
  const googleSheetsMeanRatio = ratio(entry.bilig.operationResponseMs.mean, entry.googleSheets.operationResponseMs.mean)
  const googleSheetsP95Ratio = ratio(entry.bilig.operationResponseMs.p95, entry.googleSheets.operationResponseMs.p95)
  const microsoftExcelWebMeanRatio = entry.microsoftExcelWeb
    ? ratio(entry.bilig.operationResponseMs.mean, entry.microsoftExcelWeb.operationResponseMs.mean)
    : undefined
  const microsoftExcelWebP95Ratio = entry.microsoftExcelWeb
    ? ratio(entry.bilig.operationResponseMs.p95, entry.microsoftExcelWeb.operationResponseMs.p95)
    : undefined
  if (
    entry.biligToGoogleSheetsMeanRatio !== googleSheetsMeanRatio ||
    entry.biligToGoogleSheetsP95Ratio !== googleSheetsP95Ratio ||
    entry.biligToMicrosoftExcelWebMeanRatio !== microsoftExcelWebMeanRatio ||
    entry.biligToMicrosoftExcelWebP95Ratio !== microsoftExcelWebP95Ratio
  ) {
    throw new Error(`UI responsiveness same-corpus ratio is stale: ${entry.id}`)
  }
  const scrollEventMetrics = sameCorpusScrollEventMetrics(entry.bilig, entry.googleSheets, entry.microsoftExcelWeb)
  const comparedProducts = [entry.bilig, entry.googleSheets, ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : [])]
  const postOperationFrameGuardrailPassed = comparedProducts.every(
    (measurement) => measurement.postOperationFrameMs.p95 > 0 && measurement.postOperationFrameMs.p95 <= 50,
  )
  const sourceWorkbookFingerprintGuardrailPassed = comparedProducts.every(hasCapturedSourceWorkbookFingerprint)
  const scrollMovementGuardrailPassed =
    scrollEventMetrics !== null && comparedProducts.every((measurement) => (measurement.scrollMovementPx?.min ?? 0) >= 1)
  if (scrollEventMetrics) {
    if (
      entry.biligToGoogleSheetsScrollEventMeanRatio !== scrollEventMetrics.biligToGoogleSheetsMeanRatio ||
      entry.biligToGoogleSheetsScrollEventP95Ratio !== scrollEventMetrics.biligToGoogleSheetsP95Ratio ||
      (entry.microsoftExcelWeb &&
        (entry.biligToMicrosoftExcelWebScrollEventMeanRatio !== scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio ||
          entry.biligToMicrosoftExcelWebScrollEventP95Ratio !== scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio))
    ) {
      throw new Error(`UI responsiveness same-corpus scroll-event ratio is stale: ${entry.id}`)
    }
  }
  if (
    entry.postOperationFrameGuardrailPassed !== undefined &&
    entry.postOperationFrameGuardrailPassed !== postOperationFrameGuardrailPassed
  ) {
    throw new Error(`UI responsiveness same-corpus post-frame guardrail is stale: ${entry.id}`)
  }
  if (
    entry.sourceWorkbookFingerprintGuardrailPassed !== undefined &&
    entry.sourceWorkbookFingerprintGuardrailPassed !== sourceWorkbookFingerprintGuardrailPassed
  ) {
    throw new Error(`UI responsiveness same-corpus source workbook fingerprint guardrail is stale: ${entry.id}`)
  }
  if (entry.scrollMovementGuardrailPassed !== undefined && entry.scrollMovementGuardrailPassed !== scrollMovementGuardrailPassed) {
    throw new Error(`UI responsiveness same-corpus scroll-movement guardrail is stale: ${entry.id}`)
  }
  const requiresScrollEventMetric = uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload)
  const expectedMetric = requiresScrollEventMetric ? 'scrollEventResponseMs' : 'operationResponseMs'
  if (entry.tenXMeanAndP95Metric !== expectedMetric) {
    throw new Error(`UI responsiveness same-corpus metric is stale: ${entry.id}`)
  }
  validateSameCorpusScenarioProof(entry.scenarioProof, entry.id, entry.bilig, entry.googleSheets, entry.microsoftExcelWeb)
  const visualProofGuardrailPassed = entry.scenarioProof.screenshotProof.captured && entry.scenarioProof.pixelGridProof.captured
  const timingMetricPassedAgainstGoogleSheets = requiresScrollEventMetric
    ? scrollEventMetrics !== null &&
      scrollEventMetrics.biligToGoogleSheetsMeanRatio <= 0.1 &&
      scrollEventMetrics.biligToGoogleSheetsP95Ratio <= 0.1 &&
      scrollMovementGuardrailPassed
    : googleSheetsMeanRatio <= 0.1 && googleSheetsP95Ratio <= 0.1
  const timingMetricPassedAgainstMicrosoftExcelWeb = entry.microsoftExcelWeb
    ? requiresScrollEventMetric
      ? scrollEventMetrics !== null &&
        scrollEventMetrics.biligToMicrosoftExcelWebMeanRatio <= 0.1 &&
        scrollEventMetrics.biligToMicrosoftExcelWebP95Ratio <= 0.1 &&
        scrollMovementGuardrailPassed
      : (microsoftExcelWebMeanRatio ?? Number.POSITIVE_INFINITY) <= 0.1 && (microsoftExcelWebP95Ratio ?? Number.POSITIVE_INFINITY) <= 0.1
    : undefined
  const tenXAgainstGoogleSheets =
    timingMetricPassedAgainstGoogleSheets &&
    postOperationFrameGuardrailPassed &&
    visualProofGuardrailPassed &&
    sourceWorkbookFingerprintGuardrailPassed
  const tenXAgainstMicrosoftExcelWeb =
    timingMetricPassedAgainstMicrosoftExcelWeb === undefined
      ? undefined
      : timingMetricPassedAgainstMicrosoftExcelWeb &&
        postOperationFrameGuardrailPassed &&
        visualProofGuardrailPassed &&
        sourceWorkbookFingerprintGuardrailPassed
  if (
    entry.tenXMeanAndP95AgainstGoogleSheets !== tenXAgainstGoogleSheets ||
    entry.tenXMeanAndP95AgainstMicrosoftExcelWeb !== tenXAgainstMicrosoftExcelWeb ||
    entry.passed !== tenXAgainstGoogleSheets
  ) {
    throw new Error(`UI responsiveness same-corpus pass flag is stale: ${entry.id}`)
  }
}

function hasSameCorpusScrollEvidence(measurement: UiResponsivenessSameCorpusMeasurement): boolean {
  return Boolean(
    measurement.scrollEventResponseMs &&
    measurement.scrollMovementPx &&
    measurement.scrollEventResponseMs.samples.length >= sameCorpusSampleCount &&
    measurement.scrollMovementPx.samples.length >= sameCorpusSampleCount &&
    measurement.scrollMovementPx.min >= 1,
  )
}

function hasCapturedSourceWorkbookFingerprint(measurement: UiResponsivenessSameCorpusMeasurement): boolean {
  return measurement.corpusVerification.sourceWorkbookSha256 !== null && isSha256Hex(measurement.corpusVerification.sourceWorkbookSha256)
}

function validateSameCorpusMeasurement(
  measurement: UiResponsivenessSameCorpusMeasurement,
  product: UiResponsivenessSameCorpusProduct,
  caseId: string,
  corpusCaseId: string,
  materializedCells: number,
): void {
  if (measurement.product !== product) {
    throw new Error(`UI responsiveness same-corpus product mismatch for ${caseId}`)
  }
  if (measurement.source.length === 0) {
    throw new Error(`UI responsiveness same-corpus source is missing for ${caseId}`)
  }
  validateSummary(measurement.operationResponseMs, `${caseId} ${product} operationResponseMs`, sameCorpusSampleCount)
  validateSummary(measurement.postOperationFrameMs, `${caseId} ${product} postOperationFrameMs`, sameCorpusSampleCount)
  if (measurement.scrollEventResponseMs) {
    validateSummary(measurement.scrollEventResponseMs, `${caseId} ${product} scrollEventResponseMs`, sameCorpusSampleCount)
  }
  if (measurement.scrollMovementPx) {
    validateSummary(measurement.scrollMovementPx, `${caseId} ${product} scrollMovementPx`, sameCorpusSampleCount)
  }
  validateSameCorpusCaptureVerification(measurement.corpusVerification, product, materializedCells, corpusCaseId, caseId)
}

function validateSameCorpusCaptureVerification(
  verification: SameCorpusCaptureCorpusVerification,
  product: UiResponsivenessSameCorpusProduct,
  expectedMaterializedCells: number | null,
  expectedCorpusCaseId: string,
  caseId: string,
): void {
  if (!verification.verified) {
    throw new Error(`UI responsiveness same-corpus verification is not marked verified for ${caseId} ${product}`)
  }
  if (expectedMaterializedCells !== null && verification.materializedCells !== expectedMaterializedCells) {
    throw new Error(`UI responsiveness same-corpus verification materialized cell count mismatch for ${caseId} ${product}`)
  }
  validateSameCorpusCaptureCorpusFingerprint(
    verification.corpusFingerprint,
    expectedCorpusCaseId,
    expectedMaterializedCells,
    caseId,
    product,
  )
  if (verification.sourceWorkbookSha256 !== null && !isSha256Hex(verification.sourceWorkbookSha256)) {
    throw new Error(`UI responsiveness same-corpus verification source workbook fingerprint is invalid for ${caseId} ${product}`)
  }
  if (
    product === 'bilig' &&
    verification.sourceWorkbookSha256 !== null &&
    verification.sourceWorkbookSha256 !== verification.corpusFingerprint.snapshotSha256
  ) {
    throw new Error(`UI responsiveness same-corpus Bilig source workbook fingerprint is stale for ${caseId}`)
  }
  if (product === 'bilig' && verification.method !== 'bilig-benchmark-state') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'google-sheets' && verification.method !== 'google-sheets-xlsx-export') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'microsoft-excel-web' && verification.method !== 'microsoft-excel-web-source-xlsx') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (verification.checkedCells.length < 3) {
    throw new Error(`UI responsiveness same-corpus verification must check at least 3 cells for ${caseId} ${product}`)
  }
  for (const cell of verification.checkedCells) {
    if (cell.address.trim().length === 0 || cell.expected !== cell.actual) {
      throw new Error(`UI responsiveness same-corpus verification cell mismatch for ${caseId} ${product}`)
    }
  }
}

function validateSameCorpusCaptureCorpusFingerprint(
  fingerprint: SameCorpusCaptureCorpusFingerprint,
  expectedCorpusCaseId: string,
  expectedMaterializedCells: number | null,
  caseId: string,
  product: UiResponsivenessSameCorpusProduct,
): void {
  if (fingerprint.version !== sameCorpusFingerprintVersion) {
    throw new Error(`UI responsiveness same-corpus verification fingerprint version is stale for ${caseId} ${product}`)
  }
  if (fingerprint.corpusCaseId !== expectedCorpusCaseId) {
    throw new Error(`UI responsiveness same-corpus verification corpus fingerprint mismatch for ${caseId} ${product}`)
  }
  if (expectedMaterializedCells !== null && fingerprint.materializedCells !== expectedMaterializedCells) {
    throw new Error(`UI responsiveness same-corpus verification fingerprint materialized cell count mismatch for ${caseId} ${product}`)
  }
  if (!isSha256Hex(fingerprint.snapshotSha256)) {
    throw new Error(`UI responsiveness same-corpus verification benchmark fingerprint is invalid for ${caseId} ${product}`)
  }
  if (!isWorkbookBenchmarkCorpusId(fingerprint.corpusCaseId)) {
    throw new Error(`UI responsiveness same-corpus verification uses unknown corpus fingerprint for ${caseId} ${product}`)
  }
  const expectedFingerprint = expectedCorpusFingerprint(fingerprint.corpusCaseId)
  if (JSON.stringify(fingerprint) !== JSON.stringify(expectedFingerprint)) {
    throw new Error(`UI responsiveness same-corpus verification benchmark fingerprint is stale for ${caseId} ${product}`)
  }
}

function expectedCorpusFingerprint(corpusId: WorkbookBenchmarkCorpusId): SameCorpusCaptureCorpusFingerprint {
  const cached = expectedCorpusFingerprintCache.get(corpusId)
  if (cached) {
    return cached
  }
  const fingerprint = buildSameCorpusFingerprint(buildWorkbookBenchmarkCorpus(corpusId)).corpusFingerprint
  expectedCorpusFingerprintCache.set(corpusId, fingerprint)
  return fingerprint
}
