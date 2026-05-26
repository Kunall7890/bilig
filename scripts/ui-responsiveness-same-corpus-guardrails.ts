import type {
  SameCorpusCaptureCase,
  SameCorpusCaptureMeasurement,
  SameCorpusOperationResponseProof,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'
import {
  validateSameCorpusMutationTargetProofSample,
  type SameCorpusMutationTargetProof,
  type SameCorpusProductSemanticUiProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadMutatesWorkbook,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

export const sameCorpusSampleCount = 3

export interface SameCorpusCommittedTargetProofTimingMeasurement {
  readonly product: UiResponsivenessSameCorpusProduct
  readonly committedTargetProofMs?: { readonly samples: readonly number[] } | undefined
  readonly committedTargetProofMsSamples?: readonly number[] | undefined
}

export interface SameCorpusCommittedTargetProofTimingCase {
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly scenarioProof: {
    readonly semanticUiProof: {
      readonly products: readonly SameCorpusProductSemanticUiProof[]
    }
  }
  readonly bilig: SameCorpusCommittedTargetProofTimingMeasurement
  readonly googleSheets: SameCorpusCommittedTargetProofTimingMeasurement
}

export function sameCorpusCaseOperationResponseProofGuardrailPassed(entry: UiResponsivenessSameCorpusCase): boolean {
  return sameCorpusMeasurementOperationResponseProofsPassed(entry.workload, entry.sampleCount, [
    entry.bilig,
    entry.googleSheets,
    ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : []),
  ])
}

export function sameCorpusCaptureCaseOperationResponseProofGuardrailPassed(entry: SameCorpusCaptureCase, sampleCount: number): boolean {
  return sameCorpusMeasurementOperationResponseProofsPassed(entry.workload, sampleCount, [
    entry.bilig,
    entry.googleSheets,
    ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb] : []),
  ])
}

export function sameCorpusMeasurementOperationResponseProofsPassed(
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
  measurements: readonly Pick<SameCorpusCaptureMeasurement | UiResponsivenessSameCorpusMeasurement, 'operationResponseProofs'>[],
): boolean {
  const expected = expectedSameCorpusOperationResponseProof(workload)
  return measurements.every(
    (measurement) =>
      measurement.operationResponseProofs.length >= sampleCount && measurement.operationResponseProofs.every((proof) => proof === expected),
  )
}

function expectedSameCorpusOperationResponseProof(workload: UiResponsivenessSameCorpusWorkload): SameCorpusOperationResponseProof {
  if (workload === 'open-workbook') {
    return 'load-to-ready'
  }
  if (uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
    return 'visible-scroll-movement'
  }
  return 'visible-non-scroll-response'
}

export function hasBiligAuthoritativeRenderProofTiming(
  measurement: Pick<
    SameCorpusCaptureMeasurement | UiResponsivenessSameCorpusMeasurement,
    'product' | 'authoritativeRenderProofMsSamples' | 'authoritativeRenderProofMs'
  >,
  sampleCount: number,
): boolean {
  if (measurement.product !== 'bilig') {
    return false
  }
  if ('authoritativeRenderProofMs' in measurement && measurement.authoritativeRenderProofMs) {
    return (
      measurement.authoritativeRenderProofMs.samples.length >= sampleCount &&
      measurement.authoritativeRenderProofMs.samples.every((value) => Number.isFinite(value) && value >= 0)
    )
  }
  const samples = 'authoritativeRenderProofMsSamples' in measurement ? measurement.authoritativeRenderProofMsSamples : undefined
  return Boolean(samples && samples.length >= sampleCount && samples.every((value) => Number.isFinite(value) && value >= 0))
}

export function hasCommittedTargetProofTiming(measurement: SameCorpusCommittedTargetProofTimingMeasurement, sampleCount: number): boolean {
  if ('committedTargetProofMs' in measurement && measurement.committedTargetProofMs) {
    return (
      measurement.committedTargetProofMs.samples.length >= sampleCount &&
      measurement.committedTargetProofMs.samples.every((value) => Number.isFinite(value) && value >= 0)
    )
  }
  const samples = 'committedTargetProofMsSamples' in measurement ? measurement.committedTargetProofMsSamples : undefined
  return Boolean(samples && samples.length >= sampleCount && samples.every((value) => Number.isFinite(value) && value >= 0))
}

export function hasAcceptedCommittedTargetProofTiming(
  entry: Pick<SameCorpusCommittedTargetProofTimingCase, 'scenarioProof' | 'workload'>,
  measurement: SameCorpusCommittedTargetProofTimingMeasurement,
  sampleCount: number,
): boolean {
  return acceptedCommittedTargetProofTimingSampleCount(entry, measurement, sampleCount) === sampleCount
}

export function requiredUiResponsivenessSameCorpusCommittedTargetProofTimingCaseCount(): number {
  return requiredUiResponsivenessSameCorpusWorkloads.filter(uiSameCorpusWorkloadMutatesWorkbook).length
}

export function requiredUiResponsivenessSameCorpusCommittedTargetProofTimingSampleCount(sampleCount: number): number {
  return requiredUiResponsivenessSameCorpusCommittedTargetProofTimingCaseCount() * 2 * Math.max(0, sampleCount)
}

export function sameCorpusCommittedTargetProofTimingCounts(
  cases: readonly SameCorpusCommittedTargetProofTimingCase[],
  sampleCount: number,
): {
  readonly requiredCommittedTargetProofTimingCaseCount: number
  readonly committedTargetProofTimingCaseCount: number
  readonly requiredCommittedTargetProofTimingSampleCount: number
  readonly committedTargetProofTimingSampleCount: number
} {
  return {
    requiredCommittedTargetProofTimingCaseCount: requiredUiResponsivenessSameCorpusCommittedTargetProofTimingCaseCount(),
    committedTargetProofTimingCaseCount: sameCorpusCommittedTargetProofTimingCaseCount(cases, sampleCount),
    requiredCommittedTargetProofTimingSampleCount: requiredUiResponsivenessSameCorpusCommittedTargetProofTimingSampleCount(sampleCount),
    committedTargetProofTimingSampleCount: sameCorpusCommittedTargetProofTimingSampleCount(cases, sampleCount),
  }
}

export function sameCorpusCommittedTargetProofTimingCaseCount(
  cases: readonly SameCorpusCommittedTargetProofTimingCase[],
  sampleCount: number,
): number {
  return requiredUiResponsivenessSameCorpusWorkloads.filter(
    (workload) =>
      uiSameCorpusWorkloadMutatesWorkbook(workload) &&
      cases.some(
        (entry) =>
          entry.workload === workload &&
          [entry.bilig, entry.googleSheets].every((measurement) => hasAcceptedCommittedTargetProofTiming(entry, measurement, sampleCount)),
      ),
  ).length
}

export function sameCorpusCommittedTargetProofTimingSampleCount(
  cases: readonly SameCorpusCommittedTargetProofTimingCase[],
  sampleCount: number,
): number {
  return requiredUiResponsivenessSameCorpusWorkloads.reduce((total, workload) => {
    if (!uiSameCorpusWorkloadMutatesWorkbook(workload)) {
      return total
    }
    const entry = cases.find((candidate) => candidate.workload === workload)
    if (!entry) {
      return total
    }
    return (
      total +
      acceptedCommittedTargetProofTimingSampleCount(entry, entry.bilig, sampleCount) +
      acceptedCommittedTargetProofTimingSampleCount(entry, entry.googleSheets, sampleCount)
    )
  }, 0)
}

function acceptedCommittedTargetProofTimingSampleCount(
  entry: Pick<SameCorpusCommittedTargetProofTimingCase, 'scenarioProof' | 'workload'>,
  measurement: SameCorpusCommittedTargetProofTimingMeasurement,
  sampleCount: number,
): number {
  if (!uiSameCorpusWorkloadMutatesWorkbook(entry.workload) || !hasCommittedTargetProofTiming(measurement, sampleCount)) {
    return 0
  }
  const productProof = sameCorpusSemanticProductProof(entry.scenarioProof, measurement.product)
  if (!productProof) {
    return 0
  }
  if (!sameCorpusSemanticProductContainerAccepted(productProof)) {
    return 0
  }
  const duplicateSampleIndexes = sameCorpusDuplicateMutationSampleIndexes(productProof.mutationTargetProofs)
  const duplicateScreenshotPaths = sameCorpusDuplicateMutationScreenshotPaths(productProof.mutationTargetProofs)
  let acceptedSampleCount = 0
  for (let sampleIndex = 0; sampleIndex < sampleCount; sampleIndex += 1) {
    const sample = productProof.mutationTargetProofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!sample || !committedTargetProofTimingMatchesMeasurement(measurement, sample, sampleIndex)) {
      continue
    }
    const verdict = validateSameCorpusMutationTargetProofSample(productProof, entry.workload, sample, {
      duplicateSampleIndex: duplicateSampleIndexes.has(sample.sampleIndex),
      duplicateScreenshotPath:
        sample.screenshotPath !== null &&
        duplicateScreenshotPaths.has(normalizeSameCorpusMutationTargetScreenshotPath(sample.screenshotPath)),
      sampleCount,
    })
    if (verdict.acceptedForCurrentScorecard) {
      acceptedSampleCount += 1
    }
  }
  return acceptedSampleCount
}

function committedTargetProofTimingMatchesMeasurement(
  measurement: SameCorpusCommittedTargetProofTimingMeasurement,
  sample: SameCorpusMutationTargetProof,
  sampleIndex: number,
): boolean {
  const samples = committedTargetProofTimingSamples(measurement)
  if (!samples || samples.length <= sampleIndex) {
    return false
  }
  const timing =
    'committedTargetProofMsSamples' in measurement
      ? samples[sampleIndex]
      : samples.find((value) => sameCorpusTimingValuesMatch(value, sample.committedTargetProofMs))
  return timing !== undefined && sameCorpusTimingValuesMatch(timing, sample.committedTargetProofMs)
}

function committedTargetProofTimingSamples(measurement: SameCorpusCommittedTargetProofTimingMeasurement): readonly number[] | undefined {
  if ('committedTargetProofMs' in measurement && measurement.committedTargetProofMs) {
    return measurement.committedTargetProofMs.samples
  }
  if ('committedTargetProofMsSamples' in measurement) {
    return measurement.committedTargetProofMsSamples
  }
  return undefined
}

function sameCorpusTimingValuesMatch(left: number, right: number): boolean {
  return Number.isFinite(left) && left >= 0 && Number.isFinite(right) && right >= 0 && Math.abs(left - right) <= 1
}

function sameCorpusSemanticProductProof(
  scenarioProof: SameCorpusCommittedTargetProofTimingCase['scenarioProof'],
  product: UiResponsivenessSameCorpusProduct,
): SameCorpusProductSemanticUiProof | undefined {
  return scenarioProof.semanticUiProof.products.find((entry) => entry.product === product)
}

function sameCorpusSemanticProductContainerAccepted(proof: SameCorpusProductSemanticUiProof): boolean {
  if (!proof.captured) {
    return false
  }
  if (proof.sheetName.trim().length === 0 || proof.sheetId === null || proof.sheetId.trim().length === 0) {
    return false
  }
  if (proof.selectedRange === null || proof.selectedRange.trim().length === 0) {
    return false
  }
  if (proof.screenshotSha256 === null || !/^[a-f0-9]{64}$/u.test(proof.screenshotSha256)) {
    return false
  }
  if (
    proof.checkedCells.length < 3 ||
    proof.checkedCells.some((cell) => cell.address.trim().length === 0 || cell.expected !== cell.actual)
  ) {
    return false
  }
  if (proof.product === 'bilig') {
    return (
      proof.method === 'bilig-visible-semantic-readback' &&
      proof.authoritativeRenderRevision !== null &&
      proof.authoritativeRenderRevision.trim().length > 0 &&
      proof.visibleRenderRevision !== null &&
      proof.visibleRenderRevision.trim().length > 0
    )
  }
  if (proof.product === 'google-sheets') {
    return proof.method === 'google-sheets-visible-semantic-readback'
  }
  return proof.method === 'excel-web-visible-semantic-readback'
}

function sameCorpusDuplicateMutationSampleIndexes(samples: readonly SameCorpusMutationTargetProof[]): Set<number> {
  const seen = new Set<number>()
  const duplicates = new Set<number>()
  for (const sample of samples) {
    if (seen.has(sample.sampleIndex)) {
      duplicates.add(sample.sampleIndex)
    }
    seen.add(sample.sampleIndex)
  }
  return duplicates
}

function sameCorpusDuplicateMutationScreenshotPaths(samples: readonly SameCorpusMutationTargetProof[]): Set<string> {
  const seen = new Set<string>()
  const duplicates = new Set<string>()
  for (const sample of samples) {
    if (sample.screenshotPath === null || sample.screenshotPath === undefined || sample.screenshotPath.trim().length === 0) {
      continue
    }
    const screenshotPath = normalizeSameCorpusMutationTargetScreenshotPath(sample.screenshotPath)
    if (seen.has(screenshotPath)) {
      duplicates.add(screenshotPath)
    }
    seen.add(screenshotPath)
  }
  return duplicates
}

function normalizeSameCorpusMutationTargetScreenshotPath(path: string): string {
  return path.trim().replaceAll('\\', '/')
}

export function sameCorpusCommittedTargetProofMetrics(
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
    !bilig.committedTargetProofMs ||
    !googleSheets.committedTargetProofMs ||
    (microsoftExcelWeb && !microsoftExcelWeb.committedTargetProofMs)
  ) {
    return null
  }
  return {
    biligToGoogleSheetsMeanRatio: sameCorpusTimingRatio(bilig.committedTargetProofMs.mean, googleSheets.committedTargetProofMs.mean),
    biligToGoogleSheetsP95Ratio: sameCorpusTimingRatio(bilig.committedTargetProofMs.p95, googleSheets.committedTargetProofMs.p95),
    biligToMicrosoftExcelWebMeanRatio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.committedTargetProofMs.mean, microsoftExcelWeb.committedTargetProofMs!.mean)
      : Number.POSITIVE_INFINITY,
    biligToMicrosoftExcelWebP95Ratio: microsoftExcelWeb
      ? sameCorpusTimingRatio(bilig.committedTargetProofMs.p95, microsoftExcelWeb.committedTargetProofMs!.p95)
      : Number.POSITIVE_INFINITY,
  }
}

export function hasBiligProductionRuntimeProof(
  measurement: Pick<SameCorpusCaptureMeasurement | UiResponsivenessSameCorpusMeasurement, 'product' | 'biligRuntimeProof'>,
): boolean {
  const proof = measurement.biligRuntimeProof
  return (
    measurement.product === 'bilig' &&
    proof !== undefined &&
    proof.product === 'bilig' &&
    proof.requiredBuildKind === 'production' &&
    proof.actualBuildKind === 'production' &&
    proof.prod &&
    !proof.dev &&
    proof.verified &&
    proof.sampleCount >= sameCorpusSampleCount &&
    proof.samples.length >= sameCorpusSampleCount &&
    proof.samples.every((sample) => sample.present && sample.buildKind === 'production' && sample.prod && !sample.dev)
  )
}

function sameCorpusTimingRatio(numerator: number, denominator: number): number {
  if (denominator <= 0) {
    return Number.POSITIVE_INFINITY
  }
  return numerator / denominator
}
