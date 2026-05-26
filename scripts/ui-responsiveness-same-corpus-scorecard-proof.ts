import { createHash } from 'node:crypto'

import { summarizeNumbers } from '../packages/benchmarks/src/stats.js'
import type { SameCorpusCaptureCorpusFingerprint } from './ui-responsiveness-same-corpus-fingerprint.ts'
import type { SameCorpusProductVisualProof, SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import {
  buildScorecardScenarioProof,
  sameCorpusUiRenderProofContractVersion,
  validateSameCorpusScenarioProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import {
  sameCorpusUiCaptureToolVersion,
  type SameCorpusCapture,
  type SameCorpusBiligRuntimeProof,
  type SameCorpusCaptureCase,
  type SameCorpusCaptureCorpusVerification,
  type SameCorpusCaptureMeasurement,
  type SameCorpusCaptureRunManifest,
  type SameCorpusScenarioCaseFields,
  type SameCorpusProductSourceWorkbookFingerprint,
  type UiResponsivenessSameCorpusCase,
  type UiResponsivenessSameCorpusMeasurement,
  type UiResponsivenessSameCorpusProduct,
  type UiResponsivenessSameCorpusProof,
  type UiResponsivenessSameCorpusRunManifest,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'
import {
  hasBiligAuthoritativeRenderProofTiming,
  hasBiligProductionRuntimeProof,
  sameCorpusCaptureCaseOperationResponseProofGuardrailPassed,
  sameCorpusCaseOperationResponseProofGuardrailPassed,
  sameCorpusMeasurementOperationResponseProofsPassed,
  sameCorpusSampleCount,
} from './ui-responsiveness-same-corpus-guardrails.ts'
import {
  sameCorpusScenarioCaseFields,
  sameCorpusScenarioSummaryFieldsCurrent,
  validateSameCorpusScenarioCaseFields,
} from './ui-responsiveness-same-corpus-scenario-fields.ts'
import {
  requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount,
  requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount,
  sameCorpusMutationTargetProofCaseCount,
  sameCorpusMutationTargetProofProductSummaries,
  sameCorpusMutationTargetProofSampleCount,
} from './ui-responsiveness-same-corpus-mutation-target-proof-summary.ts'
import { sameCorpusManifestInvalidReasons } from './ui-responsiveness-same-corpus-manifest-invalid-reasons.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
} from './ui-responsiveness-same-corpus-workloads.ts'
import {
  cloneSameCorpusVerification,
  isSha256Hex,
  validateSameCorpusCaptureVerification,
  validateSummary,
} from './ui-responsiveness-same-corpus-validation-helpers.ts'

export { sameCorpusUiCaptureToolVersion } from './ui-responsiveness-same-corpus-scorecard-types.ts'
export { sameCorpusScenarioCaseFields } from './ui-responsiveness-same-corpus-scenario-fields.ts'
export type {
  SameCorpusCapture,
  SameCorpusBiligRuntimeProof,
  SameCorpusBiligRuntimeProofSample,
  SameCorpusCaptureCase,
  SameCorpusCaptureCorpusVerification,
  SameCorpusCaptureMeasurement,
  SameCorpusCaptureRunManifest,
  SameCorpusCaptureVerifiedCell,
  SameCorpusOperationResponseProof,
  SameCorpusScenarioCaseFields,
  SameCorpusProductSourceWorkbookFingerprint,
  UiResponsivenessSameCorpusCase,
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
  UiResponsivenessSameCorpusProof,
  UiResponsivenessSameCorpusRunManifest,
} from './ui-responsiveness-same-corpus-scorecard-types.ts'

const requiredSameCorpusWorkloads = requiredUiResponsivenessSameCorpusWorkloads
const strictRenderedGridProofLimitation =
  'Some same-corpus cases retain timing evidence but do not satisfy strict rendered-grid proof, so they cannot count toward Google Sheets 10x UI claims.'
const authoritativeRenderProofTimingLimitation =
  'Bilig interaction-visible timing is tracked separately from authoritative rendered-proof completion; captures missing authoritativeRenderProofMsSamples cannot satisfy the current same-corpus evidence contract.'

export function buildMissingSameCorpusProof(): UiResponsivenessSameCorpusProof {
  return {
    captured: false,
    evidenceKind: 'not-captured',
    requiredProductCount: 2,
    requiredCaseCount: requiredSameCorpusWorkloads.length,
    tenXMeanAndP95CaseCount: 0,
    coveredCorpusCaseIds: [],
    runManifest: buildSameCorpusRunManifest([], null),
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
    runManifest: buildSameCorpusRunManifest(cases, capture.runManifest.captureRunSignature),
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

function buildSameCorpusRunManifest(
  cases: readonly UiResponsivenessSameCorpusCase[],
  captureRunSignature: string | null,
): UiResponsivenessSameCorpusRunManifest {
  const capturedWorkloads = cases.map((entry) => entry.workload)
  const corpusCaseIds = [...new Set(cases.map((entry) => entry.corpusCaseId))].toSorted()
  const corpusFingerprints = uniqueCorpusFingerprints(cases)
  const productSourceWorkbookFingerprints = uniqueProductSourceWorkbookFingerprints(cases)
  const materializedCellCounts = [...new Set(cases.map((entry) => entry.materializedCells))].toSorted((left, right) => left - right)
  const biligProductionRuntimeProofCaseCount = cases.filter((entry) => hasBiligProductionRuntimeProof(entry.bilig)).length
  const scenarioSummaryFieldCaseCount = cases.filter(sameCorpusScenarioSummaryFieldsCurrent).length
  const strictRenderedGridProofCaseCount = cases.filter((entry) => entry.scenarioProof.pixelGridProof.captured).length
  const visibleOperationResponseProofCaseCount = cases.filter((entry) => sameCorpusCaseOperationResponseProofGuardrailPassed(entry)).length
  const biligAuthoritativeRenderProofCaseCount = cases.filter((entry) =>
    hasBiligAuthoritativeRenderProofTiming(entry.bilig, entry.sampleCount),
  ).length
  const semanticUiProofCaseCount = cases.filter((entry) => entry.scenarioProof.semanticUiProof.captured).length
  const requiredMutationTargetProofCaseCount = requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount()
  const sampleCount = manifestSampleCount(cases)
  const mutationTargetProofCaseCount = sameCorpusMutationTargetProofCaseCount(cases, sampleCount)
  const requiredMutationTargetProofSampleCount = requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount(sampleCount)
  const mutationTargetProofSampleCount = sameCorpusMutationTargetProofSampleCount(cases, sampleCount)
  const mutationTargetProofProductSummaries = sameCorpusMutationTargetProofProductSummaries(cases, sampleCount)
  const legacyInsufficientRenderedGridProofCaseCount = cases.filter((entry) =>
    entry.scenarioProof.pixelGridProof.productVerdicts.some((verdict) => verdict.evidenceStatus === 'legacy-insufficient'),
  ).length
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length
  const invalidReasons = sameCorpusManifestInvalidReasons({
    capturedWorkloads,
    caseCount: cases.length,
    corpusCaseIds,
    corpusFingerprints,
    productSourceWorkbookFingerprints,
    biligProductionRuntimeProofCaseCount,
    scenarioSummaryFieldCaseCount,
    biligAuthoritativeRenderProofCaseCount,
    semanticUiProofCaseCount,
    requiredMutationTargetProofCaseCount,
    mutationTargetProofCaseCount,
    requiredMutationTargetProofSampleCount,
    mutationTargetProofSampleCount,
    legacyInsufficientRenderedGridProofCaseCount,
    materializedCellCounts,
    strictRenderedGridProofCaseCount,
    visibleOperationResponseProofCaseCount,
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
    biligProductionRuntimeProofCaseCount,
    sampleCount,
    caseCount: cases.length,
    scenarioSummaryFieldCaseCount,
    strictRenderedGridProofCaseCount,
    visibleOperationResponseProofCaseCount,
    biligAuthoritativeRenderProofCaseCount,
    semanticUiProofCaseCount,
    requiredMutationTargetProofCaseCount,
    mutationTargetProofCaseCount,
    requiredMutationTargetProofSampleCount,
    mutationTargetProofSampleCount,
    mutationTargetProofProductSummaries,
    legacyInsufficientRenderedGridProofCaseCount,
    tenXMeanAndP95CaseCount,
    currentContractEvidenceComplete: !invalidReasons.some(
      (reason) => reason !== 'not every required workload is 10x against Google Sheets',
    ),
    googleSheetsTenXRequirementSatisfied: invalidReasons.length === 0,
    captureRunSignature,
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
  const biligProductionRuntimeProofCaseCount = cases.filter((entry) => hasBiligProductionRuntimeProof(entry.bilig)).length
  const scenarioSummaryFieldCaseCount = cases.filter(sameCorpusScenarioSummaryFieldsCurrent).length
  const strictRenderedGridProofCaseCount = cases.filter((entry) => entry.scenarioProof.pixelGridProof.captured).length
  const visibleOperationResponseProofCaseCount = cases.filter((entry) =>
    sameCorpusCaptureCaseOperationResponseProofGuardrailPassed(entry, sampleCount),
  ).length
  const biligAuthoritativeRenderProofCaseCount = cases.filter((entry) =>
    hasBiligAuthoritativeRenderProofTiming(entry.bilig, sampleCount),
  ).length
  const semanticUiProofCaseCount = cases.filter((entry) => entry.scenarioProof.semanticUiProof.captured).length
  const requiredMutationTargetProofCaseCount = requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount()
  const requiredMutationTargetProofSampleCount = requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount(sampleCount)
  const mutationTargetProofCaseCount = sameCorpusMutationTargetProofCaseCount(cases, sampleCount)
  const mutationTargetProofSampleCount = sameCorpusMutationTargetProofSampleCount(cases, sampleCount)
  const mutationTargetProofProductSummaries = sameCorpusMutationTargetProofProductSummaries(cases, sampleCount)
  const legacyInsufficientRenderedGridProofCaseCount = cases.filter((entry) =>
    entry.scenarioProof.pixelGridProof.productVerdicts.some((verdict) => verdict.evidenceStatus === 'legacy-insufficient'),
  ).length
  const tenXMeanAndP95CaseCount = cases.map(buildSameCorpusCase).filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length
  const invalidReasons = sameCorpusManifestInvalidReasons({
    capturedWorkloads,
    caseCount: cases.length,
    corpusCaseIds,
    corpusFingerprints,
    productSourceWorkbookFingerprints,
    biligProductionRuntimeProofCaseCount,
    scenarioSummaryFieldCaseCount,
    biligAuthoritativeRenderProofCaseCount,
    semanticUiProofCaseCount,
    requiredMutationTargetProofCaseCount,
    mutationTargetProofCaseCount,
    requiredMutationTargetProofSampleCount,
    mutationTargetProofSampleCount,
    legacyInsufficientRenderedGridProofCaseCount,
    materializedCellCounts,
    strictRenderedGridProofCaseCount,
    visibleOperationResponseProofCaseCount,
    tenXMeanAndP95CaseCount,
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
    biligProductionRuntimeProofCaseCount,
    sampleCount,
    caseCount: cases.length,
    scenarioSummaryFieldCaseCount,
    strictRenderedGridProofCaseCount,
    visibleOperationResponseProofCaseCount,
    biligAuthoritativeRenderProofCaseCount,
    semanticUiProofCaseCount,
    requiredMutationTargetProofCaseCount,
    mutationTargetProofCaseCount,
    requiredMutationTargetProofSampleCount,
    mutationTargetProofSampleCount,
    mutationTargetProofProductSummaries,
    legacyInsufficientRenderedGridProofCaseCount,
    tenXMeanAndP95CaseCount,
    currentContractEvidenceComplete: !invalidReasons.some(
      (reason) => reason !== 'not every required workload is 10x against Google Sheets',
    ),
    googleSheetsTenXRequirementSatisfied: invalidReasons.length === 0,
    captureRunSignature: sameCorpusCaptureRunSignature(cases),
    invalidReasons,
  }
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

function manifestSampleCount(cases: readonly UiResponsivenessSameCorpusCase[]): number {
  return cases.length === 0 ? 0 : Math.min(...cases.map((entry) => entry.sampleCount))
}

function captureCaseSampleCount(entry: SameCorpusCaptureCase): number {
  return Math.min(
    entry.bilig.operationResponseMsSamples.length,
    entry.googleSheets.operationResponseMsSamples.length,
    ...(entry.microsoftExcelWeb ? [entry.microsoftExcelWeb.operationResponseMsSamples.length] : []),
  )
}

function validateSameCorpusRunManifest(proof: UiResponsivenessSameCorpusProof): void {
  if (!proof.runManifest) {
    throw new Error('UI responsiveness same-corpus proof is missing run manifest')
  }
  if (proof.captured) {
    if (!isSha256Hex(proof.runManifest.captureRunSignature)) {
      throw new Error('UI responsiveness same-corpus proof run manifest must bind to a captureRunSignature')
    }
  } else if (proof.runManifest.captureRunSignature !== null) {
    throw new Error('UI responsiveness missing same-corpus proof must not claim a captureRunSignature')
  }
  const expected = buildSameCorpusRunManifest(proof.cases, proof.runManifest.captureRunSignature)
  if (stableJsonString(proof.runManifest) !== stableJsonString(expected)) {
    throw new Error('UI responsiveness same-corpus run manifest is stale')
  }
}

function sameCorpusProofLimitations(captureLimitations: readonly string[], cases: readonly UiResponsivenessSameCorpusCase[]): string[] {
  const limitations = [...captureLimitations]
  if (cases.some((entry) => !entry.scenarioProof.pixelGridProof.captured) && !limitations.includes(strictRenderedGridProofLimitation)) {
    limitations.push(strictRenderedGridProofLimitation)
  }
  if (
    cases.some((entry) => !hasBiligAuthoritativeRenderProofTiming(entry.bilig, entry.sampleCount)) &&
    !limitations.includes(authoritativeRenderProofTimingLimitation)
  ) {
    limitations.push(authoritativeRenderProofTimingLimitation)
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
        measurement.operationResponseProofs.length < capture.sampleCount ||
        measurement.postOperationFrameMsSamples.length < capture.sampleCount
      ) {
        throw new Error(`UI responsiveness same-corpus capture has too few samples for ${entry.id}`)
      }
      if (measurement.product === 'bilig' && measurement.authoritativeRenderProofMsSamples) {
        if (measurement.authoritativeRenderProofMsSamples.length < capture.sampleCount) {
          throw new Error(
            `UI responsiveness same-corpus capture has too few Bilig authoritative render proof timing samples for ${entry.id}`,
          )
        }
        for (const value of measurement.authoritativeRenderProofMsSamples) {
          if (!Number.isFinite(value) || value < 0) {
            throw new Error(`UI responsiveness same-corpus capture has invalid Bilig authoritative render proof timing for ${entry.id}`)
          }
        }
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
      if (measurement.product === 'bilig' && measurement.biligRuntimeProof) {
        validateBiligRuntimeProof(measurement.biligRuntimeProof, measurement.source, entry.id)
      }
    }
    validateSameCorpusScenarioProof(
      entry.scenarioProof,
      entry.id,
      buildSameCorpusMeasurement(entry.bilig),
      buildSameCorpusMeasurement(entry.googleSheets),
      entry.workload,
      entry.microsoftExcelWeb ? buildSameCorpusMeasurement(entry.microsoftExcelWeb) : undefined,
    )
    validateSameCorpusScenarioCaseFields(entry, 'capture')
  }
}

export function validateSameCorpusCaptureRunManifest(capture: SameCorpusCapture): void {
  const expected = buildSameCorpusCaptureRunManifest(capture.cases, capture.sampleCount)
  if (stableJsonString(capture.runManifest) !== stableJsonString(expected)) {
    throw new Error('UI responsiveness same-corpus capture run manifest is stale')
  }
}

function sameCorpusCaptureRunSignature(cases: readonly SameCorpusCaptureCase[]): string {
  return createHash('sha256').update(stableJsonString(cases)).digest('hex')
}

function stableJsonString(value: unknown): string {
  return JSON.stringify(stableJsonValue(value))
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, stableJsonValue(entry)]),
    )
  }
  return value
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
  const operationResponseProofGuardrailPassed = sameCorpusMeasurementOperationResponseProofsPassed(
    captureCase.workload,
    captureCaseSampleCount(captureCase),
    [captureCase.bilig, captureCase.googleSheets, ...(captureCase.microsoftExcelWeb ? [captureCase.microsoftExcelWeb] : [])],
  )
  const sourceWorkbookFingerprintGuardrailPassed = comparedProducts.every(hasCapturedSourceWorkbookFingerprint)
  const biligRuntimeProofGuardrailPassed = hasBiligProductionRuntimeProof(bilig)
  const authoritativeRenderProofGuardrailPassed = hasBiligAuthoritativeRenderProofTiming(bilig, captureCaseSampleCount(captureCase))
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
    workload: captureCase.workload,
    sampleCount: captureCaseSampleCount(captureCase),
    visualProofs: scenarioProofVisualProofs(captureCase.scenarioProof),
  })
  const visualProofGuardrailPassed = scenarioProof.screenshotProof.captured && scenarioProof.pixelGridProof.captured
  const semanticUiProofGuardrailPassed = scenarioProof.semanticUiProof.captured
  const tenXMeanAndP95AgainstGoogleSheets =
    timingMetricPassedAgainstGoogleSheets &&
    operationResponseProofGuardrailPassed &&
    authoritativeRenderProofGuardrailPassed &&
    postOperationFrameGuardrailPassed &&
    visualProofGuardrailPassed &&
    semanticUiProofGuardrailPassed &&
    biligRuntimeProofGuardrailPassed &&
    sourceWorkbookFingerprintGuardrailPassed
  const tenXMeanAndP95AgainstMicrosoftExcelWeb =
    timingMetricPassedAgainstMicrosoftExcelWeb === undefined
      ? undefined
      : timingMetricPassedAgainstMicrosoftExcelWeb &&
        operationResponseProofGuardrailPassed &&
        authoritativeRenderProofGuardrailPassed &&
        postOperationFrameGuardrailPassed &&
        visualProofGuardrailPassed &&
        semanticUiProofGuardrailPassed &&
        biligRuntimeProofGuardrailPassed &&
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
    ...sameCorpusScenarioCaseFields(scenarioProof),
    scenarioProof,
    postOperationFrameGuardrailPassed,
    operationResponseProofGuardrailPassed,
    authoritativeRenderProofGuardrailPassed,
    biligRuntimeProofGuardrailPassed,
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
    semanticUiProof: proof.semanticUiProof.products.find((semanticProof) => semanticProof.product === entry.product),
  }))
}

function buildSameCorpusMeasurement(capture: SameCorpusCaptureMeasurement): UiResponsivenessSameCorpusMeasurement {
  return {
    product: capture.product,
    source: capture.source,
    operationResponseMs: summarizeNumbers(capture.operationResponseMsSamples),
    operationResponseProofs: [...capture.operationResponseProofs],
    ...(capture.authoritativeRenderProofMsSamples
      ? { authoritativeRenderProofMs: summarizeNumbers(capture.authoritativeRenderProofMsSamples) }
      : {}),
    postOperationFrameMs: summarizeNumbers(capture.postOperationFrameMsSamples),
    ...(capture.scrollEventResponseMsSamples ? { scrollEventResponseMs: summarizeNumbers(capture.scrollEventResponseMsSamples) } : {}),
    ...(capture.scrollMovementPxSamples ? { scrollMovementPx: summarizeNumbers(capture.scrollMovementPxSamples) } : {}),
    ...(capture.biligRuntimeProof ? { biligRuntimeProof: cloneBiligRuntimeProof(capture.biligRuntimeProof) } : {}),
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
  const operationResponseProofGuardrailPassed = sameCorpusCaseOperationResponseProofGuardrailPassed(entry)
  const sourceWorkbookFingerprintGuardrailPassed = comparedProducts.every(hasCapturedSourceWorkbookFingerprint)
  const biligRuntimeProofGuardrailPassed = hasBiligProductionRuntimeProof(entry.bilig)
  const authoritativeRenderProofGuardrailPassed = hasBiligAuthoritativeRenderProofTiming(entry.bilig, entry.sampleCount)
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
    entry.operationResponseProofGuardrailPassed !== undefined &&
    entry.operationResponseProofGuardrailPassed !== operationResponseProofGuardrailPassed
  ) {
    throw new Error(`UI responsiveness same-corpus operation-response proof guardrail is stale: ${entry.id}`)
  }
  if (
    entry.sourceWorkbookFingerprintGuardrailPassed !== undefined &&
    entry.sourceWorkbookFingerprintGuardrailPassed !== sourceWorkbookFingerprintGuardrailPassed
  ) {
    throw new Error(`UI responsiveness same-corpus source workbook fingerprint guardrail is stale: ${entry.id}`)
  }
  if (entry.biligRuntimeProofGuardrailPassed !== undefined && entry.biligRuntimeProofGuardrailPassed !== biligRuntimeProofGuardrailPassed) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof guardrail is stale: ${entry.id}`)
  }
  if (
    entry.authoritativeRenderProofGuardrailPassed !== undefined &&
    entry.authoritativeRenderProofGuardrailPassed !== authoritativeRenderProofGuardrailPassed
  ) {
    throw new Error(`UI responsiveness same-corpus authoritative render proof guardrail is stale: ${entry.id}`)
  }
  if (entry.scrollMovementGuardrailPassed !== undefined && entry.scrollMovementGuardrailPassed !== scrollMovementGuardrailPassed) {
    throw new Error(`UI responsiveness same-corpus scroll-movement guardrail is stale: ${entry.id}`)
  }
  const requiresScrollEventMetric = uiSameCorpusWorkloadRequiresScrollEventEvidence(entry.workload)
  const expectedMetric = requiresScrollEventMetric ? 'scrollEventResponseMs' : 'operationResponseMs'
  if (entry.tenXMeanAndP95Metric !== expectedMetric) {
    throw new Error(`UI responsiveness same-corpus metric is stale: ${entry.id}`)
  }
  validateSameCorpusScenarioProof(entry.scenarioProof, entry.id, entry.bilig, entry.googleSheets, entry.workload, entry.microsoftExcelWeb)
  validateSameCorpusScenarioCaseFields(entry, 'scorecard')
  const visualProofGuardrailPassed = entry.scenarioProof.screenshotProof.captured && entry.scenarioProof.pixelGridProof.captured
  const semanticUiProofGuardrailPassed = entry.scenarioProof.semanticUiProof.captured
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
    operationResponseProofGuardrailPassed &&
    authoritativeRenderProofGuardrailPassed &&
    postOperationFrameGuardrailPassed &&
    visualProofGuardrailPassed &&
    semanticUiProofGuardrailPassed &&
    biligRuntimeProofGuardrailPassed &&
    sourceWorkbookFingerprintGuardrailPassed
  const tenXAgainstMicrosoftExcelWeb =
    timingMetricPassedAgainstMicrosoftExcelWeb === undefined
      ? undefined
      : timingMetricPassedAgainstMicrosoftExcelWeb &&
        operationResponseProofGuardrailPassed &&
        authoritativeRenderProofGuardrailPassed &&
        postOperationFrameGuardrailPassed &&
        visualProofGuardrailPassed &&
        semanticUiProofGuardrailPassed &&
        biligRuntimeProofGuardrailPassed &&
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

function cloneBiligRuntimeProof(proof: SameCorpusBiligRuntimeProof): SameCorpusBiligRuntimeProof {
  return {
    ...proof,
    samples: proof.samples.map((sample) => ({ ...sample })),
  }
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
  if (product === 'bilig' && measurement.biligRuntimeProof) {
    validateBiligRuntimeProof(measurement.biligRuntimeProof, measurement.source, caseId)
  }
  validateSummary(measurement.operationResponseMs, `${caseId} ${product} operationResponseMs`, sameCorpusSampleCount)
  if (measurement.operationResponseProofs.length < sameCorpusSampleCount) {
    throw new Error(`UI responsiveness same-corpus proof is missing operation-response proof for ${caseId}`)
  }
  if (measurement.authoritativeRenderProofMs) {
    validateSummary(measurement.authoritativeRenderProofMs, `${caseId} ${product} authoritativeRenderProofMs`, sameCorpusSampleCount)
  }
  validateSummary(measurement.postOperationFrameMs, `${caseId} ${product} postOperationFrameMs`, sameCorpusSampleCount)
  if (measurement.scrollEventResponseMs) {
    validateSummary(measurement.scrollEventResponseMs, `${caseId} ${product} scrollEventResponseMs`, sameCorpusSampleCount)
  }
  if (measurement.scrollMovementPx) {
    validateSummary(measurement.scrollMovementPx, `${caseId} ${product} scrollMovementPx`, sameCorpusSampleCount)
  }
  validateSameCorpusCaptureVerification(measurement.corpusVerification, product, materializedCells, corpusCaseId, caseId)
}

function validateBiligRuntimeProof(proof: SameCorpusBiligRuntimeProof, source: string, caseId: string): void {
  if (proof.product !== 'bilig' || proof.source !== source || proof.verificationMethod !== 'window.__biligRuntimeBuild') {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof identity mismatch for ${caseId}`)
  }
  if (proof.requiredBuildKind !== 'production') {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof has stale required build kind for ${caseId}`)
  }
  if (!['development', 'production', 'unknown'].includes(proof.actualBuildKind)) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof has invalid build kind for ${caseId}`)
  }
  if (proof.sampleCount !== proof.samples.length || proof.sampleCount <= 0) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample count is stale for ${caseId}`)
  }
  for (const sample of proof.samples) {
    if (!Number.isInteger(sample.sampleIndex) || sample.sampleIndex < 0) {
      throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample index is invalid for ${caseId}`)
    }
    if (!['development', 'production', 'unknown'].includes(sample.buildKind)) {
      throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample build kind is invalid for ${caseId}`)
    }
  }
  const verified = hasBiligProductionRuntimeProof({ product: 'bilig', biligRuntimeProof: proof })
  if (proof.verified !== verified) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof verified flag is stale for ${caseId}`)
  }
}
