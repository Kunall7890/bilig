import type { SameCorpusCaptureCorpusFingerprint } from './ui-responsiveness-same-corpus-fingerprint.ts'
import type { SameCorpusProductSourceWorkbookFingerprint } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

export function sameCorpusManifestInvalidReasons(args: {
  readonly capturedWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
  readonly caseCount: number
  readonly corpusCaseIds: readonly string[]
  readonly corpusFingerprints: readonly SameCorpusCaptureCorpusFingerprint[]
  readonly productSourceWorkbookFingerprints: readonly SameCorpusProductSourceWorkbookFingerprint[]
  readonly biligProductionRuntimeProofCaseCount: number
  readonly scenarioSummaryFieldCaseCount: number
  readonly biligAuthoritativeRenderProofCaseCount: number
  readonly semanticUiProofCaseCount: number
  readonly requiredMutationTargetProofCaseCount: number
  readonly mutationTargetProofCaseCount: number
  readonly requiredMutationTargetProofSampleCount: number
  readonly mutationTargetProofSampleCount: number
  readonly requiredCommittedTargetProofTimingCaseCount: number
  readonly committedTargetProofTimingCaseCount: number
  readonly requiredCommittedTargetProofTimingSampleCount: number
  readonly committedTargetProofTimingSampleCount: number
  readonly legacyInsufficientRenderedGridProofCaseCount: number
  readonly materializedCellCounts: readonly number[]
  readonly strictRenderedGridProofCaseCount: number
  readonly visibleOperationResponseProofCaseCount: number
  readonly tenXMeanAndP95CaseCount: number
}): string[] {
  const invalidReasons: string[] = []
  const requiredWorkloads = requiredUiResponsivenessSameCorpusWorkloads
  if (args.caseCount !== requiredWorkloads.length) {
    invalidReasons.push('required workload count is incomplete')
  }
  const missingWorkloads = requiredWorkloads.filter((workload) => !args.capturedWorkloads.includes(workload))
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
  if (args.biligProductionRuntimeProofCaseCount !== requiredWorkloads.length) {
    invalidReasons.push(
      `Bilig production runtime proof covers ${String(args.biligProductionRuntimeProofCaseCount)}/${String(requiredWorkloads.length)} cases`,
    )
  }
  if (args.scenarioSummaryFieldCaseCount !== requiredWorkloads.length) {
    invalidReasons.push(
      `first-class scenario summary fields cover ${String(args.scenarioSummaryFieldCaseCount)}/${String(requiredWorkloads.length)} cases`,
    )
  }
  if (args.strictRenderedGridProofCaseCount !== requiredWorkloads.length) {
    invalidReasons.push(
      `strict rendered-grid proof covers ${String(args.strictRenderedGridProofCaseCount)}/${String(requiredWorkloads.length)} cases`,
    )
  }
  if (args.visibleOperationResponseProofCaseCount !== requiredWorkloads.length) {
    invalidReasons.push(
      `visible operation-response proof covers ${String(args.visibleOperationResponseProofCaseCount)}/${String(requiredWorkloads.length)} cases`,
    )
  }
  if (args.biligAuthoritativeRenderProofCaseCount !== requiredWorkloads.length) {
    invalidReasons.push(
      `Bilig authoritative render proof timing covers ${String(args.biligAuthoritativeRenderProofCaseCount)}/${String(requiredWorkloads.length)} cases`,
    )
  }
  if (args.semanticUiProofCaseCount !== requiredWorkloads.length) {
    invalidReasons.push(`semantic UI proof covers ${String(args.semanticUiProofCaseCount)}/${String(requiredWorkloads.length)} cases`)
  }
  if (args.mutationTargetProofCaseCount !== args.requiredMutationTargetProofCaseCount) {
    invalidReasons.push(
      `mutation target proof covers ${String(args.mutationTargetProofCaseCount)}/${String(
        args.requiredMutationTargetProofCaseCount,
      )} mutating cases`,
    )
  }
  if (args.mutationTargetProofSampleCount !== args.requiredMutationTargetProofSampleCount) {
    invalidReasons.push(
      `mutation target proof covers ${String(args.mutationTargetProofSampleCount)}/${String(
        args.requiredMutationTargetProofSampleCount,
      )} required per-sample product proofs`,
    )
  }
  if (args.committedTargetProofTimingCaseCount !== args.requiredCommittedTargetProofTimingCaseCount) {
    invalidReasons.push(
      `committed target proof timing covers ${String(args.committedTargetProofTimingCaseCount)}/${String(
        args.requiredCommittedTargetProofTimingCaseCount,
      )} mutating cases`,
    )
  }
  if (args.committedTargetProofTimingSampleCount !== args.requiredCommittedTargetProofTimingSampleCount) {
    invalidReasons.push(
      `committed target proof timing covers ${String(args.committedTargetProofTimingSampleCount)}/${String(
        args.requiredCommittedTargetProofTimingSampleCount,
      )} required per-sample product timings`,
    )
  }
  if (args.legacyInsufficientRenderedGridProofCaseCount > 0) {
    invalidReasons.push(
      `legacy-insufficient rendered-grid proof covers ${String(args.legacyInsufficientRenderedGridProofCaseCount)}/${String(requiredWorkloads.length)} cases`,
    )
  }
  if (args.tenXMeanAndP95CaseCount !== requiredWorkloads.length) {
    invalidReasons.push('not every required workload is 10x against Google Sheets')
  }
  return invalidReasons
}

function requiredProductSourceWorkbookFingerprintsPresent(fingerprints: readonly SameCorpusProductSourceWorkbookFingerprint[]): boolean {
  const requiredProducts = ['bilig', 'google-sheets'] as const
  return requiredProducts.every(
    (product) => fingerprints.filter((entry) => entry.product === product && entry.sourceWorkbookSha256 !== null).length === 1,
  )
}
