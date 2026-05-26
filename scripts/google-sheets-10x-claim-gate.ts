#!/usr/bin/env bun

import { pathToFileURL } from 'node:url'

import { loadBiligDominanceScorecardInput } from './bilig-dominance-scorecard-input.ts'
import { buildBiligDominanceScorecard } from './gen-bilig-dominance-scorecard.ts'
import {
  requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount,
  requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount,
} from './ui-responsiveness-same-corpus-mutation-target-proof-summary.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from './ui-responsiveness-same-corpus-workloads.ts'

export interface GoogleSheetsTenXClaimGateReport {
  readonly passed: boolean
  readonly requiredWorkloadCount: number
  readonly tenXMeanAndP95CaseCount: number | null
  readonly failures: readonly string[]
}

export interface GoogleSheetsTenXClaimGateInput {
  readonly scorecard: unknown
  readonly uiScorecard: unknown
  readonly browserGatesSkipped?: boolean
}

export function buildGoogleSheetsTenXClaimGateReport(input: GoogleSheetsTenXClaimGateInput): GoogleSheetsTenXClaimGateReport {
  const requiredWorkloadCount = requiredUiResponsivenessSameCorpusWorkloads.length
  const failures = googleSheetsTenXClaimGateFailures(input, requiredWorkloadCount)
  return {
    passed: failures.length === 0,
    requiredWorkloadCount,
    tenXMeanAndP95CaseCount: sameCorpusTenXMeanAndP95CaseCount(input.uiScorecard),
    failures,
  }
}

export function scorecardPermitsGoogleSheetsTenXClaim(scorecard: unknown): boolean {
  return scorecardLevelFailures(scorecard).length === 0
}

function googleSheetsTenXClaimGateFailures(input: GoogleSheetsTenXClaimGateInput, requiredWorkloadCount: number): string[] {
  return [
    ...(input.browserGatesSkipped
      ? ['Google Sheets 10x claim gate cannot run with BILIG_CI_SKIP_BROWSER=1; live browser proof must be enabled']
      : []),
    ...scorecardLevelFailures(input.scorecard),
    ...sameCorpusUiProofFailures(input.uiScorecard, requiredWorkloadCount),
  ]
}

function scorecardLevelFailures(scorecard: unknown): string[] {
  if (!isRecord(scorecard)) {
    return ['dominance scorecard must be an object']
  }
  const claimPolicy = asRecord(scorecard['claimPolicy'])
  const completionAudit = asRecord(scorecard['completionAudit'])
  const googleSheetsStatus = asRecord(scorecard['overallGoogleSheets10xStatus'])
  const googleSheetsCategories = Array.isArray(googleSheetsStatus?.['categories']) ? googleSheetsStatus['categories'] : []
  const failures: string[] = []
  if (scorecard['goalStatus'] !== 'achieved') {
    failures.push(`dominance scorecard goalStatus is ${String(scorecard['goalStatus'])}, expected achieved`)
  }
  if (claimPolicy?.['blanketTenXClaimAllowed'] !== true) {
    failures.push('dominance scorecard claimPolicy.blanketTenXClaimAllowed is not true')
  }
  if (Array.isArray(claimPolicy?.['unmetRequirements']) && claimPolicy['unmetRequirements'].length > 0) {
    failures.push(`dominance scorecard claimPolicy still has unmet requirements: ${claimPolicy['unmetRequirements'].join('; ')}`)
  }
  if (completionAudit?.['allCriteriaPassed'] !== true) {
    failures.push('dominance scorecard completionAudit.allCriteriaPassed is not true')
  }
  if (Array.isArray(completionAudit?.['unmetRequirements']) && completionAudit['unmetRequirements'].length > 0) {
    failures.push(`dominance scorecard completionAudit still has unmet requirements: ${completionAudit['unmetRequirements'].join('; ')}`)
  }
  if (googleSheetsStatus?.['passed'] !== true || googleSheetsStatus?.['status'] !== 'passed') {
    failures.push('dominance scorecard overallGoogleSheets10xStatus is not passed')
  }
  if (Array.isArray(googleSheetsStatus?.['unmetRequirements']) && googleSheetsStatus['unmetRequirements'].length > 0) {
    failures.push(`Google Sheets 10x status still has unmet requirements: ${googleSheetsStatus['unmetRequirements'].join('; ')}`)
  }
  for (const category of googleSheetsCategories) {
    if (!isRecord(category)) {
      failures.push('Google Sheets 10x status has a malformed category')
      continue
    }
    if (category['passed'] !== true) {
      failures.push(`Google Sheets 10x category ${String(category['id'])} is not passed`)
    }
    if (Array.isArray(category['gaps']) && category['gaps'].length > 0) {
      failures.push(`Google Sheets 10x category ${String(category['id'])} still has gaps: ${category['gaps'].join('; ')}`)
    }
  }
  return failures
}

function sameCorpusUiProofFailures(uiScorecard: unknown, requiredWorkloadCount: number): string[] {
  if (!isRecord(uiScorecard)) {
    return ['UI responsiveness scorecard must be an object']
  }
  const proof = asRecord(uiScorecard['sameCorpusProof'])
  if (!proof) {
    return ['UI responsiveness scorecard is missing sameCorpusProof']
  }
  const runManifest = asRecord(proof['runManifest'])
  const cases = Array.isArray(proof['cases']) ? proof['cases'] : []
  const failures: string[] = []
  if (proof['captured'] !== true) {
    failures.push('same-corpus UI proof is not captured')
  }
  if (proof['requiredCaseCount'] !== requiredWorkloadCount) {
    failures.push(
      `same-corpus UI proof requiredCaseCount is ${String(proof['requiredCaseCount'])}, expected ${String(requiredWorkloadCount)}`,
    )
  }
  if (cases.length !== requiredWorkloadCount) {
    failures.push(`same-corpus UI proof has ${String(cases.length)}/${String(requiredWorkloadCount)} required cases`)
  }
  if (proof['tenXMeanAndP95CaseCount'] !== requiredWorkloadCount) {
    failures.push(
      `same-corpus UI proof has ${String(proof['tenXMeanAndP95CaseCount'])}/${String(requiredWorkloadCount)} 10x mean+p95 cases`,
    )
  }
  if (!runManifest) {
    failures.push('same-corpus UI proof is missing runManifest')
    return failures
  }
  if (runManifest['currentContractEvidenceComplete'] !== true) {
    failures.push('same-corpus UI runManifest.currentContractEvidenceComplete is not true')
  }
  if (runManifest['googleSheetsTenXRequirementSatisfied'] !== true) {
    failures.push('same-corpus UI runManifest.googleSheetsTenXRequirementSatisfied is not true')
  }
  if (Array.isArray(runManifest['invalidReasons']) && runManifest['invalidReasons'].length > 0) {
    failures.push(`same-corpus UI runManifest still has invalid reasons: ${runManifest['invalidReasons'].join('; ')}`)
  }
  for (const [field, label] of [
    ['scenarioSummaryFieldCaseCount', 'first-class scenario summary fields'],
    ['strictRenderedGridProofCaseCount', 'strict rendered-grid proof'],
    ['visibleOperationResponseProofCaseCount', 'visible operation-response proof'],
    ['biligAuthoritativeRenderProofCaseCount', 'Bilig authoritative render proof timing'],
    ['semanticUiProofCaseCount', 'semantic UI proof'],
    ['tenXMeanAndP95CaseCount', '10x mean+p95 cases'],
  ] as const) {
    if (runManifest[field] !== requiredWorkloadCount) {
      failures.push(`same-corpus UI ${label} covers ${String(runManifest[field])}/${String(requiredWorkloadCount)} cases`)
    }
  }
  const requiredMutationTargetProofCaseCount =
    typeof runManifest['requiredMutationTargetProofCaseCount'] === 'number'
      ? runManifest['requiredMutationTargetProofCaseCount']
      : requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount()
  if (runManifest['mutationTargetProofCaseCount'] !== requiredMutationTargetProofCaseCount) {
    failures.push(
      `same-corpus UI mutation target proof covers ${String(runManifest['mutationTargetProofCaseCount'])}/${String(
        requiredMutationTargetProofCaseCount,
      )} mutating cases`,
    )
  }
  const requiredMutationTargetProofSampleCount =
    typeof runManifest['requiredMutationTargetProofSampleCount'] === 'number'
      ? runManifest['requiredMutationTargetProofSampleCount']
      : requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount(Number(runManifest['sampleCount'] ?? 0))
  if (runManifest['mutationTargetProofSampleCount'] !== requiredMutationTargetProofSampleCount) {
    failures.push(
      `same-corpus UI mutation target proof covers ${String(runManifest['mutationTargetProofSampleCount'])}/${String(
        requiredMutationTargetProofSampleCount,
      )} required per-sample product proofs`,
    )
  }
  if (runManifest['legacyInsufficientRenderedGridProofCaseCount'] !== 0) {
    failures.push(
      `same-corpus UI legacy-insufficient rendered-grid proof covers ${String(
        runManifest['legacyInsufficientRenderedGridProofCaseCount'],
      )}/${String(requiredWorkloadCount)} cases`,
    )
  }
  const workloads = new Set(cases.flatMap((entry) => (isRecord(entry) && typeof entry['workload'] === 'string' ? [entry['workload']] : [])))
  const missingWorkloads = requiredUiResponsivenessSameCorpusWorkloads.filter((workload) => !workloads.has(workload))
  if (missingWorkloads.length > 0) {
    failures.push(`same-corpus UI proof is missing workloads: ${missingWorkloads.join(', ')}`)
  }
  const failedCases = cases.flatMap((entry) => {
    if (!isRecord(entry)) {
      return ['<malformed>']
    }
    return entry['passed'] === true && entry['tenXMeanAndP95AgainstGoogleSheets'] === true ? [] : [String(entry['id'] ?? entry['workload'])]
  })
  if (failedCases.length > 0) {
    failures.push(`same-corpus UI cases are not all passed 10x cases: ${failedCases.join(', ')}`)
  }
  return failures
}

function sameCorpusTenXMeanAndP95CaseCount(uiScorecard: unknown): number | null {
  const proof = isRecord(uiScorecard) ? asRecord(uiScorecard['sameCorpusProof']) : null
  const value = proof?.['tenXMeanAndP95CaseCount']
  return typeof value === 'number' ? value : null
}

function main(): void {
  const input = loadBiligDominanceScorecardInput()
  const report = buildGoogleSheetsTenXClaimGateReport({
    scorecard: buildBiligDominanceScorecard(input),
    uiScorecard: input.uiResponsivenessLiveBrowserScorecard,
    browserGatesSkipped: process.env['BILIG_CI_SKIP_BROWSER'] === '1',
  })
  if (!report.passed) {
    throw new Error(`Google Sheets 10x release-claim gate failed:\n${report.failures.map((failure) => `- ${failure}`).join('\n')}`)
  }
  console.log(JSON.stringify(report, null, 2))
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return isRecord(value) ? value : null
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  main()
}
