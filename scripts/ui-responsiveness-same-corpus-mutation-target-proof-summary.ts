import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadMutatesWorkbook,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

interface SameCorpusMutationTargetProofSummaryCase {
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly scenarioProof: SameCorpusScenarioProof
}

export const requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads =
  requiredUiResponsivenessSameCorpusWorkloads.filter(uiSameCorpusWorkloadMutatesWorkbook)

export function requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount(): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.length
}

export function sameCorpusMutationTargetProofCaseCount(cases: readonly SameCorpusMutationTargetProofSummaryCase[]): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.filter((workload) =>
    cases.some((entry) => entry.workload === workload && entry.scenarioProof.semanticUiProof.captured),
  ).length
}
