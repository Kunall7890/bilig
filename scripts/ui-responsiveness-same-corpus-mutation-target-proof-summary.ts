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

export function requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount(sampleCount: number): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount() * 2 * Math.max(0, sampleCount)
}

export function sameCorpusMutationTargetProofCaseCount(cases: readonly SameCorpusMutationTargetProofSummaryCase[]): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.filter((workload) =>
    cases.some((entry) => entry.workload === workload && entry.scenarioProof.semanticUiProof.captured),
  ).length
}

export function sameCorpusMutationTargetProofSampleCount(cases: readonly SameCorpusMutationTargetProofSummaryCase[]): number {
  const requiredProducts = ['bilig', 'google-sheets'] as const
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.reduce((total, workload) => {
    const entry = cases.find((candidate) => candidate.workload === workload)
    if (!entry?.scenarioProof.semanticUiProof.captured) {
      return total
    }
    return (
      total +
      requiredProducts.reduce((productTotal, product) => {
        const productProof = entry.scenarioProof.semanticUiProof.products.find((proof) => proof.product === product)
        return productTotal + (productProof?.mutationTargetProofs.length ?? 0)
      }, 0)
    )
  }, 0)
}
