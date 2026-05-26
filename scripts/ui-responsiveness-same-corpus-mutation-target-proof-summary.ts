import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import { validateSameCorpusProductSemanticUiProof } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadMutatesWorkbook,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

interface SameCorpusMutationTargetProofSummaryCase {
  readonly sampleCount?: number
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly scenarioProof: SameCorpusScenarioProof
}

const requiredMutationTargetProofProducts = ['bilig', 'google-sheets'] as const

export const requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads =
  requiredUiResponsivenessSameCorpusWorkloads.filter(uiSameCorpusWorkloadMutatesWorkbook)

export function requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount(): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.length
}

export function requiredUiResponsivenessSameCorpusMutationTargetProofSampleCount(sampleCount: number): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofCaseCount() * 2 * Math.max(0, sampleCount)
}

export function sameCorpusMutationTargetProofCaseCount(
  cases: readonly SameCorpusMutationTargetProofSummaryCase[],
  expectedSampleCount?: number,
): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.filter((workload) =>
    cases.some((entry) => entry.workload === workload && sameCorpusMutationTargetProofCaseComplete(entry, expectedSampleCount)),
  ).length
}

export function sameCorpusMutationTargetProofSampleCount(
  cases: readonly SameCorpusMutationTargetProofSummaryCase[],
  expectedSampleCount?: number,
): number {
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.reduce((total, workload) => {
    const entry = cases.find((candidate) => candidate.workload === workload)
    if (!entry) {
      return total
    }
    return (
      total +
      requiredMutationTargetProofProducts.reduce((productTotal, product) => {
        const productProof = entry.scenarioProof.semanticUiProof.products.find((proof) => proof.product === product)
        return (
          productTotal +
          (sameCorpusMutationTargetProductProofComplete(entry, product, expectedSampleCount)
            ? (productProof?.mutationTargetProofs.length ?? 0)
            : 0)
        )
      }, 0)
    )
  }, 0)
}

function sameCorpusMutationTargetProofCaseComplete(
  entry: SameCorpusMutationTargetProofSummaryCase,
  expectedSampleCount: number | undefined,
): boolean {
  return requiredMutationTargetProofProducts.every((product) =>
    sameCorpusMutationTargetProductProofComplete(entry, product, expectedSampleCount),
  )
}

function sameCorpusMutationTargetProductProofComplete(
  entry: SameCorpusMutationTargetProofSummaryCase,
  product: (typeof requiredMutationTargetProofProducts)[number],
  expectedSampleCount: number | undefined,
): boolean {
  const productProof = entry.scenarioProof.semanticUiProof.products.find((proof) => proof.product === product)
  if (!productProof) {
    return false
  }
  return validateSameCorpusProductSemanticUiProof(productProof, {
    sampleCount: expectedSampleCount ?? entry.sampleCount ?? productProof.mutationTargetProofs.length,
    workload: entry.workload,
  }).acceptedForCurrentScorecard
}
