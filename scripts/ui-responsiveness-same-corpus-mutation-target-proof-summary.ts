import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import { validateSameCorpusProductSemanticUiProof } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { SameCorpusMutationTargetProofProductSummary } from './ui-responsiveness-same-corpus-scorecard-types.ts'
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

export function sameCorpusMutationTargetProofProductSummaries(
  cases: readonly SameCorpusMutationTargetProofSummaryCase[],
  expectedSampleCount?: number,
): readonly SameCorpusMutationTargetProofProductSummary[] {
  return requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.flatMap((workload) => {
    const entry = cases.find((candidate) => candidate.workload === workload)
    return requiredMutationTargetProofProducts.map((product) =>
      sameCorpusMutationTargetProductProofSummary(entry, product, workload, expectedSampleCount),
    )
  })
}

function sameCorpusMutationTargetProofCaseComplete(
  entry: SameCorpusMutationTargetProofSummaryCase,
  expectedSampleCount: number | undefined,
): boolean {
  return requiredMutationTargetProofProducts.every((product) =>
    sameCorpusMutationTargetProductProofComplete(entry, product, expectedSampleCount),
  )
}

function sameCorpusMutationTargetProductProofSummary(
  entry: SameCorpusMutationTargetProofSummaryCase | undefined,
  product: (typeof requiredMutationTargetProofProducts)[number],
  workload: UiResponsivenessSameCorpusWorkload,
  expectedSampleCount: number | undefined,
): SameCorpusMutationTargetProofProductSummary {
  const productProof = entry?.scenarioProof.semanticUiProof.products.find((proof) => proof.product === product)
  const requiredSampleCount = expectedSampleCount ?? entry?.sampleCount ?? productProof?.mutationTargetProofs.length ?? 0
  if (!entry || !productProof) {
    return {
      workload,
      product,
      requiredSampleCount,
      rawSampleCount: 0,
      acceptedSampleCount: 0,
      accepted: false,
      samples: sameCorpusMissingMutationTargetSamples(requiredSampleCount),
      invalidReasons: [`semantic UI mutation target proof for ${workload} is missing ${product} proof`],
    }
  }
  const verdict = validateSameCorpusProductSemanticUiProof(productProof, {
    sampleCount: requiredSampleCount,
    workload,
  })
  const accepted = verdict.acceptedForCurrentScorecard
  return {
    workload,
    product,
    requiredSampleCount,
    rawSampleCount: productProof.mutationTargetProofs.length,
    acceptedSampleCount: accepted ? productProof.mutationTargetProofs.length : 0,
    accepted,
    samples: Array.from({ length: requiredSampleCount }, (_, sampleIndex) => {
      const sample = productProof.mutationTargetProofs.find((candidate) => candidate.sampleIndex === sampleIndex)
      return {
        sampleIndex,
        present: sample !== undefined,
        accepted: accepted && sample !== undefined,
        targetRange: sample?.targetRange ?? null,
        screenshotPath: sample?.screenshotPath ?? null,
        screenshotSha256: sample?.screenshotSha256 ?? null,
      }
    }),
    invalidReasons: verdict.invalidReasons,
  }
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

function sameCorpusMissingMutationTargetSamples(requiredSampleCount: number): SameCorpusMutationTargetProofProductSummary['samples'] {
  return Array.from({ length: Math.max(0, requiredSampleCount) }, (_, sampleIndex) => ({
    sampleIndex,
    present: false,
    accepted: false,
    targetRange: null,
    screenshotPath: null,
    screenshotSha256: null,
  }))
}
