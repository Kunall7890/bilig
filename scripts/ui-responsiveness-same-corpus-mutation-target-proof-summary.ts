import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import {
  validateSameCorpusMutationTargetProofSample,
  validateSameCorpusProductSemanticUiProof,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { SameCorpusMutationTargetProof, SameCorpusProductSemanticUiProof } from './ui-responsiveness-same-corpus-semantic-proof.ts'
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
        return productTotal + sameCorpusMutationTargetProductProofSummary(entry, product, workload, expectedSampleCount).acceptedSampleCount
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
  const baseVerdict = validateSameCorpusProductSemanticUiProof(productProof)
  const sampleSummaries = sameCorpusMutationTargetProofSampleSummaries(
    productProof,
    workload,
    requiredSampleCount,
    baseVerdict.invalidReasons,
  )
  const acceptedSampleCount = sampleSummaries.filter((sample) => sample.accepted).length
  const accepted = acceptedSampleCount === requiredSampleCount && verdict.acceptedForCurrentScorecard
  return {
    workload,
    product,
    requiredSampleCount,
    rawSampleCount: productProof.mutationTargetProofs.length,
    acceptedSampleCount,
    accepted,
    samples: sampleSummaries,
    invalidReasons: verdict.invalidReasons,
  }
}

function sameCorpusMutationTargetProofSampleSummaries(
  productProof: SameCorpusProductSemanticUiProof,
  workload: UiResponsivenessSameCorpusWorkload,
  requiredSampleCount: number,
  baseInvalidReasons: readonly string[],
): SameCorpusMutationTargetProofProductSummary['samples'] {
  const duplicateSampleIndexes = duplicateSameCorpusSampleIndexes(productProof.mutationTargetProofs)
  const duplicateScreenshotPaths = duplicateSameCorpusScreenshotPaths(productProof.mutationTargetProofs)
  return Array.from({ length: requiredSampleCount }, (_, sampleIndex) => {
    const sample = productProof.mutationTargetProofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!sample) {
      return {
        sampleIndex,
        present: false,
        accepted: false,
        committedTargetProofMs: null,
        sheetName: null,
        sheetId: null,
        targetRange: null,
        intendedOperation: null,
        intendedPayload: null,
        before: null,
        after: null,
        restored: null,
        visibleAfter: null,
        visibleRestored: null,
        visibleAfterSelectedRange: null,
        visibleRestoredSelectedRange: null,
        authoritativeReadbackRevision: null,
        visibleRenderRevision: null,
        targetScreenshots: null,
        screenshotPath: null,
        screenshotSha256: null,
        undoRestoreStatus: null,
        invalidReasons: [`semantic UI mutation target proof for ${workload} is missing sample ${String(sampleIndex + 1)}`],
      }
    }
    const sampleVerdict = validateSameCorpusMutationTargetProofSample(productProof, workload, sample, {
      duplicateSampleIndex: duplicateSampleIndexes.has(sample.sampleIndex),
      duplicateScreenshotPath:
        sample.screenshotPath !== null && sample.screenshotPath !== undefined
          ? duplicateScreenshotPaths.has(normalizeSameCorpusMutationTargetScreenshotPath(sample.screenshotPath))
          : false,
      sampleCount: requiredSampleCount,
    })
    const invalidReasons = [...baseInvalidReasons, ...sampleVerdict.invalidReasons]
    return {
      sampleIndex,
      present: true,
      accepted: invalidReasons.length === 0,
      committedTargetProofMs: sample.committedTargetProofMs,
      sheetName: sample.sheetName,
      sheetId: sample.sheetId,
      targetRange: sample.targetRange,
      intendedOperation: sample.intendedOperation,
      intendedPayload: sample.intendedPayload,
      before: sample.before,
      after: sample.after,
      restored: sample.restored,
      visibleAfter: sample.visibleAfter,
      visibleRestored: sample.visibleRestored,
      visibleAfterSelectedRange: sample.visibleAfterSelectedRange,
      visibleRestoredSelectedRange: sample.visibleRestoredSelectedRange,
      authoritativeReadbackRevision: sample.authoritativeReadbackRevision,
      visibleRenderRevision: sample.visibleRenderRevision,
      targetScreenshots: sample.targetScreenshots,
      screenshotPath: sample.screenshotPath,
      screenshotSha256: sample.screenshotSha256,
      undoRestoreStatus: sample.undoRestoreStatus,
      invalidReasons,
    }
  })
}

function duplicateSameCorpusSampleIndexes(samples: readonly SameCorpusMutationTargetProof[]): ReadonlySet<number> {
  const counts = new Map<number, number>()
  for (const sample of samples) {
    counts.set(sample.sampleIndex, (counts.get(sample.sampleIndex) ?? 0) + 1)
  }
  return new Set([...counts].filter((entry) => entry[1] > 1).map((entry) => entry[0]))
}

function duplicateSameCorpusScreenshotPaths(samples: readonly SameCorpusMutationTargetProof[]): ReadonlySet<string> {
  const counts = new Map<string, number>()
  for (const sample of samples) {
    if (sample.screenshotPath === null || sample.screenshotPath === undefined || sample.screenshotPath.trim().length === 0) {
      continue
    }
    const path = normalizeSameCorpusMutationTargetScreenshotPath(sample.screenshotPath)
    counts.set(path, (counts.get(path) ?? 0) + 1)
  }
  return new Set([...counts].filter((entry) => entry[1] > 1).map((entry) => entry[0]))
}

function normalizeSameCorpusMutationTargetScreenshotPath(path: string): string {
  return path.trim().replaceAll('\\', '/')
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
    committedTargetProofMs: null,
    sheetName: null,
    sheetId: null,
    targetRange: null,
    intendedOperation: null,
    intendedPayload: null,
    before: null,
    after: null,
    restored: null,
    visibleAfter: null,
    visibleRestored: null,
    visibleAfterSelectedRange: null,
    visibleRestoredSelectedRange: null,
    authoritativeReadbackRevision: null,
    visibleRenderRevision: null,
    targetScreenshots: null,
    screenshotPath: null,
    screenshotSha256: null,
    undoRestoreStatus: null,
    invalidReasons: ['semantic UI mutation target proof sample is missing'],
  }))
}
