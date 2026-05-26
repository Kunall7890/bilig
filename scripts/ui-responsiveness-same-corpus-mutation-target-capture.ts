import { mkdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import {
  captureSameCorpusMutationTargetScreenshotProof,
  readSameCorpusMutationTargetReadback,
  readSameCorpusMutationTargetRevisionProof,
  readSameCorpusVisibleMutationTargetReadback,
  selectSameCorpusMutationTargetRange,
  type SameCorpusMutationTargetSelection,
} from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type {
  SameCorpusMutationTargetProof,
  SameCorpusMutationTargetReadback,
  SameCorpusMutationTargetScreenshotProof,
} from './ui-responsiveness-same-corpus-proof.ts'
import { readSameCorpusVisibleSelectedRange } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import {
  restoreProductWorkbookMutation,
  sameCorpusFillColorExpectedColor,
  sameCorpusFillColorSwatchLabel,
} from './ui-responsiveness-same-corpus-workload-runner.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export async function captureSameCorpusMutationTargetProofForSample(args: {
  readonly before: SameCorpusMutationTargetReadback
  readonly beforeScreenshot: SameCorpusMutationTargetScreenshotProof
  readonly caseId?: string
  readonly operationStartedAt: number
  readonly outputPath: string
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetProof> {
  await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target })
  const after = await readSameCorpusMutationTargetReadback({ page: args.page, product: args.product, target: args.target })
  const visibleAfter = await readSameCorpusVisibleMutationTargetReadback({
    page: args.page,
    product: args.product,
    target: args.target,
    workload: args.workload,
  })
  const visibleAfterSelectedRange = await readSameCorpusVisibleSelectedRange(args.page, args.product)
  const afterScreenshot = await captureTargetScreenshot(args, 'after')
  const revisions = await readSameCorpusMutationTargetRevisionProof({
    page: args.page,
    product: args.product,
    readback: after,
    screenshotSha256: afterScreenshot.screenshotSha256,
    target: args.target,
  })
  const committedTargetProofMs = Math.max(0, performance.now() - args.operationStartedAt)
  await restoreProductWorkbookMutation(args.page, args.workload)
  await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target })
  const restored = await readSameCorpusMutationTargetReadback({ page: args.page, product: args.product, target: args.target })
  const visibleRestored = await readSameCorpusVisibleMutationTargetReadback({
    page: args.page,
    product: args.product,
    target: args.target,
    workload: args.workload,
  })
  const visibleRestoredSelectedRange = await readSameCorpusVisibleSelectedRange(args.page, args.product)
  const restoredScreenshot = await captureTargetScreenshot(args, 'restored')
  return {
    after,
    authoritativeReadbackRevision: revisions.authoritativeReadbackRevision,
    before: args.before,
    committedTargetProofMs,
    intendedOperation: args.workload,
    intendedPayload: intendedMutationTargetPayload(args.product, args.workload, args.sampleIndex),
    restored,
    sampleIndex: args.sampleIndex,
    screenshotPath: afterScreenshot.screenshotPath,
    screenshotSha256: afterScreenshot.screenshotSha256,
    sheetId: args.target.sheetId,
    sheetName: args.target.sheetName,
    targetRange: args.target.targetRange,
    targetScreenshots: {
      after: afterScreenshot,
      before: args.beforeScreenshot,
      restored: restoredScreenshot,
    },
    undoRestoreStatus: sameCorpusMutationReadbacksEqual(args.before, restored) ? 'verified' : 'failed',
    visibleAfter,
    visibleAfterSelectedRange,
    visibleRenderRevision: revisions.visibleRenderRevision,
    visibleRestored,
    visibleRestoredSelectedRange,
    workload: args.workload,
  }
}

export async function captureSameCorpusMutationTargetPhaseScreenshot(args: {
  readonly caseId?: string
  readonly outputPath: string
  readonly page: Page
  readonly phase: SameCorpusMutationTargetScreenshotProof['phase']
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetScreenshotProof> {
  const screenshotPath = mutationTargetScreenshotArtifactPath(args)
  mkdirSync(dirname(screenshotPath), { recursive: true })
  return await captureSameCorpusMutationTargetScreenshotProof({
    page: args.page,
    phase: args.phase,
    product: args.product,
    relativeScreenshotPath: repoRelativePath(screenshotPath),
    screenshotPath,
    target: args.target,
  })
}

function captureTargetScreenshot(
  args: Omit<Parameters<typeof captureSameCorpusMutationTargetPhaseScreenshot>[0], 'phase'>,
  phase: SameCorpusMutationTargetScreenshotProof['phase'],
): Promise<SameCorpusMutationTargetScreenshotProof> {
  return captureSameCorpusMutationTargetPhaseScreenshot({ ...args, phase })
}

function intendedMutationTargetPayload(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['intendedPayload'] {
  if (workload === 'formula-edit') {
    return { kind: 'formula', formula: `=${String(sampleIndex + 1)}+1` }
  }
  if (workload === 'fill-format-change') {
    return {
      kind: 'fill-color',
      expectedFillColor: sameCorpusFillColorExpectedColor(sampleIndex),
      swatchLabel: sameCorpusFillColorSwatchLabel(sampleIndex),
    }
  }
  return { kind: 'cell-value', value: `${product}-same-corpus-${String(sampleIndex + 1)}` }
}

function mutationTargetScreenshotArtifactPath(args: {
  readonly caseId?: string
  readonly outputPath: string
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
  readonly phase: SameCorpusMutationTargetScreenshotProof['phase']
}): string {
  const caseId = args.caseId ?? `same-corpus-${args.workload}`
  return resolve(
    `${args.outputPath}.proof`,
    caseId,
    'mutation-target',
    `${args.product}-sample-${String(args.sampleIndex + 1)}-${args.phase}.png`,
  )
}

function repoRelativePath(path: string): string {
  return relative(process.cwd(), path)
}

function sameCorpusMutationReadbacksEqual(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return (
    left.value === right.value &&
    left.formula === right.formula &&
    left.fillColor === right.fillColor &&
    left.visibleText === right.visibleText &&
    left.source === right.source
  )
}
