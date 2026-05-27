import { mkdirSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import type { Page } from '@playwright/test'

import type { UiResponsivenessSameCorpusProduct } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import {
  buildSameCorpusCommittedStateProof,
  captureSameCorpusCommittedStatePhaseProof,
  SameCorpusCommittedStateMismatchError,
  sameCorpusCommittedStateProofArtifactPath,
  type SameCorpusMutationTargetCommittedStatePhaseProof,
} from './ui-responsiveness-same-corpus-committed-state-proof.ts'
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
import {
  readSameCorpusVisibleSelectedRange,
  type SameCorpusMutationTargetIntendedPayload,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'
import {
  SameCorpusMutationTargetCaptureDiagnosticError,
  writeSameCorpusMutationTargetCaptureFailureDiagnostic,
  type SameCorpusMutationTargetCaptureFailurePhase,
  type SameCorpusMutationTargetFailureCleanupStatus,
} from './ui-responsiveness-same-corpus-mutation-target-diagnostic.ts'
import {
  restoreProductWorkbookMutation,
  sameCorpusEditVisibleCellValue,
  sameCorpusFillColorExpectedColor,
  sameCorpusFillColorSwatchLabel,
  sameCorpusFormulaEditFormula,
} from './ui-responsiveness-same-corpus-workload-runner.ts'
import { sameCorpusMutationTargetProofSignature } from './ui-responsiveness-same-corpus-mutation-target-signature.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export async function captureSameCorpusMutationTargetProofForSample(args: {
  readonly before: SameCorpusMutationTargetReadback
  readonly beforeCommittedStateProof?: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly beforeScreenshot: SameCorpusMutationTargetScreenshotProof
  readonly caseId?: string
  readonly intendedPayload?: SameCorpusMutationTargetIntendedPayload | null
  readonly operationStartedAt: number
  readonly outputPath: string
  readonly page: Page
  readonly product: UiResponsivenessSameCorpusProduct
  readonly allowIncompleteCommittedStateProof?: boolean
  readonly readyTimeoutMs?: number
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): Promise<SameCorpusMutationTargetProof> {
  const intendedPayload = args.intendedPayload ?? intendedMutationTargetPayload(args.workload, args.sampleIndex)
  let restoredAfterMutation = false
  let failurePhase: SameCorpusMutationTargetCaptureFailurePhase = 'select-target'
  let after: SameCorpusMutationTargetReadback | null = null
  let visibleAfter: SameCorpusMutationTargetReadback | null = null
  let visibleAfterSelectedRange: string | null = null
  let afterScreenshot: SameCorpusMutationTargetScreenshotProof | null = null
  let afterCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null = null
  let restored: SameCorpusMutationTargetReadback | null = null
  let visibleRestored: SameCorpusMutationTargetReadback | null = null
  let visibleRestoredSelectedRange: string | null = null
  let restoredScreenshot: SameCorpusMutationTargetScreenshotProof | null = null
  let restoredCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null = null
  try {
    await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target })
    failurePhase = 'read-after'
    after = await readSameCorpusMutationTargetReadback({
      page: args.page,
      product: args.product,
      target: args.target,
      workload: args.workload,
    })
    failurePhase = 'read-visible-after'
    visibleAfter = await readSameCorpusVisibleMutationTargetReadback({
      page: args.page,
      product: args.product,
      target: args.target,
      workload: args.workload,
    })
    failurePhase = 'read-visible-after-selection'
    visibleAfterSelectedRange = await readSameCorpusVisibleSelectedRange(args.page, args.product)
    failurePhase = 'capture-after-screenshot'
    afterScreenshot = await captureTargetScreenshot({ ...args, semanticReadback: visibleAfter }, 'after')
    const visibleTargetRenderCapturedAtMs = performance.now()
    const visibleTargetRenderMs = Math.max(0, visibleTargetRenderCapturedAtMs - args.operationStartedAt)
    failurePhase = 'capture-after-committed-state'
    afterCommittedStateProof = await maybeCaptureIncompleteCommittedStatePhaseProof(
      {
        allowIncompleteCommittedStateProof: args.allowIncompleteCommittedStateProof === true,
      },
      captureSameCorpusCommittedStatePhaseProof({
        artifactPath: committedStateArtifactPath(args, 'after'),
        expectedReadback: sameCorpusCommittedStateExpectedReadback({
          before: args.before,
          intendedPayload,
          phase: 'after',
          phaseReadback: after,
          sampleIndex: args.sampleIndex,
          workload: args.workload,
        }),
        page: args.page,
        phase: 'after',
        product: args.product,
        timeoutMs: args.readyTimeoutMs,
        sampleIndex: args.sampleIndex,
        target: args.target,
        workload: args.workload,
      }),
    )
    failurePhase = 'read-revisions'
    const revisions = await readSameCorpusMutationTargetRevisionProof({
      page: args.page,
      product: args.product,
      readback: after,
      screenshotSha256: afterScreenshot.screenshotSha256,
      target: args.target,
    })
    const postMutationProofCapturedAtMs = performance.now()
    const committedTargetProofMs = Math.max(0, postMutationProofCapturedAtMs - args.operationStartedAt)
    const committedStateValidationMs = Math.max(0, postMutationProofCapturedAtMs - visibleTargetRenderCapturedAtMs)
    failurePhase = 'restore-mutation'
    await restoreProductWorkbookMutation(args.page, args.workload)
    restoredAfterMutation = true
    await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target })
    failurePhase = 'read-restored'
    restored = await readSameCorpusMutationTargetReadback({
      page: args.page,
      product: args.product,
      target: args.target,
      workload: args.workload,
    })
    failurePhase = 'read-visible-restored'
    visibleRestored = await readSameCorpusVisibleMutationTargetReadback({
      page: args.page,
      product: args.product,
      target: args.target,
      workload: args.workload,
    })
    failurePhase = 'read-visible-restored-selection'
    visibleRestoredSelectedRange = await readSameCorpusVisibleSelectedRange(args.page, args.product)
    failurePhase = 'capture-restored-screenshot'
    restoredScreenshot = await captureTargetScreenshot({ ...args, semanticReadback: visibleRestored }, 'restored')
    failurePhase = 'capture-restored-committed-state'
    restoredCommittedStateProof = await maybeCaptureIncompleteCommittedStatePhaseProof(
      {
        allowIncompleteCommittedStateProof: args.allowIncompleteCommittedStateProof === true,
      },
      captureSameCorpusCommittedStatePhaseProof({
        artifactPath: committedStateArtifactPath(args, 'restored'),
        expectedReadback: sameCorpusCommittedStateExpectedReadback({
          before: args.before,
          intendedPayload,
          phase: 'restored',
          phaseReadback: restored,
          sampleIndex: args.sampleIndex,
          workload: args.workload,
        }),
        page: args.page,
        phase: 'restored',
        product: args.product,
        timeoutMs: args.readyTimeoutMs,
        sampleIndex: args.sampleIndex,
        target: args.target,
        workload: args.workload,
      }),
    )
    const restoreProofCapturedAtMs = performance.now()
    const restoreValidationMs = Math.max(0, restoreProofCapturedAtMs - postMutationProofCapturedAtMs)
    failurePhase = 'build-proof'
    const proof: Omit<SameCorpusMutationTargetProof, 'targetProofSignature'> = {
      after,
      authoritativeReadbackRevision: revisions.authoritativeReadbackRevision,
      before: args.before,
      committedStateValidationMs,
      committedTargetProofMs,
      committedStateProof: buildSameCorpusCommittedStateProof({
        after: afterCommittedStateProof,
        before: args.beforeCommittedStateProof ?? null,
        product: args.product,
        restored: restoredCommittedStateProof,
        sampleIndex: args.sampleIndex,
        target: args.target,
        workload: args.workload,
      }),
      intendedOperation: args.workload,
      intendedPayload,
      operationStartedAtMs: args.operationStartedAt,
      postMutationProofCapturedAtMs,
      product: args.product,
      restored,
      restoreValidationMs,
      restoreProofCapturedAtMs,
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
      visibleTargetRenderCapturedAtMs,
      visibleTargetRenderMs,
      visibleRestored,
      visibleRestoredSelectedRange,
      workload: args.workload,
    }
    return {
      ...proof,
      targetProofSignature: sameCorpusMutationTargetProofSignature(proof),
    }
  } catch (error: unknown) {
    let restoreStatus: SameCorpusMutationTargetFailureCleanupStatus = restoredAfterMutation ? 'not-needed' : 'restored'
    let restoreError: unknown = null
    if (!restoredAfterMutation) {
      try {
        await restoreProductWorkbookMutation(args.page, args.workload)
        restoredAfterMutation = true
      } catch (cleanupError: unknown) {
        restoreStatus = 'failed'
        restoreError = cleanupError
      }
    }
    let reselectStatus: SameCorpusMutationTargetFailureCleanupStatus = 'restored'
    let reselectError: unknown = null
    try {
      await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target })
    } catch (cleanupError: unknown) {
      reselectStatus = 'failed'
      reselectError = cleanupError
    }
    let diagnosticArtifactPath: string | null = null
    try {
      diagnosticArtifactPath = writeSameCorpusMutationTargetCaptureFailureDiagnostic({
        after,
        afterCommittedStateProof,
        afterScreenshot,
        before: args.before,
        beforeCommittedStateProof: args.beforeCommittedStateProof ?? null,
        beforeScreenshot: args.beforeScreenshot,
        caseId: args.caseId,
        error,
        failurePhase,
        intendedPayload,
        operationStartedAt: args.operationStartedAt,
        outputPath: args.outputPath,
        product: args.product,
        restoreError,
        restored,
        restoredCommittedStateProof,
        restoredScreenshot,
        restoreStatus,
        reselectError,
        reselectStatus,
        sampleIndex: args.sampleIndex,
        target: args.target,
        visibleAfter,
        visibleAfterSelectedRange,
        visibleRestored,
        visibleRestoredSelectedRange,
        workload: args.workload,
      })
    } catch {
      // Preserve the original capture failure; accepted or stale committed-state phase artifacts remain on disk.
    }
    if (diagnosticArtifactPath) {
      throw new SameCorpusMutationTargetCaptureDiagnosticError({
        artifactPath: diagnosticArtifactPath,
        cause: error,
        failurePhase,
        product: args.product,
        sampleIndex: args.sampleIndex,
        targetRange: args.target.targetRange,
        workload: args.workload,
      })
    }
    throw error
  } finally {
    if (!restoredAfterMutation) {
      await restoreProductWorkbookMutation(args.page, args.workload).catch(() => undefined)
      await selectSameCorpusMutationTargetRange({ page: args.page, product: args.product, target: args.target }).catch(() => undefined)
    }
  }
}

export async function captureSameCorpusMutationTargetPhaseScreenshot(args: {
  readonly caseId?: string
  readonly outputPath: string
  readonly page: Page
  readonly phase: SameCorpusMutationTargetScreenshotProof['phase']
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly semanticReadback: SameCorpusMutationTargetReadback
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
    sampleIndex: args.sampleIndex,
    semanticReadback: args.semanticReadback,
    screenshotPath,
    target: args.target,
    workload: args.workload,
  })
}

function captureTargetScreenshot(
  args: Omit<Parameters<typeof captureSameCorpusMutationTargetPhaseScreenshot>[0], 'phase'>,
  phase: SameCorpusMutationTargetScreenshotProof['phase'],
): Promise<SameCorpusMutationTargetScreenshotProof> {
  return captureSameCorpusMutationTargetPhaseScreenshot({ ...args, phase })
}

function committedStateArtifactPath(
  args: Pick<Parameters<typeof captureSameCorpusMutationTargetProofForSample>[0], 'caseId' | 'outputPath' | 'sampleIndex' | 'workload'>,
  phase: SameCorpusMutationTargetCommittedStatePhaseProof['phase'],
): string {
  return sameCorpusCommittedStateProofArtifactPath({
    caseId: args.caseId,
    outputPath: args.outputPath,
    phase,
    sampleIndex: args.sampleIndex,
    workload: args.workload,
  })
}

function intendedMutationTargetPayload(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['intendedPayload'] {
  if (workload === 'formula-edit') {
    return { kind: 'formula', formula: sameCorpusFormulaEditFormula(sampleIndex) }
  }
  if (workload === 'fill-format-change') {
    return {
      kind: 'fill-color',
      expectedFillColor: sameCorpusFillColorExpectedColor(sampleIndex),
      swatchLabel: sameCorpusFillColorSwatchLabel(sampleIndex),
    }
  }
  return { kind: 'cell-value', value: sameCorpusEditVisibleCellValue(sampleIndex) }
}

export async function maybeCaptureIncompleteCommittedStatePhaseProof(
  options: { readonly allowIncompleteCommittedStateProof: boolean },
  proof: Promise<SameCorpusMutationTargetCommittedStatePhaseProof>,
): Promise<SameCorpusMutationTargetCommittedStatePhaseProof | null> {
  try {
    return await proof
  } catch (error: unknown) {
    if (options.allowIncompleteCommittedStateProof && error instanceof SameCorpusCommittedStateMismatchError) {
      return null
    }
    throw error
  }
}

export function sameCorpusCommittedStateExpectedReadback(args: {
  readonly before: SameCorpusMutationTargetReadback
  readonly intendedPayload?: SameCorpusMutationTargetIntendedPayload | null
  readonly phase: 'after' | 'restored'
  readonly phaseReadback: SameCorpusMutationTargetReadback
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): SameCorpusMutationTargetReadback {
  if (args.phase === 'restored') {
    return args.before
  }
  if (args.workload === 'fill-format-change') {
    const fillColor =
      args.intendedPayload?.kind === 'fill-color'
        ? args.intendedPayload.expectedFillColor
        : sameCorpusFillColorExpectedColor(args.sampleIndex)
    return { ...args.phaseReadback, fillColor }
  }
  return args.phaseReadback
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
