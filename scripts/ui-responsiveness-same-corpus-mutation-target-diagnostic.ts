import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'

import {
  SameCorpusCommittedStateMismatchError,
  sameCorpusCommittedStateProofArtifactPath,
  type SameCorpusCommittedStateMismatchDiagnostic,
  type SameCorpusMutationTargetCommittedStatePhaseProof,
} from './ui-responsiveness-same-corpus-committed-state-proof.ts'
import type { SameCorpusMutationTargetSelection } from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { SameCorpusMutationTargetReadback, SameCorpusMutationTargetScreenshotProof } from './ui-responsiveness-same-corpus-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export type SameCorpusMutationTargetCaptureFailurePhase =
  | 'select-target'
  | 'read-after'
  | 'read-visible-after'
  | 'read-visible-after-selection'
  | 'capture-after-screenshot'
  | 'capture-after-committed-state'
  | 'read-revisions'
  | 'restore-mutation'
  | 'read-restored'
  | 'read-visible-restored'
  | 'read-visible-restored-selection'
  | 'capture-restored-screenshot'
  | 'capture-restored-committed-state'
  | 'build-proof'

export type SameCorpusMutationTargetFailureCleanupStatus = 'not-needed' | 'restored' | 'failed'

export interface SameCorpusMutationTargetCaptureFailureDiagnostic {
  readonly schemaVersion: 1
  readonly artifactKind: 'same-corpus-mutation-target-failure-diagnostic'
  readonly product: UiResponsivenessSameCorpusProduct
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
  readonly sampleIndex: number
  readonly sheetName: string
  readonly sheetId: string | null
  readonly targetRange: string
  readonly failurePhase: SameCorpusMutationTargetCaptureFailurePhase
  readonly operationStartedAtMs: number
  readonly failedAtMs: number
  readonly cleanup: {
    readonly restoreStatus: SameCorpusMutationTargetFailureCleanupStatus
    readonly reselectStatus: SameCorpusMutationTargetFailureCleanupStatus
    readonly restoreError: SerializedSameCorpusCaptureError | null
    readonly reselectError: SerializedSameCorpusCaptureError | null
  }
  readonly intendedPayload:
    | {
        readonly kind: 'cell-value'
        readonly value: string
      }
    | {
        readonly kind: 'formula'
        readonly formula: string
      }
    | {
        readonly kind: 'fill-color'
        readonly expectedFillColor: string
        readonly swatchLabel: string
      }
  readonly before: SameCorpusMutationTargetReadback
  readonly after: SameCorpusMutationTargetReadback | null
  readonly restored: SameCorpusMutationTargetReadback | null
  readonly visibleAfter: SameCorpusMutationTargetReadback | null
  readonly visibleRestored: SameCorpusMutationTargetReadback | null
  readonly visibleAfterSelectedRange: string | null
  readonly visibleRestoredSelectedRange: string | null
  readonly beforeScreenshot: SameCorpusMutationTargetScreenshotProof
  readonly afterScreenshot: SameCorpusMutationTargetScreenshotProof | null
  readonly restoredScreenshot: SameCorpusMutationTargetScreenshotProof | null
  readonly beforeCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly afterCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly restoredCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly committedStateMismatch: SameCorpusCommittedStateMismatchDiagnostic | null
  readonly error: SerializedSameCorpusCaptureError
}

export interface SerializedSameCorpusCaptureError {
  readonly name: string
  readonly message: string
  readonly stack: string | null
  readonly cause: SerializedSameCorpusCaptureError | null
}

export class SameCorpusMutationTargetCaptureDiagnosticError extends Error {
  readonly diagnosticArtifactPath: string
  readonly failurePhase: SameCorpusMutationTargetCaptureFailurePhase

  constructor(args: {
    readonly artifactPath: string
    readonly cause: unknown
    readonly failurePhase: SameCorpusMutationTargetCaptureFailurePhase
    readonly product: UiResponsivenessSameCorpusProduct
    readonly sampleIndex: number
    readonly targetRange: string
    readonly workload: UiResponsivenessSameCorpusMutatingWorkload
  }) {
    const causeMessage = args.cause instanceof Error ? args.cause.message : String(args.cause)
    super(
      `Same-corpus mutation target proof failed for ${args.product} ${args.workload} sample ${String(args.sampleIndex + 1)} ${
        args.targetRange
      } during ${args.failurePhase}. Diagnostic artifact: ${args.artifactPath}. Cause: ${causeMessage}`,
      { cause: args.cause },
    )
    this.name = 'SameCorpusMutationTargetCaptureDiagnosticError'
    this.diagnosticArtifactPath = args.artifactPath
    this.failurePhase = args.failurePhase
  }
}

export function sameCorpusMutationTargetFailureDiagnosticArtifactPath(args: {
  readonly caseId?: string
  readonly outputPath: string
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): string {
  const basePath = sameCorpusCommittedStateProofArtifactPath({
    caseId: args.caseId,
    outputPath: args.outputPath,
    phase: 'after',
    sampleIndex: args.sampleIndex,
    workload: args.workload,
  })
  return resolve(
    dirname(dirname(basePath)),
    'mutation-target-diagnostics',
    `${args.product}-sample-${String(args.sampleIndex + 1)}-failure.json`,
  )
}

export function writeSameCorpusMutationTargetCaptureFailureDiagnostic(args: {
  readonly after: SameCorpusMutationTargetReadback | null
  readonly afterCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly afterScreenshot: SameCorpusMutationTargetScreenshotProof | null
  readonly before: SameCorpusMutationTargetReadback
  readonly beforeCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly beforeScreenshot: SameCorpusMutationTargetScreenshotProof
  readonly caseId?: string
  readonly error: unknown
  readonly failurePhase: SameCorpusMutationTargetCaptureFailurePhase
  readonly intendedPayload: SameCorpusMutationTargetCaptureFailureDiagnostic['intendedPayload']
  readonly operationStartedAt: number
  readonly outputPath: string
  readonly product: UiResponsivenessSameCorpusProduct
  readonly restoreError: unknown
  readonly restored: SameCorpusMutationTargetReadback | null
  readonly restoredCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly restoredScreenshot: SameCorpusMutationTargetScreenshotProof | null
  readonly restoreStatus: SameCorpusMutationTargetFailureCleanupStatus
  readonly reselectError: unknown
  readonly reselectStatus: SameCorpusMutationTargetFailureCleanupStatus
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection
  readonly visibleAfter: SameCorpusMutationTargetReadback | null
  readonly visibleAfterSelectedRange: string | null
  readonly visibleRestored: SameCorpusMutationTargetReadback | null
  readonly visibleRestoredSelectedRange: string | null
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}): string {
  const artifactPath = sameCorpusMutationTargetFailureDiagnosticArtifactPath(args)
  mkdirSync(dirname(artifactPath), { recursive: true })
  const diagnostic: SameCorpusMutationTargetCaptureFailureDiagnostic = {
    schemaVersion: 1,
    artifactKind: 'same-corpus-mutation-target-failure-diagnostic',
    product: args.product,
    workload: args.workload,
    sampleIndex: args.sampleIndex,
    sheetName: args.target.sheetName,
    sheetId: args.target.sheetId,
    targetRange: args.target.targetRange,
    failurePhase: args.failurePhase,
    operationStartedAtMs: args.operationStartedAt,
    failedAtMs: performance.now(),
    cleanup: {
      restoreStatus: args.restoreStatus,
      reselectStatus: args.reselectStatus,
      restoreError: args.restoreError === null ? null : serializeSameCorpusCaptureError(args.restoreError),
      reselectError: args.reselectError === null ? null : serializeSameCorpusCaptureError(args.reselectError),
    },
    intendedPayload: args.intendedPayload,
    before: args.before,
    after: args.after,
    restored: args.restored,
    visibleAfter: args.visibleAfter,
    visibleRestored: args.visibleRestored,
    visibleAfterSelectedRange: args.visibleAfterSelectedRange,
    visibleRestoredSelectedRange: args.visibleRestoredSelectedRange,
    beforeScreenshot: args.beforeScreenshot,
    afterScreenshot: args.afterScreenshot,
    restoredScreenshot: args.restoredScreenshot,
    beforeCommittedStateProof: args.beforeCommittedStateProof,
    afterCommittedStateProof: args.afterCommittedStateProof,
    restoredCommittedStateProof: args.restoredCommittedStateProof,
    committedStateMismatch: args.error instanceof SameCorpusCommittedStateMismatchError ? args.error.diagnostic : null,
    error: serializeSameCorpusCaptureError(args.error),
  }
  writeFileSync(artifactPath, `${JSON.stringify(stableJsonValue(diagnostic), null, 2)}\n`)
  return repoRelativePath(artifactPath)
}

function serializeSameCorpusCaptureError(error: unknown): SerializedSameCorpusCaptureError {
  if (!(error instanceof Error)) {
    return {
      name: typeof error,
      message: String(error),
      stack: null,
      cause: null,
    }
  }
  const cause = Reflect.get(error, 'cause')
  return {
    name: error.name,
    message: error.message,
    stack: error.stack ?? null,
    cause: cause === undefined ? null : serializeSameCorpusCaptureError(cause),
  }
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

function repoRelativePath(path: string): string {
  return relative(resolve(new URL('..', import.meta.url).pathname), path).replaceAll('\\', '/')
}
