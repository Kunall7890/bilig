import type { SameCorpusMutationTargetCommittedStatePhaseProof } from './ui-responsiveness-same-corpus-committed-state-proof.ts'
import type { SameCorpusMutationTargetSelection } from './ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import type {
  SameCorpusMutationTargetReadback,
  SameCorpusMutationTargetScreenshotProof,
} from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from './ui-responsiveness-same-corpus-workloads.ts'
import { sameCorpusMutationTargetBrowserVisibleReadbackSourceAccepted } from './ui-responsiveness-same-corpus-visible-readback-source.ts'

interface SameCorpusMutationTargetPreflightProof {
  readonly before: SameCorpusMutationTargetReadback | null
  readonly beforeCommittedStateProof: SameCorpusMutationTargetCommittedStatePhaseProof | null
  readonly beforeScreenshot: SameCorpusMutationTargetScreenshotProof | null
  readonly product: UiResponsivenessSameCorpusProduct
  readonly sampleIndex: number
  readonly target: SameCorpusMutationTargetSelection | null
  readonly visibleBefore: SameCorpusMutationTargetReadback | null
  readonly workload: UiResponsivenessSameCorpusMutatingWorkload
}

export function assertSameCorpusMutationTargetPreflightProof(args: SameCorpusMutationTargetPreflightProof): void {
  const invalidReasons = sameCorpusMutationTargetPreflightInvalidReasons(args)
  if (invalidReasons.length === 0) {
    return
  }
  throw new Error(
    `same-corpus UI mutation preflight failed for ${args.product} ${args.workload} sample ${String(args.sampleIndex + 1)}: ${invalidReasons.join(
      '; ',
    )}`,
  )
}

export function sameCorpusMutationTargetPreflightInvalidReasons(args: SameCorpusMutationTargetPreflightProof): readonly string[] {
  const invalidReasons: string[] = []
  const target = args.target
  if (!target) {
    invalidReasons.push('missing declared mutation target range')
  } else {
    if (target.sheetName.trim().length === 0) {
      invalidReasons.push('mutation target is missing sheet name')
    }
    if (target.targetRange.trim().length === 0) {
      invalidReasons.push('mutation target is missing target range')
    }
  }
  if (!args.before || args.before.source === 'unknown') {
    invalidReasons.push('missing authoritative before readback for target range')
  }
  if (!args.visibleBefore || !sameCorpusMutationTargetBrowserVisibleReadbackSourceAccepted(args.visibleBefore.source)) {
    invalidReasons.push('missing browser-visible before readback for target cell')
  }
  invalidReasons.push(...sameCorpusBeforeScreenshotInvalidReasons(args))
  if (args.product === 'google-sheets') {
    invalidReasons.push(...sameCorpusGoogleCommittedStatePreflightInvalidReasons(args))
  }
  return [...new Set(invalidReasons)]
}

function sameCorpusBeforeScreenshotInvalidReasons(args: SameCorpusMutationTargetPreflightProof): readonly string[] {
  const invalidReasons: string[] = []
  const screenshot = args.beforeScreenshot
  if (!screenshot) {
    return ['missing before target-cell screenshot proof']
  }
  if (screenshot.phase !== 'before') {
    invalidReasons.push('before screenshot proof has the wrong phase')
  }
  if (screenshot.product !== args.product) {
    invalidReasons.push('before screenshot proof has mismatched product')
  }
  if (screenshot.workload !== args.workload) {
    invalidReasons.push('before screenshot proof has mismatched workload')
  }
  if (screenshot.sampleIndex !== args.sampleIndex) {
    invalidReasons.push('before screenshot proof has mismatched sample')
  }
  if (screenshot.scope !== 'target-cell') {
    invalidReasons.push('before screenshot proof is not scoped to the target cell')
  }
  if (!screenshot.screenshotPath || screenshot.screenshotPath.trim().length === 0) {
    invalidReasons.push('before screenshot proof is missing artifact path')
  }
  if (!isSha256Hex(screenshot.screenshotSha256)) {
    invalidReasons.push('before screenshot proof is missing screenshot SHA256')
  }
  if (args.target) {
    if (screenshot.sheetName !== args.target.sheetName || screenshot.sheetId !== args.target.sheetId) {
      invalidReasons.push('before screenshot proof has mismatched sheet identity')
    }
    if (normalizeRange(screenshot.targetRange) !== normalizeRange(args.target.targetRange)) {
      invalidReasons.push('before screenshot proof has mismatched target range')
    }
  }
  if (args.visibleBefore && !sameCorpusMutationTargetReadbacksEqual(screenshot.semanticReadback, args.visibleBefore)) {
    invalidReasons.push('before screenshot proof semantic readback does not match browser-visible before readback')
  }
  return invalidReasons
}

function sameCorpusGoogleCommittedStatePreflightInvalidReasons(args: SameCorpusMutationTargetPreflightProof): readonly string[] {
  const invalidReasons: string[] = []
  const proof = args.beforeCommittedStateProof
  if (!proof) {
    return ['missing independent Google Sheets before committed-state proof']
  }
  if (proof.product !== 'google-sheets' || proof.product !== args.product) {
    invalidReasons.push('before committed-state proof is not from Google Sheets')
  }
  if (proof.phase !== 'before') {
    invalidReasons.push('before committed-state proof has the wrong phase')
  }
  if (proof.workload !== args.workload) {
    invalidReasons.push('before committed-state proof has mismatched workload')
  }
  if (proof.sampleIndex !== args.sampleIndex) {
    invalidReasons.push('before committed-state proof has mismatched sample')
  }
  if (args.target) {
    if (proof.sheetName !== args.target.sheetName || proof.sheetId !== args.target.sheetId) {
      invalidReasons.push('before committed-state proof has mismatched sheet identity')
    }
    if (normalizeRange(proof.targetRange) !== normalizeRange(args.target.targetRange)) {
      invalidReasons.push('before committed-state proof has mismatched target range')
    }
  }
  if (!proof.artifactPath || proof.artifactPath.trim().length === 0) {
    invalidReasons.push('before committed-state proof is missing artifact path')
  }
  if (!isSha256Hex(proof.artifactSha256)) {
    invalidReasons.push('before committed-state proof is missing artifact SHA256')
  }
  if (args.before && !sameCorpusMutationTargetReadbackMatchesWorkload(args.workload, proof.readback, args.before)) {
    invalidReasons.push('before committed-state proof readback does not match authoritative before readback')
  }
  return invalidReasons
}

function sameCorpusMutationTargetReadbacksEqual(left: SameCorpusMutationTargetReadback, right: SameCorpusMutationTargetReadback): boolean {
  return (
    left.value === right.value &&
    left.formula === right.formula &&
    left.fillColor === right.fillColor &&
    left.visibleText === right.visibleText &&
    left.source === right.source
  )
}

function sameCorpusMutationTargetReadbackMatchesWorkload(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  left: SameCorpusMutationTargetReadback,
  right: SameCorpusMutationTargetReadback,
): boolean {
  if (workload === 'formula-edit') {
    return left.formula === right.formula
  }
  if (workload === 'fill-format-change') {
    return normalizeColor(left.fillColor) === normalizeColor(right.fillColor)
  }
  return left.value === right.value || left.visibleText === right.visibleText
}

function isSha256Hex(value: string | null | undefined): boolean {
  return typeof value === 'string' && /^[a-f0-9]{64}$/u.test(value)
}

function normalizeRange(value: string): string {
  return value.trim().toUpperCase()
}

function normalizeColor(value: string | null): string | null {
  return value ? value.trim().toLowerCase() : null
}
