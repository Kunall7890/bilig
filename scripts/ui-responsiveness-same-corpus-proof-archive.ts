import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, isAbsolute, resolve } from 'node:path'

import { strFromU8, unzipSync } from 'fflate'

import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import type { SameCorpusMutationTargetReadback } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import { validateSameCorpusProductSemanticUiProof } from './ui-responsiveness-same-corpus-semantic-proof.ts'
import type { SameCorpusCapture, UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { arrayField, asObject, booleanField, literalField, numberField, stringField } from './json-scorecard-helpers.ts'
import { sameCorpusFillColorsMatch } from './ui-responsiveness-same-corpus-fill-proof.ts'
import { requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads } from './ui-responsiveness-same-corpus-mutation-target-proof-summary.ts'
import { sameCorpusFillColorExpectedColor } from './ui-responsiveness-same-corpus-workload-runner.ts'
import {
  isUiResponsivenessSameCorpusWorkload,
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadMutatesWorkbook,
  type UiResponsivenessSameCorpusWorkload,
} from './ui-responsiveness-same-corpus-workloads.ts'

interface SameCorpusProofArchiveSummaryCase {
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly scenarioProof: SameCorpusScenarioProof
}

export interface SameCorpusProofArchiveManifest {
  readonly schemaVersion: 1
  readonly suite: 'ui-responsiveness-same-corpus-proof-archive'
  readonly captureRunSignature: string
  readonly requiredArtifactCount: number
  readonly artifactCount: number
  readonly filesVerified: boolean
  readonly complete: boolean
  readonly fileVerification: SameCorpusProofArchiveFileVerification
  readonly artifacts: readonly SameCorpusProofArchiveArtifact[]
}

export type SameCorpusProofArchiveArtifact =
  | SameCorpusScenarioScreenshotArchiveArtifact
  | SameCorpusMutationTargetScreenshotArchiveArtifact
  | SameCorpusGoogleSheetsCommittedStateArchiveArtifact

export interface SameCorpusScenarioScreenshotArchiveArtifact {
  readonly kind: 'scenario-screenshot'
  readonly product: UiResponsivenessSameCorpusProduct
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly path: string
  readonly screenshotSha256: string
}

export interface SameCorpusMutationTargetScreenshotArchiveArtifact {
  readonly kind: 'mutation-target-screenshot'
  readonly product: UiResponsivenessSameCorpusProduct
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleIndex: number
  readonly phase: 'before' | 'after' | 'restored'
  readonly path: string
  readonly screenshotSha256: string
}

export interface SameCorpusGoogleSheetsCommittedStateArchiveArtifact {
  readonly kind: 'google-sheets-committed-state-export'
  readonly product: 'google-sheets'
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleIndex: number
  readonly phase: 'before' | 'after' | 'restored'
  readonly sheetName: string
  readonly sheetId: string | null
  readonly targetRange: string
  readonly capturedAtMs: number
  readonly artifactPath: string
  readonly artifactSha256: string
  readonly exportUrl: string
  readonly workbookByteSize: number
  readonly workbookSha256: string
  readonly readback: SameCorpusMutationTargetReadback
  readonly readbackSha256: string
}

export interface SameCorpusProofArchiveBuildOptions {
  readonly artifactBaseDir?: string
}

export interface SameCorpusProofArchiveZipVerificationOptions {
  readonly entryRootDir?: string
  readonly manifestEntryPath?: string
}

export interface SameCorpusProofArchiveZipVerification {
  readonly schemaVersion: 1
  readonly archivePath: string
  readonly manifestEntryPath: string
  readonly manifestSha256: string
  readonly filesVerified: boolean
  readonly complete: boolean
  readonly manifest: SameCorpusProofArchiveManifest
  readonly fileVerification: SameCorpusProofArchiveFileVerification
}

export interface SameCorpusProofArchiveFileVerification {
  readonly schemaVersion: 1
  readonly checkedArtifactCount: number
  readonly verifiedArtifactCount: number
  readonly missingArtifactCount: number
  readonly mismatchedArtifactCount: number
  readonly complete: boolean
  readonly entries: readonly SameCorpusProofArchiveFileVerificationEntry[]
}

export interface SameCorpusProofArchiveFileVerificationEntry {
  readonly status: 'verified' | 'missing' | 'hash-mismatch' | 'identity-mismatch'
  readonly kind: SameCorpusProofArchiveArtifact['kind']
  readonly product: UiResponsivenessSameCorpusProduct
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly sampleIndex?: number
  readonly phase?: 'before' | 'after' | 'restored'
  readonly path: string
  readonly resolvedPath: string
  readonly expectedSha256: string
  readonly actualSha256?: string
  readonly identityMismatchReason?: string
}

const requiredProofArchiveProducts = ['bilig', 'google-sheets'] as const satisfies readonly UiResponsivenessSameCorpusProduct[]
const mutationTargetScreenshotPhaseCount = 3
const googleSheetsCommittedStatePhaseCount = 3

export function buildSameCorpusProofArchiveManifest(
  capture: SameCorpusCapture,
  options: SameCorpusProofArchiveBuildOptions = {},
): SameCorpusProofArchiveManifest {
  const artifacts = sameCorpusProofArchiveArtifacts(capture.cases, capture.sampleCount)
  const requiredArtifactCount = capture.runManifest.requiredProofArchiveArtifactCount
  const fileVerification = verifySameCorpusProofArchiveFiles(artifacts, options)
  const allRequiredArtifactsPresent = artifacts.length === requiredArtifactCount
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-proof-archive',
    captureRunSignature: capture.runManifest.captureRunSignature,
    requiredArtifactCount,
    artifactCount: artifacts.length,
    filesVerified: allRequiredArtifactsPresent && fileVerification.complete,
    complete: allRequiredArtifactsPresent && fileVerification.complete,
    fileVerification,
    artifacts,
  }
}

export function proofArchiveManifestPath(outputPath: string): string {
  return resolve(`${outputPath}.proof`, 'proof-archive-manifest.json')
}

export function writeSameCorpusProofArchiveManifest(capture: SameCorpusCapture, outputPath: string): SameCorpusProofArchiveManifest {
  const manifest = buildSameCorpusProofArchiveManifest(capture)
  const manifestPath = proofArchiveManifestPath(outputPath)
  mkdirSync(dirname(manifestPath), { recursive: true })
  writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
  return manifest
}

export function readSameCorpusProofArchiveManifest(manifestPath: string): SameCorpusProofArchiveManifest {
  return parseSameCorpusProofArchiveManifest(JSON.parse(readFileSync(manifestPath, 'utf8')) as unknown)
}

function parseSameCorpusProofArchiveManifest(value: unknown): SameCorpusProofArchiveManifest {
  const record = asObject(value, 'same-corpus proof archive manifest')
  return {
    schemaVersion: literalField(record, 'schemaVersion', 1),
    suite: literalField(record, 'suite', 'ui-responsiveness-same-corpus-proof-archive'),
    captureRunSignature: stringField(record, 'captureRunSignature'),
    requiredArtifactCount: numberField(record, 'requiredArtifactCount'),
    artifactCount: numberField(record, 'artifactCount'),
    filesVerified: booleanField(record, 'filesVerified'),
    complete: booleanField(record, 'complete'),
    fileVerification: parseSameCorpusProofArchiveFileVerification(record.fileVerification),
    artifacts: arrayField(record, 'artifacts').map(parseSameCorpusProofArchiveArtifact),
  }
}

function parseSameCorpusProofArchiveFileVerification(value: unknown): SameCorpusProofArchiveFileVerification {
  const record = asObject(value, 'same-corpus proof archive file verification')
  return {
    schemaVersion: literalField(record, 'schemaVersion', 1),
    checkedArtifactCount: numberField(record, 'checkedArtifactCount'),
    verifiedArtifactCount: numberField(record, 'verifiedArtifactCount'),
    missingArtifactCount: numberField(record, 'missingArtifactCount'),
    mismatchedArtifactCount: numberField(record, 'mismatchedArtifactCount'),
    complete: booleanField(record, 'complete'),
    entries: arrayField(record, 'entries').map(parseSameCorpusProofArchiveFileVerificationEntry),
  }
}

function parseSameCorpusProofArchiveFileVerificationEntry(value: unknown): SameCorpusProofArchiveFileVerificationEntry {
  const record = asObject(value, 'same-corpus proof archive file verification entry')
  const actualSha256 = Object.hasOwn(record, 'actualSha256') ? { actualSha256: stringField(record, 'actualSha256') } : {}
  const identityMismatchReason = Object.hasOwn(record, 'identityMismatchReason')
    ? { identityMismatchReason: stringField(record, 'identityMismatchReason') }
    : {}
  return {
    status: parseSameCorpusProofArchiveFileVerificationStatus(stringField(record, 'status')),
    kind: parseSameCorpusProofArchiveArtifactKind(stringField(record, 'kind')),
    product: parseSameCorpusProofArchiveProduct(stringField(record, 'product')),
    workload: parseSameCorpusProofArchiveWorkload(stringField(record, 'workload')),
    ...(Object.hasOwn(record, 'sampleIndex') ? { sampleIndex: numberField(record, 'sampleIndex') } : {}),
    ...(Object.hasOwn(record, 'phase') ? { phase: parseSameCorpusProofArchivePhase(stringField(record, 'phase')) } : {}),
    path: stringField(record, 'path'),
    resolvedPath: stringField(record, 'resolvedPath'),
    expectedSha256: stringField(record, 'expectedSha256'),
    ...actualSha256,
    ...identityMismatchReason,
  }
}

function parseSameCorpusProofArchiveArtifact(value: unknown): SameCorpusProofArchiveArtifact {
  const record = asObject(value, 'same-corpus proof archive artifact')
  const kind = parseSameCorpusProofArchiveArtifactKind(stringField(record, 'kind'))
  const workload = parseSameCorpusProofArchiveWorkload(stringField(record, 'workload'))
  if (kind === 'scenario-screenshot') {
    return {
      kind,
      product: parseSameCorpusProofArchiveProduct(stringField(record, 'product')),
      workload,
      path: stringField(record, 'path'),
      screenshotSha256: stringField(record, 'screenshotSha256'),
    }
  }
  if (kind === 'mutation-target-screenshot') {
    return {
      kind,
      product: parseSameCorpusProofArchiveProduct(stringField(record, 'product')),
      workload,
      sampleIndex: numberField(record, 'sampleIndex'),
      phase: parseSameCorpusProofArchivePhase(stringField(record, 'phase')),
      path: stringField(record, 'path'),
      screenshotSha256: stringField(record, 'screenshotSha256'),
    }
  }
  return {
    kind,
    product: literalField(record, 'product', 'google-sheets'),
    workload,
    sampleIndex: numberField(record, 'sampleIndex'),
    phase: parseSameCorpusProofArchivePhase(stringField(record, 'phase')),
    sheetName: stringField(record, 'sheetName'),
    sheetId: Object.hasOwn(record, 'sheetId') && record.sheetId !== null ? stringField(record, 'sheetId') : null,
    targetRange: stringField(record, 'targetRange'),
    capturedAtMs: numberField(record, 'capturedAtMs'),
    artifactPath: stringField(record, 'artifactPath'),
    artifactSha256: stringField(record, 'artifactSha256'),
    exportUrl: stringField(record, 'exportUrl'),
    workbookByteSize: numberField(record, 'workbookByteSize'),
    workbookSha256: stringField(record, 'workbookSha256'),
    readback: parseCommittedStateReadback(record.readback),
    readbackSha256: stringField(record, 'readbackSha256'),
  }
}

function parseSameCorpusProofArchiveArtifactKind(value: string): SameCorpusProofArchiveArtifact['kind'] {
  if (value === 'scenario-screenshot' || value === 'mutation-target-screenshot' || value === 'google-sheets-committed-state-export') {
    return value
  }
  throw new Error(`Unexpected same-corpus proof archive artifact kind: ${value}`)
}

function parseSameCorpusProofArchiveFileVerificationStatus(value: string): SameCorpusProofArchiveFileVerificationEntry['status'] {
  if (value === 'verified' || value === 'missing' || value === 'hash-mismatch' || value === 'identity-mismatch') {
    return value
  }
  throw new Error(`Unexpected same-corpus proof archive artifact verification status: ${value}`)
}

function parseCommittedStateReadback(value: unknown): SameCorpusMutationTargetReadback {
  const record = asObject(value, 'same-corpus committed-state readback')
  return {
    value: nullableStringField(record, 'value'),
    formula: nullableStringField(record, 'formula'),
    fillColor: nullableStringField(record, 'fillColor'),
    visibleText: nullableStringField(record, 'visibleText'),
    source: literalField(record, 'source', 'google-sheets-xlsx-export'),
  }
}

function nullableStringField(record: Record<string, unknown>, key: string): string | null {
  return Object.hasOwn(record, key) && record[key] !== null ? stringField(record, key) : null
}

function parseSameCorpusProofArchiveProduct(value: string): UiResponsivenessSameCorpusProduct {
  if (value === 'bilig' || value === 'google-sheets' || value === 'microsoft-excel-web') {
    return value
  }
  throw new Error(`Unexpected same-corpus proof archive product: ${value}`)
}

function parseSameCorpusProofArchiveWorkload(value: string): UiResponsivenessSameCorpusWorkload {
  if (isUiResponsivenessSameCorpusWorkload(value)) {
    return value
  }
  throw new Error(`Unexpected same-corpus proof archive workload: ${value}`)
}

function parseSameCorpusProofArchivePhase(value: string): 'before' | 'after' | 'restored' {
  if (value === 'before' || value === 'after' || value === 'restored') {
    return value
  }
  throw new Error(`Unexpected same-corpus proof archive phase: ${value}`)
}

export function verifySameCorpusProofArchiveManifestPath(
  manifestPath: string,
  options: SameCorpusProofArchiveBuildOptions = {},
): SameCorpusProofArchiveManifest {
  const manifest = readSameCorpusProofArchiveManifest(manifestPath)
  const fileVerification = verifySameCorpusProofArchiveFiles(manifest.artifacts, options)
  const allRequiredArtifactsPresent = manifest.artifactCount === manifest.requiredArtifactCount
  return {
    ...manifest,
    filesVerified: allRequiredArtifactsPresent && fileVerification.complete,
    complete: allRequiredArtifactsPresent && fileVerification.complete,
    fileVerification,
  }
}

export function verifySameCorpusProofArchiveZipPath(
  archivePath: string,
  options: SameCorpusProofArchiveZipVerificationOptions = {},
): SameCorpusProofArchiveZipVerification {
  const resolvedArchivePath = resolve(archivePath)
  const entriesByPath = normalizeSameCorpusProofArchiveZipEntries(unzipSync(new Uint8Array(readFileSync(resolvedArchivePath))))
  const manifestEntryPath = resolveSameCorpusProofArchiveZipManifestEntry(entriesByPath, options.manifestEntryPath)
  const manifestBytes = entriesByPath.get(manifestEntryPath)
  if (!manifestBytes) {
    throw new Error(`UI responsiveness same-corpus proof archive ZIP manifest is missing: ${manifestEntryPath}`)
  }
  const manifest = parseSameCorpusProofArchiveManifest(JSON.parse(strFromU8(manifestBytes)) as unknown)
  const fileVerification = verifySameCorpusProofArchiveZipFiles(manifest.artifacts, entriesByPath, {
    archivePath: resolvedArchivePath,
    entryRootDir: options.entryRootDir,
  })
  const allRequiredArtifactsPresent = manifest.artifactCount === manifest.requiredArtifactCount
  return {
    schemaVersion: 1,
    archivePath: resolvedArchivePath,
    manifestEntryPath,
    manifestSha256: sha256Hex(manifestBytes),
    filesVerified: allRequiredArtifactsPresent && fileVerification.complete,
    complete: allRequiredArtifactsPresent && fileVerification.complete,
    manifest,
    fileVerification,
  }
}

export function verifySameCorpusProofArchiveFiles(
  artifacts: readonly SameCorpusProofArchiveArtifact[],
  options: SameCorpusProofArchiveBuildOptions = {},
): SameCorpusProofArchiveFileVerification {
  const entries = artifacts.map((artifact) => verifySameCorpusProofArchiveFile(artifact, options))
  return summarizeSameCorpusProofArchiveFileVerification(entries)
}

function verifySameCorpusProofArchiveZipFiles(
  artifacts: readonly SameCorpusProofArchiveArtifact[],
  entriesByPath: ReadonlyMap<string, Uint8Array>,
  options: { readonly archivePath: string; readonly entryRootDir?: string },
): SameCorpusProofArchiveFileVerification {
  const entries = artifacts.map((artifact) => verifySameCorpusProofArchiveZipFile(artifact, entriesByPath, options))
  return summarizeSameCorpusProofArchiveFileVerification(entries)
}

function summarizeSameCorpusProofArchiveFileVerification(
  entries: readonly SameCorpusProofArchiveFileVerificationEntry[],
): SameCorpusProofArchiveFileVerification {
  const verifiedArtifactCount = entries.filter((entry) => entry.status === 'verified').length
  const missingArtifactCount = entries.filter((entry) => entry.status === 'missing').length
  const mismatchedArtifactCount = entries.filter((entry) => entry.status === 'hash-mismatch' || entry.status === 'identity-mismatch').length
  return {
    schemaVersion: 1,
    checkedArtifactCount: entries.length,
    verifiedArtifactCount,
    missingArtifactCount,
    mismatchedArtifactCount,
    complete: entries.length === verifiedArtifactCount,
    entries,
  }
}

export function requiredUiResponsivenessSameCorpusProofArchiveArtifactCount(sampleCount: number): number {
  const normalizedSampleCount = Math.max(0, sampleCount)
  return (
    requiredUiResponsivenessSameCorpusWorkloads.length * requiredProofArchiveProducts.length +
    requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads.length *
      normalizedSampleCount *
      (requiredProofArchiveProducts.length * mutationTargetScreenshotPhaseCount + googleSheetsCommittedStatePhaseCount)
  )
}

export function sameCorpusProofArchiveArtifactCount(cases: readonly SameCorpusProofArchiveSummaryCase[], sampleCount: number): number {
  return (
    sameCorpusScenarioScreenshotArtifactCount(cases, sampleCount) +
    sameCorpusMutationTargetScreenshotArtifactCount(cases, sampleCount) +
    sameCorpusGoogleSheetsCommittedStateArtifactCount(cases, sampleCount)
  )
}

function sameCorpusScenarioScreenshotArtifactCount(cases: readonly SameCorpusProofArchiveSummaryCase[], sampleCount: number): number {
  return cases.reduce((total, entry) => {
    return total + sameCorpusScenarioScreenshotArchiveArtifacts(entry, sampleCount).length
  }, 0)
}

function sameCorpusProofArchiveArtifacts(
  cases: readonly SameCorpusProofArchiveSummaryCase[],
  sampleCount: number,
): SameCorpusProofArchiveArtifact[] {
  return cases.flatMap((entry) => [
    ...sameCorpusScenarioScreenshotArchiveArtifacts(entry, sampleCount),
    ...sameCorpusMutationTargetScreenshotArchiveArtifacts(entry, sampleCount),
    ...sameCorpusGoogleSheetsCommittedStateArchiveArtifacts(entry, sampleCount),
  ])
}

function sameCorpusScenarioScreenshotArchiveArtifacts(
  entry: SameCorpusProofArchiveSummaryCase,
  sampleCount: number,
): SameCorpusScenarioScreenshotArchiveArtifact[] {
  return requiredProofArchiveProducts.flatMap((product) => {
    const path = entry.scenarioProof.screenshotProof.artifactPaths.find((artifactPath) =>
      artifactPath.replaceAll('\\', '/').endsWith(`/${product}-sample-1.png`),
    )
    const proof = entry.scenarioProof.semanticUiProof.products.find((candidate) => candidate.product === product)
    const screenshotSha256 = proof?.screenshotSha256 ?? null
    const semanticProofAccepted =
      proof !== undefined &&
      validateSameCorpusProductSemanticUiProof(proof, { workload: entry.workload, sampleCount }).acceptedForCurrentScorecard
    return hasText(path) && isSha256(screenshotSha256) && semanticProofAccepted
      ? [
          {
            kind: 'scenario-screenshot',
            product,
            workload: entry.workload,
            path,
            screenshotSha256,
          },
        ]
      : []
  })
}

function sameCorpusMutationTargetScreenshotArtifactCount(cases: readonly SameCorpusProofArchiveSummaryCase[], sampleCount: number): number {
  return cases.reduce((total, entry) => {
    if (!uiSameCorpusWorkloadMutatesWorkbook(entry.workload)) {
      return total
    }
    return (
      total +
      requiredProofArchiveProducts.reduce((productTotal, product) => {
        const proof = entry.scenarioProof.semanticUiProof.products.find((candidate) => candidate.product === product)
        if (!proof) {
          return productTotal
        }
        return productTotal + sameCorpusTargetScreenshotArtifactCountForProduct(proof, entry.workload, sampleCount)
      }, 0)
    )
  }, 0)
}

function sameCorpusMutationTargetScreenshotArchiveArtifacts(
  entry: SameCorpusProofArchiveSummaryCase,
  sampleCount: number,
): SameCorpusMutationTargetScreenshotArchiveArtifact[] {
  if (!uiSameCorpusWorkloadMutatesWorkbook(entry.workload)) {
    return []
  }
  return requiredProofArchiveProducts.flatMap((product) => {
    const proof = entry.scenarioProof.semanticUiProof.products.find((candidate) => candidate.product === product)
    return proof ? sameCorpusTargetScreenshotArchiveArtifactsForProduct(proof, entry.workload, sampleCount) : []
  })
}

function sameCorpusGoogleSheetsCommittedStateArtifactCount(
  cases: readonly SameCorpusProofArchiveSummaryCase[],
  sampleCount: number,
): number {
  return cases.reduce((total, entry) => {
    if (!uiSameCorpusWorkloadMutatesWorkbook(entry.workload)) {
      return total
    }
    const proof = entry.scenarioProof.semanticUiProof.products.find((candidate) => candidate.product === 'google-sheets')
    if (!proof) {
      return total
    }
    return total + sameCorpusCommittedStateArtifactCountForProduct(proof, entry.workload, sampleCount)
  }, 0)
}

function sameCorpusGoogleSheetsCommittedStateArchiveArtifacts(
  entry: SameCorpusProofArchiveSummaryCase,
  sampleCount: number,
): SameCorpusGoogleSheetsCommittedStateArchiveArtifact[] {
  if (!uiSameCorpusWorkloadMutatesWorkbook(entry.workload)) {
    return []
  }
  const proof = entry.scenarioProof.semanticUiProof.products.find((candidate) => candidate.product === 'google-sheets')
  return proof ? sameCorpusCommittedStateArchiveArtifactsForProduct(proof, entry.workload, sampleCount) : []
}

function sameCorpusTargetScreenshotArtifactCountForProduct(
  proof: SameCorpusScenarioProof['semanticUiProof']['products'][number],
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
): number {
  if (!sameCorpusProductSemanticProofAccepted(proof, workload, sampleCount)) {
    return 0
  }
  let artifactCount = 0
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const sample = proof.mutationTargetProofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!sample?.targetScreenshots) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const screenshot = sample.targetScreenshots[phase]
      if (hasText(screenshot.screenshotPath) && isSha256(screenshot.screenshotSha256)) {
        artifactCount += 1
      }
    }
  }
  return artifactCount
}

function sameCorpusTargetScreenshotArchiveArtifactsForProduct(
  proof: SameCorpusScenarioProof['semanticUiProof']['products'][number],
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
): SameCorpusMutationTargetScreenshotArchiveArtifact[] {
  if (!sameCorpusProductSemanticProofAccepted(proof, workload, sampleCount)) {
    return []
  }
  const artifacts: SameCorpusMutationTargetScreenshotArchiveArtifact[] = []
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const sample = proof.mutationTargetProofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!sample?.targetScreenshots) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const screenshot = sample.targetScreenshots[phase]
      if (hasText(screenshot.screenshotPath) && isSha256(screenshot.screenshotSha256)) {
        artifacts.push({
          kind: 'mutation-target-screenshot',
          product: screenshot.product,
          workload: screenshot.workload,
          sampleIndex,
          phase,
          path: screenshot.screenshotPath,
          screenshotSha256: screenshot.screenshotSha256,
        })
      }
    }
  }
  return artifacts
}

function sameCorpusCommittedStateArtifactCountForProduct(
  proof: SameCorpusScenarioProof['semanticUiProof']['products'][number],
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
): number {
  if (!sameCorpusProductSemanticProofAccepted(proof, workload, sampleCount)) {
    return 0
  }
  let artifactCount = 0
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const sample = proof.mutationTargetProofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!sample?.committedStateProof) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const phaseProof = sample.committedStateProof[phase]
      if (
        hasText(phaseProof.artifactPath) &&
        isSha256(phaseProof.artifactSha256) &&
        hasText(phaseProof.exportUrl) &&
        phaseProof.workbookByteSize > 0 &&
        isSha256(phaseProof.workbookSha256)
      ) {
        artifactCount += 1
      }
    }
  }
  return artifactCount
}

function sameCorpusCommittedStateArchiveArtifactsForProduct(
  proof: SameCorpusScenarioProof['semanticUiProof']['products'][number],
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
): SameCorpusGoogleSheetsCommittedStateArchiveArtifact[] {
  if (!sameCorpusProductSemanticProofAccepted(proof, workload, sampleCount)) {
    return []
  }
  const artifacts: SameCorpusGoogleSheetsCommittedStateArchiveArtifact[] = []
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const sample = proof.mutationTargetProofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!sample?.committedStateProof) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const phaseProof = sample.committedStateProof[phase]
      if (
        hasText(phaseProof.artifactPath) &&
        isSha256(phaseProof.artifactSha256) &&
        hasText(phaseProof.exportUrl) &&
        phaseProof.workbookByteSize > 0 &&
        isSha256(phaseProof.workbookSha256)
      ) {
        artifacts.push({
          kind: 'google-sheets-committed-state-export',
          product: 'google-sheets',
          workload: phaseProof.workload,
          sampleIndex,
          phase,
          artifactPath: phaseProof.artifactPath,
          artifactSha256: phaseProof.artifactSha256,
          sheetName: phaseProof.sheetName,
          sheetId: phaseProof.sheetId,
          targetRange: phaseProof.targetRange,
          capturedAtMs: phaseProof.capturedAtMs,
          exportUrl: phaseProof.exportUrl,
          workbookByteSize: phaseProof.workbookByteSize,
          workbookSha256: phaseProof.workbookSha256,
          readback: phaseProof.readback,
          readbackSha256: sha256Hex(stableJsonBytes(phaseProof.readback)),
        })
      }
    }
  }
  return artifacts
}

function sameCorpusProductSemanticProofAccepted(
  proof: SameCorpusScenarioProof['semanticUiProof']['products'][number],
  workload: UiResponsivenessSameCorpusWorkload,
  sampleCount: number,
): boolean {
  return validateSameCorpusProductSemanticUiProof(proof, { workload, sampleCount }).acceptedForCurrentScorecard
}

function verifySameCorpusProofArchiveFile(
  artifact: SameCorpusProofArchiveArtifact,
  options: SameCorpusProofArchiveBuildOptions,
): SameCorpusProofArchiveFileVerificationEntry {
  const path = sameCorpusProofArchiveArtifactPath(artifact)
  const resolvedPath = resolveSameCorpusProofArchiveArtifactPath(path, options.artifactBaseDir)
  const expectedSha256 = sameCorpusProofArchiveArtifactSha256(artifact)
  const baseEntry = {
    kind: artifact.kind,
    product: artifact.product,
    workload: artifact.workload,
    ...('sampleIndex' in artifact ? { sampleIndex: artifact.sampleIndex, phase: artifact.phase } : {}),
    path,
    resolvedPath,
    expectedSha256,
  }
  if (!existsSync(resolvedPath)) {
    return { ...baseEntry, status: 'missing' }
  }
  const bytes = readFileSync(resolvedPath)
  const actualSha256 = sha256Hex(bytes)
  if (actualSha256 !== expectedSha256) {
    return { ...baseEntry, status: 'hash-mismatch', actualSha256 }
  }
  const identityMismatchReason =
    artifact.kind === 'google-sheets-committed-state-export' ? committedStateArtifactIdentityMismatchReason(artifact, bytes) : null
  return identityMismatchReason
    ? { ...baseEntry, status: 'identity-mismatch', actualSha256, identityMismatchReason }
    : { ...baseEntry, status: 'verified', actualSha256 }
}

function verifySameCorpusProofArchiveZipFile(
  artifact: SameCorpusProofArchiveArtifact,
  entriesByPath: ReadonlyMap<string, Uint8Array>,
  options: { readonly archivePath: string; readonly entryRootDir?: string },
): SameCorpusProofArchiveFileVerificationEntry {
  const path = sameCorpusProofArchiveArtifactPath(artifact)
  const entryPath = resolveSameCorpusProofArchiveZipArtifactEntry(path, options.entryRootDir)
  const expectedSha256 = sameCorpusProofArchiveArtifactSha256(artifact)
  const baseEntry = {
    kind: artifact.kind,
    product: artifact.product,
    workload: artifact.workload,
    ...('sampleIndex' in artifact ? { sampleIndex: artifact.sampleIndex, phase: artifact.phase } : {}),
    path,
    resolvedPath: `${options.archivePath}#${entryPath}`,
    expectedSha256,
  }
  const bytes = entriesByPath.get(entryPath)
  if (!bytes) {
    return { ...baseEntry, status: 'missing' }
  }
  const actualSha256 = sha256Hex(bytes)
  if (actualSha256 !== expectedSha256) {
    return { ...baseEntry, status: 'hash-mismatch', actualSha256 }
  }
  const identityMismatchReason =
    artifact.kind === 'google-sheets-committed-state-export' ? committedStateArtifactIdentityMismatchReason(artifact, bytes) : null
  return identityMismatchReason
    ? { ...baseEntry, status: 'identity-mismatch', actualSha256, identityMismatchReason }
    : { ...baseEntry, status: 'verified', actualSha256 }
}

function sameCorpusProofArchiveArtifactPath(artifact: SameCorpusProofArchiveArtifact): string {
  return artifact.kind === 'google-sheets-committed-state-export' ? artifact.artifactPath : artifact.path
}

function sameCorpusProofArchiveArtifactSha256(artifact: SameCorpusProofArchiveArtifact): string {
  return artifact.kind === 'google-sheets-committed-state-export' ? artifact.artifactSha256 : artifact.screenshotSha256
}

function resolveSameCorpusProofArchiveArtifactPath(path: string, artifactBaseDir: string | undefined): string {
  return isAbsolute(path) ? path : resolve(artifactBaseDir ?? process.cwd(), path)
}

function normalizeSameCorpusProofArchiveZipEntries(zipEntries: Record<string, Uint8Array>): Map<string, Uint8Array> {
  const entriesByPath = new Map<string, Uint8Array>()
  for (const [path, bytes] of Object.entries(zipEntries)) {
    const normalizedPath = normalizeSameCorpusProofArchiveZipEntryPath(path, 'ZIP entry')
    if (entriesByPath.has(normalizedPath)) {
      throw new Error(`UI responsiveness same-corpus proof archive ZIP has duplicate normalized entry: ${normalizedPath}`)
    }
    entriesByPath.set(normalizedPath, bytes)
  }
  return entriesByPath
}

function resolveSameCorpusProofArchiveZipManifestEntry(
  entriesByPath: ReadonlyMap<string, Uint8Array>,
  manifestEntryPath: string | undefined,
): string {
  if (manifestEntryPath) {
    return normalizeSameCorpusProofArchiveZipEntryPath(manifestEntryPath, 'manifest entry')
  }
  const candidates = [...entriesByPath.keys()].filter(
    (entryPath) => entryPath.endsWith('/proof-archive-manifest.json') || entryPath === 'proof-archive-manifest.json',
  )
  if (candidates.length === 1) {
    return candidates[0]
  }
  if (candidates.length === 0) {
    throw new Error('UI responsiveness same-corpus proof archive ZIP manifest is missing')
  }
  throw new Error('UI responsiveness same-corpus proof archive ZIP contains multiple proof manifests; pass manifestEntryPath')
}

function resolveSameCorpusProofArchiveZipArtifactEntry(path: string, entryRootDir: string | undefined): string {
  const normalizedPath = normalizeSameCorpusProofArchiveZipEntryPath(path, 'artifact entry')
  if (!entryRootDir) {
    return normalizedPath
  }
  return `${normalizeSameCorpusProofArchiveZipEntryPath(entryRootDir, 'entry root')}/${normalizedPath}`
}

function normalizeSameCorpusProofArchiveZipEntryPath(path: string, label: string): string {
  const normalizedPath = path
    .replaceAll('\\', '/')
    .replace(/^\.\/+/u, '')
    .trim()
  if (normalizedPath.length === 0) {
    throw new Error(`UI responsiveness same-corpus proof archive ZIP ${label} path is empty`)
  }
  if (normalizedPath.startsWith('/') || normalizedPath.split('/').includes('..')) {
    throw new Error(`UI responsiveness same-corpus proof archive ZIP ${label} escapes the archive: ${path}`)
  }
  return normalizedPath
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function committedStateArtifactIdentityMismatchReason(
  artifact: SameCorpusGoogleSheetsCommittedStateArchiveArtifact,
  bytes: Uint8Array,
): string | null {
  const record = parseJsonRecord(bytes, artifact.artifactPath)
  const expectedFields: ReadonlyArray<readonly [string, unknown]> = [
    ['product', artifact.product],
    ['workload', artifact.workload],
    ['sampleIndex', artifact.sampleIndex],
    ['phase', artifact.phase],
    ['sheetName', artifact.sheetName],
    ['sheetId', artifact.sheetId],
    ['targetRange', artifact.targetRange],
    ['capturedAtMs', artifact.capturedAtMs],
    ['artifactPath', artifact.artifactPath],
    ['exportUrl', artifact.exportUrl],
    ['workbookByteSize', artifact.workbookByteSize],
    ['workbookSha256', artifact.workbookSha256],
  ]
  const mismatchedField = expectedFields.find(([key, expected]) => !sameStableJsonValue(record[key], expected))?.[0]
  if (mismatchedField) {
    return `committed-state ${mismatchedField} does not match archive manifest`
  }
  const readbackHash = sha256Hex(stableJsonBytes(record.readback ?? null))
  if (readbackHash !== artifact.readbackSha256) {
    return 'committed-state readback does not match archive manifest'
  }
  return committedStateArtifactSemanticMismatchReason(artifact)
}

function committedStateArtifactSemanticMismatchReason(artifact: SameCorpusGoogleSheetsCommittedStateArchiveArtifact): string | null {
  if (artifact.workload !== 'fill-format-change') {
    return null
  }
  const expectedFillColor = sameCorpusFillColorExpectedColor(artifact.sampleIndex)
  const hasExpectedFill = sameCorpusFillColorsMatch(artifact.readback.fillColor, expectedFillColor)
  if (artifact.phase === 'after') {
    return hasExpectedFill ? null : 'committed-state after fill does not match intended swatch'
  }
  return hasExpectedFill ? `committed-state ${artifact.phase} fill still matches intended swatch` : null
}

function parseJsonRecord(bytes: Uint8Array, path: string): Record<string, unknown> {
  try {
    return asObject(JSON.parse(new TextDecoder().decode(bytes)) as unknown, `same-corpus committed-state artifact ${path}`)
  } catch (error: unknown) {
    throw new Error(`Unable to parse same-corpus committed-state artifact ${path}`, { cause: error })
  }
}

function sameStableJsonValue(left: unknown, right: unknown): boolean {
  return JSON.stringify(stableJsonValue(left)) === JSON.stringify(stableJsonValue(right))
}

function stableJsonBytes(value: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(stableJsonValue(value)))
}

function stableJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(stableJsonValue)
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entryValue]) => entryValue !== undefined)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entryValue]) => [key, stableJsonValue(entryValue)]),
    )
  }
  return value
}

function sampleIndexes(sampleCount: number): number[] {
  return Array.from({ length: Math.max(0, sampleCount) }, (_, sampleIndex) => sampleIndex)
}

function hasText(value: string | null | undefined): value is string {
  return value !== null && value !== undefined && value.trim().length > 0
}

function isSha256(value: string | null | undefined): boolean {
  return /^[a-f0-9]{64}$/u.test(value?.trim().toLowerCase() ?? '')
}
