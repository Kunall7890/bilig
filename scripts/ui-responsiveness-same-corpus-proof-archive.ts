import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'

import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import type { SameCorpusCapture, UiResponsivenessSameCorpusProduct } from './ui-responsiveness-same-corpus-scorecard-types.ts'
import { requiredUiResponsivenessSameCorpusMutationTargetProofWorkloads } from './ui-responsiveness-same-corpus-mutation-target-proof-summary.ts'
import {
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
  readonly complete: boolean
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
  readonly exportUrl: string
  readonly workbookByteSize: number
  readonly workbookSha256: string
}

type SameCorpusMutationTargetProof = SameCorpusScenarioProof['semanticUiProof']['products'][number]['mutationTargetProofs'][number]

const requiredProofArchiveProducts = ['bilig', 'google-sheets'] as const satisfies readonly UiResponsivenessSameCorpusProduct[]
const mutationTargetScreenshotPhaseCount = 3
const googleSheetsCommittedStatePhaseCount = 3

export function buildSameCorpusProofArchiveManifest(capture: SameCorpusCapture): SameCorpusProofArchiveManifest {
  const artifacts = sameCorpusProofArchiveArtifacts(capture.cases, capture.sampleCount)
  const requiredArtifactCount = capture.runManifest.requiredProofArchiveArtifactCount
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-proof-archive',
    captureRunSignature: capture.runManifest.captureRunSignature,
    requiredArtifactCount,
    artifactCount: artifacts.length,
    complete: artifacts.length === requiredArtifactCount,
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
    sameCorpusScenarioScreenshotArtifactCount(cases) +
    sameCorpusMutationTargetScreenshotArtifactCount(cases, sampleCount) +
    sameCorpusGoogleSheetsCommittedStateArtifactCount(cases, sampleCount)
  )
}

function sameCorpusScenarioScreenshotArtifactCount(cases: readonly SameCorpusProofArchiveSummaryCase[]): number {
  return cases.reduce((total, entry) => {
    return total + sameCorpusScenarioScreenshotArchiveArtifacts(entry).length
  }, 0)
}

function sameCorpusProofArchiveArtifacts(
  cases: readonly SameCorpusProofArchiveSummaryCase[],
  sampleCount: number,
): SameCorpusProofArchiveArtifact[] {
  return cases.flatMap((entry) => [
    ...sameCorpusScenarioScreenshotArchiveArtifacts(entry),
    ...sameCorpusMutationTargetScreenshotArchiveArtifacts(entry, sampleCount),
    ...sameCorpusGoogleSheetsCommittedStateArchiveArtifacts(entry, sampleCount),
  ])
}

function sameCorpusScenarioScreenshotArchiveArtifacts(
  entry: SameCorpusProofArchiveSummaryCase,
): SameCorpusScenarioScreenshotArchiveArtifact[] {
  return requiredProofArchiveProducts.flatMap((product) => {
    const path = entry.scenarioProof.screenshotProof.artifactPaths.find((artifactPath) =>
      artifactPath.replaceAll('\\', '/').endsWith(`/${product}-sample-1.png`),
    )
    return hasText(path)
      ? [
          {
            kind: 'scenario-screenshot',
            product,
            workload: entry.workload,
            path,
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
        return productTotal + sameCorpusTargetScreenshotArtifactCountForProduct(proof.mutationTargetProofs, sampleCount)
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
    return proof ? sameCorpusTargetScreenshotArchiveArtifactsForProduct(proof.mutationTargetProofs, sampleCount) : []
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
    return total + sameCorpusCommittedStateArtifactCountForProduct(proof.mutationTargetProofs, sampleCount)
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
  return proof ? sameCorpusCommittedStateArchiveArtifactsForProduct(proof.mutationTargetProofs, sampleCount) : []
}

function sameCorpusTargetScreenshotArtifactCountForProduct(proofs: readonly SameCorpusMutationTargetProof[], sampleCount: number): number {
  let artifactCount = 0
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const proof = proofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!proof?.targetScreenshots) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const screenshot = proof.targetScreenshots[phase]
      if (hasText(screenshot.screenshotPath) && isSha256(screenshot.screenshotSha256)) {
        artifactCount += 1
      }
    }
  }
  return artifactCount
}

function sameCorpusTargetScreenshotArchiveArtifactsForProduct(
  proofs: readonly SameCorpusMutationTargetProof[],
  sampleCount: number,
): SameCorpusMutationTargetScreenshotArchiveArtifact[] {
  const artifacts: SameCorpusMutationTargetScreenshotArchiveArtifact[] = []
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const proof = proofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!proof?.targetScreenshots) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const screenshot = proof.targetScreenshots[phase]
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

function sameCorpusCommittedStateArtifactCountForProduct(proofs: readonly SameCorpusMutationTargetProof[], sampleCount: number): number {
  let artifactCount = 0
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const proof = proofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!proof?.committedStateProof) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const phaseProof = proof.committedStateProof[phase]
      if (hasText(phaseProof.exportUrl) && phaseProof.workbookByteSize > 0 && isSha256(phaseProof.workbookSha256)) {
        artifactCount += 1
      }
    }
  }
  return artifactCount
}

function sameCorpusCommittedStateArchiveArtifactsForProduct(
  proofs: readonly SameCorpusMutationTargetProof[],
  sampleCount: number,
): SameCorpusGoogleSheetsCommittedStateArchiveArtifact[] {
  const artifacts: SameCorpusGoogleSheetsCommittedStateArchiveArtifact[] = []
  for (const sampleIndex of sampleIndexes(sampleCount)) {
    const proof = proofs.find((candidate) => candidate.sampleIndex === sampleIndex)
    if (!proof?.committedStateProof) {
      continue
    }
    for (const phase of ['before', 'after', 'restored'] as const) {
      const phaseProof = proof.committedStateProof[phase]
      if (hasText(phaseProof.exportUrl) && phaseProof.workbookByteSize > 0 && isSha256(phaseProof.workbookSha256)) {
        artifacts.push({
          kind: 'google-sheets-committed-state-export',
          product: 'google-sheets',
          workload: phaseProof.workload,
          sampleIndex,
          phase,
          exportUrl: phaseProof.exportUrl,
          workbookByteSize: phaseProof.workbookByteSize,
          workbookSha256: phaseProof.workbookSha256,
        })
      }
    }
  }
  return artifacts
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
