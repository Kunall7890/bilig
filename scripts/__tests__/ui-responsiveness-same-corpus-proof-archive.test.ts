import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { buildWorkbookBenchmarkCorpus } from '../../packages/benchmarks/src/workbook-corpus.js'
import { buildSameCorpusCaptureArtifact } from '../capture-ui-responsiveness-same-corpus.ts'
import { buildSameCorpusFingerprint } from '../ui-responsiveness-same-corpus-fingerprint.ts'
import {
  buildSameCorpusProofArchiveManifest,
  proofArchiveManifestPath,
  type SameCorpusProofArchiveArtifact,
  type SameCorpusProofArchiveManifest,
  verifySameCorpusProofArchiveFiles,
  verifySameCorpusProofArchiveManifestPath,
  verifySameCorpusProofArchiveZipPath,
  writeSameCorpusProofArchiveManifest,
} from '../ui-responsiveness-same-corpus-proof-archive.ts'
import type { SameCorpusProductSemanticUiProof, SameCorpusScenarioProof } from '../ui-responsiveness-same-corpus-proof.ts'
import type {
  SameCorpusCaptureCase,
  SameCorpusCaptureMeasurement,
  SameCorpusOperationResponseProof,
  UiResponsivenessSameCorpusProduct,
} from '../ui-responsiveness-same-corpus-scorecard-types.ts'
import type { UiResponsivenessSameCorpusWorkload } from '../ui-responsiveness-same-corpus-workloads.ts'

const corpusFingerprint = buildSameCorpusFingerprint(buildWorkbookBenchmarkCorpus('wide-mixed-250k')).corpusFingerprint
const requiredProducts = ['bilig', 'google-sheets'] as const

describe('same-corpus proof archive manifest', () => {
  it('counts required archive artifacts in capture run manifests', () => {
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: [sameCorpusCaptureCase('open-workbook')],
    })

    expect(capture.runManifest).toMatchObject({
      requiredProofArchiveArtifactCount: 99,
      proofArchiveArtifactCount: 2,
    })
    expect(capture.runManifest.invalidReasons).toContain('proof archive covers 2/99 required proof artifacts')
  })

  it('builds and writes an auditable proof archive manifest beside capture artifacts', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-'))
    const outputPath = join(rootDir, 'capture.json')
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: [sameCorpusCaptureCase('open-workbook')],
    })

    const manifest = buildSameCorpusProofArchiveManifest(capture)
    const writtenManifest = writeSameCorpusProofArchiveManifest(capture, outputPath)
    const written = JSON.parse(readFileSync(proofArchiveManifestPath(outputPath), 'utf8'))

    expect(manifest).toMatchObject({
      schemaVersion: 1,
      suite: 'ui-responsiveness-same-corpus-proof-archive',
      captureRunSignature: capture.runManifest.captureRunSignature,
      requiredArtifactCount: 99,
      artifactCount: 2,
      complete: false,
      fileVerification: {
        checkedArtifactCount: 2,
        verifiedArtifactCount: 0,
        missingArtifactCount: 2,
        mismatchedArtifactCount: 0,
        complete: false,
      },
    })
    expect(manifest.artifacts).toContainEqual({
      kind: 'scenario-screenshot',
      product: 'bilig',
      workload: 'open-workbook',
      path: 'tmp/same-corpus-wide-mixed-250k-open-workbook/bilig-sample-1.png',
      screenshotSha256: 'a'.repeat(64),
    })
    expect(writtenManifest).toEqual(manifest)
    expect(written).toMatchObject({
      captureRunSignature: capture.runManifest.captureRunSignature,
      artifactCount: 2,
      complete: false,
      fileVerification: {
        checkedArtifactCount: 2,
        missingArtifactCount: 2,
        complete: false,
      },
    })
  })

  it('does not count path-only screenshots that lack accepted semantic screenshot proof', () => {
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: [sameCorpusCaptureCase('open-workbook', { includeSemanticScreenshotProof: false })],
    })

    expect(capture.runManifest).toMatchObject({
      requiredProofArchiveArtifactCount: 99,
      proofArchiveArtifactCount: 0,
    })
    expect(capture.runManifest.invalidReasons).toContain('proof archive covers 0/99 required proof artifacts')
    expect(buildSameCorpusProofArchiveManifest(capture)).toMatchObject({
      artifactCount: 0,
      complete: false,
      artifacts: [],
    })
  })

  it('verifies file-backed archive artifacts by reading bytes and matching sha256 digests', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-'))
    const scenarioBytes = 'scenario screenshot bytes'
    const committedReadback = {
      value: 'same-corpus-edit-1',
      formula: null,
      fillColor: null,
      visibleText: 'same-corpus-edit-1',
      source: 'google-sheets-xlsx-export' as const,
    }
    const committedBytes = `${JSON.stringify(
      {
        artifactPath: 'google-sheets-sample-1-after.json',
        capturedAtMs: 10,
        exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
        phase: 'after',
        product: 'google-sheets',
        readback: committedReadback,
        sampleIndex: 0,
        sheetId: 'gid:0',
        sheetName: 'WideGrid',
        targetRange: 'C5',
        workbookByteSize: 123,
        workbookSha256: 'c'.repeat(64),
        workload: 'edit-visible-cell',
      },
      null,
      2,
    )}\n`
    writeFileSync(join(rootDir, 'bilig-sample-1.png'), scenarioBytes)
    writeFileSync(join(rootDir, 'google-sheets-sample-1-after.json'), committedBytes)

    const verification = verifySameCorpusProofArchiveFiles(
      [
        {
          kind: 'scenario-screenshot',
          product: 'bilig',
          workload: 'open-workbook',
          path: 'bilig-sample-1.png',
          screenshotSha256: sha256Hex(scenarioBytes),
        },
        {
          kind: 'google-sheets-committed-state-export',
          product: 'google-sheets',
          workload: 'edit-visible-cell',
          sampleIndex: 0,
          phase: 'after',
          sheetName: 'WideGrid',
          sheetId: 'gid:0',
          targetRange: 'C5',
          capturedAtMs: 10,
          artifactPath: 'google-sheets-sample-1-after.json',
          artifactSha256: sha256Hex(committedBytes),
          exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
          workbookByteSize: 123,
          workbookSha256: 'c'.repeat(64),
          readback: committedReadback,
          readbackSha256: stableJsonSha256(committedReadback),
        },
      ],
      { artifactBaseDir: rootDir },
    )

    expect(verification).toMatchObject({
      checkedArtifactCount: 2,
      verifiedArtifactCount: 2,
      missingArtifactCount: 0,
      mismatchedArtifactCount: 0,
      complete: true,
    })
    expect(verification.entries.map((entry) => entry.status)).toEqual(['verified', 'verified'])
  })

  it('verifies final proof archive ZIP contents instead of trusting loose local files', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-zip-'))
    const archivePath = join(rootDir, 'same-corpus-proof.zip')
    const scenarioBytes = 'zip scenario screenshot bytes'
    const committedReadback = {
      value: 'same-corpus-edit-1',
      formula: null,
      fillColor: null,
      visibleText: 'same-corpus-edit-1',
      source: 'google-sheets-xlsx-export' as const,
    }
    const committedBytes = sameCorpusCommittedStateArtifactJson({
      artifactPath: 'google-sheets-sample-1-after.json',
      capturedAtMs: 10,
      exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
      phase: 'after',
      product: 'google-sheets',
      readback: committedReadback,
      sampleIndex: 0,
      sheetId: 'gid:0',
      sheetName: 'WideGrid',
      targetRange: 'C5',
      workbookByteSize: 123,
      workbookSha256: 'c'.repeat(64),
      workload: 'edit-visible-cell',
    })
    const artifacts: SameCorpusProofArchiveArtifact[] = [
      {
        kind: 'scenario-screenshot',
        product: 'bilig',
        workload: 'open-workbook',
        path: 'bilig-sample-1.png',
        screenshotSha256: sha256Hex(scenarioBytes),
      },
      {
        kind: 'google-sheets-committed-state-export',
        product: 'google-sheets',
        workload: 'edit-visible-cell',
        sampleIndex: 0,
        phase: 'after',
        sheetName: 'WideGrid',
        sheetId: 'gid:0',
        targetRange: 'C5',
        capturedAtMs: 10,
        artifactPath: 'google-sheets-sample-1-after.json',
        artifactSha256: sha256Hex(committedBytes),
        exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
        workbookByteSize: 123,
        workbookSha256: 'c'.repeat(64),
        readback: committedReadback,
        readbackSha256: stableJsonSha256(committedReadback),
      },
    ]
    writeSameCorpusProofZip(archivePath, {
      'proof-archive-manifest.json': sameCorpusProofArchiveManifestJson(artifacts),
      'bilig-sample-1.png': scenarioBytes,
      'google-sheets-sample-1-after.json': committedBytes,
    })

    const verification = verifySameCorpusProofArchiveZipPath(archivePath)

    expect(verification).toMatchObject({
      archivePath,
      manifestEntryPath: 'proof-archive-manifest.json',
      filesVerified: true,
      complete: true,
      fileVerification: {
        checkedArtifactCount: 2,
        verifiedArtifactCount: 2,
        missingArtifactCount: 0,
        mismatchedArtifactCount: 0,
        complete: true,
      },
    })
    expect(verification.fileVerification.entries.map((entry) => entry.status)).toEqual(['verified', 'verified'])
  })

  it('rejects stale final proof archive ZIP bytes even when a matching loose file exists', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-zip-stale-'))
    const archivePath = join(rootDir, 'same-corpus-proof.zip')
    const artifactPath = 'bilig-sample-1.png'
    const freshBytes = 'fresh screenshot bytes'
    const staleBytes = 'stale screenshot bytes'
    writeFileSync(join(rootDir, artifactPath), freshBytes)
    const artifacts: SameCorpusProofArchiveArtifact[] = [
      {
        kind: 'scenario-screenshot',
        product: 'bilig',
        workload: 'open-workbook',
        path: artifactPath,
        screenshotSha256: sha256Hex(freshBytes),
      },
    ]
    writeSameCorpusProofZip(archivePath, {
      'proof-archive-manifest.json': sameCorpusProofArchiveManifestJson(artifacts),
      [artifactPath]: staleBytes,
    })

    const verification = verifySameCorpusProofArchiveZipPath(archivePath)

    expect(verification).toMatchObject({
      filesVerified: false,
      complete: false,
      fileVerification: {
        checkedArtifactCount: 1,
        verifiedArtifactCount: 0,
        missingArtifactCount: 0,
        mismatchedArtifactCount: 1,
        complete: false,
      },
    })
    expect(verification.fileVerification.entries[0]).toMatchObject({
      status: 'hash-mismatch',
      actualSha256: sha256Hex(staleBytes),
      expectedSha256: sha256Hex(freshBytes),
    })
  })

  it('rejects committed-state archive files whose embedded target identity drifts from the manifest', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-'))
    const committedPayload = {
      artifactPath: 'google-sheets-sample-1-after.json',
      capturedAtMs: 10,
      exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
      phase: 'after',
      product: 'google-sheets',
      readback: {
        value: 'same-corpus-edit-1',
        formula: null,
        fillColor: null,
        visibleText: 'same-corpus-edit-1',
        source: 'google-sheets-xlsx-export',
      },
      sampleIndex: 0,
      sheetId: 'gid:0',
      sheetName: 'WideGrid',
      targetRange: 'D9',
      workbookByteSize: 123,
      workbookSha256: 'c'.repeat(64),
      workload: 'edit-visible-cell',
    }
    const committedBytes = `${JSON.stringify(committedPayload, null, 2)}\n`
    writeFileSync(join(rootDir, 'google-sheets-sample-1-after.json'), committedBytes)

    const verification = verifySameCorpusProofArchiveFiles(
      [
        {
          kind: 'google-sheets-committed-state-export',
          product: 'google-sheets',
          workload: 'edit-visible-cell',
          sampleIndex: 0,
          phase: 'after',
          sheetName: 'WideGrid',
          sheetId: 'gid:0',
          targetRange: 'C5',
          capturedAtMs: 10,
          artifactPath: 'google-sheets-sample-1-after.json',
          artifactSha256: sha256Hex(committedBytes),
          exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
          workbookByteSize: 123,
          workbookSha256: 'c'.repeat(64),
          readback: committedPayload.readback,
          readbackSha256: stableJsonSha256(committedPayload.readback),
        },
      ],
      { artifactBaseDir: rootDir },
    )

    expect(verification).toMatchObject({
      checkedArtifactCount: 1,
      verifiedArtifactCount: 0,
      mismatchedArtifactCount: 1,
      complete: false,
    })
    expect(verification.entries[0]).toMatchObject({
      status: 'identity-mismatch',
      identityMismatchReason: 'committed-state targetRange does not match archive manifest',
    })
  })

  it('rejects internally consistent Google Sheets fill artifacts that never commit the intended target fill', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-fill-'))
    const committedPayload = {
      artifactPath: 'google-sheets-sample-1-after.json',
      capturedAtMs: 10,
      exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
      phase: 'after' as const,
      product: 'google-sheets' as const,
      readback: {
        value: 'segment-5',
        formula: null,
        fillColor: null,
        visibleText: 'segment-5',
        source: 'google-sheets-xlsx-export' as const,
      },
      sampleIndex: 0,
      sheetId: 'gid:0',
      sheetName: 'WideGrid',
      targetRange: 'B5',
      workbookByteSize: 123,
      workbookSha256: 'c'.repeat(64),
      workload: 'fill-format-change' as const,
    }
    const committedBytes = sameCorpusCommittedStateArtifactJson(committedPayload)
    writeFileSync(join(rootDir, 'google-sheets-sample-1-after.json'), committedBytes)

    const verification = verifySameCorpusProofArchiveFiles(
      [
        {
          kind: 'google-sheets-committed-state-export',
          product: 'google-sheets',
          workload: 'fill-format-change',
          sampleIndex: 0,
          phase: 'after',
          sheetName: 'WideGrid',
          sheetId: 'gid:0',
          targetRange: 'B5',
          capturedAtMs: 10,
          artifactPath: 'google-sheets-sample-1-after.json',
          artifactSha256: sha256Hex(committedBytes),
          exportUrl: 'https://docs.google.com/spreadsheets/d/example/export?format=xlsx',
          workbookByteSize: 123,
          workbookSha256: 'c'.repeat(64),
          readback: committedPayload.readback,
          readbackSha256: stableJsonSha256(committedPayload.readback),
        },
      ],
      { artifactBaseDir: rootDir },
    )

    expect(verification).toMatchObject({
      checkedArtifactCount: 1,
      verifiedArtifactCount: 0,
      mismatchedArtifactCount: 1,
      complete: false,
    })
    expect(verification.entries[0]).toMatchObject({
      status: 'identity-mismatch',
      identityMismatchReason: 'committed-state after fill does not match intended swatch',
    })
  })

  it('rejects missing and mismatched proof archive files', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-'))
    writeFileSync(join(rootDir, 'stale.png'), 'stale bytes')

    const verification = verifySameCorpusProofArchiveFiles(
      [
        {
          kind: 'scenario-screenshot',
          product: 'bilig',
          workload: 'open-workbook',
          path: 'missing.png',
          screenshotSha256: sha256Hex('missing bytes'),
        },
        {
          kind: 'scenario-screenshot',
          product: 'google-sheets',
          workload: 'open-workbook',
          path: 'stale.png',
          screenshotSha256: sha256Hex('fresh bytes'),
        },
      ],
      { artifactBaseDir: rootDir },
    )

    expect(verification).toMatchObject({
      checkedArtifactCount: 2,
      verifiedArtifactCount: 0,
      missingArtifactCount: 1,
      mismatchedArtifactCount: 1,
      complete: false,
    })
    expect(verification.entries.map((entry) => entry.status)).toEqual(['missing', 'hash-mismatch'])
  })

  it('recomputes manifest completeness from file verification when reading a written manifest', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-proof-archive-'))
    const outputPath = join(rootDir, 'capture.json')
    const capture = buildSameCorpusCaptureArtifact({
      sampleCount: 3,
      limitations: ['test limitation'],
      cases: [sameCorpusCaptureCase('open-workbook')],
    })
    writeSameCorpusProofArchiveManifest(capture, outputPath)

    const verified = verifySameCorpusProofArchiveManifestPath(proofArchiveManifestPath(outputPath))

    expect(verified).toMatchObject({
      artifactCount: 2,
      complete: false,
      fileVerification: {
        checkedArtifactCount: 2,
        missingArtifactCount: 2,
        complete: false,
      },
    })
  })
})

function sameCorpusCaptureCase(
  workload: UiResponsivenessSameCorpusWorkload,
  options: { readonly includeSemanticScreenshotProof?: boolean } = {},
): SameCorpusCaptureCase {
  const scenarioProof = sameCorpusScenarioProof(workload, options)
  return {
    id: `same-corpus-wide-mixed-250k-${workload}`,
    corpusCaseId: 'wide-mixed-250k',
    materializedCells: 250_000,
    workload,
    biligMeanMs: scenarioProof.biligMeanMs,
    biligP95Ms: scenarioProof.biligP95Ms,
    googleMeanMs: scenarioProof.googleMeanMs,
    googleP95Ms: scenarioProof.googleP95Ms,
    meanRatio: scenarioProof.meanRatio,
    p95Ratio: scenarioProof.p95Ratio,
    screenshotProof: scenarioProof.screenshotProof,
    pixelGridProof: scenarioProof.pixelGridProof,
    semanticUiProof: scenarioProof.semanticUiProof,
    scenarioProof,
    bilig: sameCorpusMeasurement('bilig', 'load-to-ready'),
    googleSheets: sameCorpusMeasurement('google-sheets', 'load-to-ready'),
  }
}

function sameCorpusScenarioProof(
  workload: UiResponsivenessSameCorpusWorkload,
  options: { readonly includeSemanticScreenshotProof?: boolean } = {},
): SameCorpusScenarioProof {
  const artifactPaths = requiredProducts.map((product) => `tmp/same-corpus-wide-mixed-250k-${workload}/${product}-sample-1.png`)
  const includeSemanticScreenshotProof = options.includeSemanticScreenshotProof ?? true
  return {
    biligMeanMs: 10,
    biligP95Ms: 12,
    googleMeanMs: 20,
    googleP95Ms: 24,
    meanRatio: 2,
    p95Ratio: 2,
    screenshotProof: {
      captured: true,
      requiredProducts,
      artifactPaths,
      missingProducts: [],
    },
    pixelGridProof: {
      captured: false,
      requiredProducts,
      products: [],
      productVerdicts: [],
      missingProducts: [...requiredProducts],
    },
    semanticUiProof: {
      captured: includeSemanticScreenshotProof,
      requiredProducts,
      products: includeSemanticScreenshotProof ? requiredProducts.map((product) => sameCorpusSemanticUiProof(product)) : [],
      productVerdicts: [],
      missingProducts: includeSemanticScreenshotProof ? [] : [...requiredProducts],
    },
  }
}

function sameCorpusSemanticUiProof(product: UiResponsivenessSameCorpusProduct): SameCorpusProductSemanticUiProof {
  return {
    product,
    captured: true,
    method: product === 'bilig' ? 'bilig-visible-semantic-readback' : 'google-sheets-visible-semantic-readback',
    sheetName: 'WideGrid',
    sheetId: product === 'bilig' ? 'sheet-wide-grid' : 'gid:0',
    selectedRange: 'A1',
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
    ],
    authoritativeRenderRevision: product === 'bilig' ? 'authoritative-1' : null,
    visibleRenderRevision: product === 'bilig' ? 'visible-1' : null,
    screenshotSha256: product === 'bilig' ? 'a'.repeat(64) : 'b'.repeat(64),
    mutationTargetProofs: [],
    evidence: ['semanticUiProofVersion=semantic-ui-v1'],
  }
}

function sameCorpusMeasurement(
  product: UiResponsivenessSameCorpusProduct,
  operationResponseProof: SameCorpusOperationResponseProof,
): SameCorpusCaptureMeasurement {
  return {
    product,
    source: product === 'bilig' ? 'http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k' : 'https://example.com/sheet',
    operationResponseMsSamples: [10, 11, 12],
    operationResponseProofs: [operationResponseProof, operationResponseProof, operationResponseProof],
    ...(product === 'bilig' ? { authoritativeRenderProofMsSamples: [15, 16, 17] } : {}),
    postOperationFrameMsSamples: [8, 9, 10],
    corpusVerification: {
      verified: true,
      method: product === 'bilig' ? 'bilig-benchmark-state' : 'google-sheets-xlsx-export',
      sheetName: 'WideGrid',
      materializedCells: 250_000,
      corpusFingerprint,
      sourceWorkbookSha256: product === 'bilig' ? 'a'.repeat(64) : 'b'.repeat(64),
      checkedCells: [
        { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
        { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
        { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
      ],
    },
    limitations: [],
  }
}

function writeSameCorpusProofZip(archivePath: string, entries: Record<string, string>): void {
  writeFileSync(
    archivePath,
    zipSync(Object.fromEntries(Object.entries(entries).map(([entryPath, contents]) => [entryPath, strToU8(contents)]))),
  )
}

function sameCorpusProofArchiveManifestJson(artifacts: readonly SameCorpusProofArchiveArtifact[]): string {
  const manifest: SameCorpusProofArchiveManifest = {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-proof-archive',
    captureRunSignature: 'd'.repeat(64),
    requiredArtifactCount: artifacts.length,
    artifactCount: artifacts.length,
    filesVerified: true,
    complete: true,
    fileVerification: {
      schemaVersion: 1,
      checkedArtifactCount: artifacts.length,
      verifiedArtifactCount: artifacts.length,
      missingArtifactCount: 0,
      mismatchedArtifactCount: 0,
      complete: true,
      entries: [],
    },
    artifacts,
  }
  return `${JSON.stringify(manifest, null, 2)}\n`
}

function sameCorpusCommittedStateArtifactJson(value: {
  readonly artifactPath: string
  readonly capturedAtMs: number
  readonly exportUrl: string
  readonly phase: 'before' | 'after' | 'restored'
  readonly product: 'google-sheets'
  readonly readback: unknown
  readonly sampleIndex: number
  readonly sheetId: string | null
  readonly sheetName: string
  readonly targetRange: string
  readonly workbookByteSize: number
  readonly workbookSha256: string
  readonly workload: UiResponsivenessSameCorpusWorkload
}): string {
  return `${JSON.stringify(value, null, 2)}\n`
}

function sha256Hex(value: string): string {
  return createHash('sha256').update(value).digest('hex')
}

function stableJsonSha256(value: unknown): string {
  return createHash('sha256')
    .update(JSON.stringify(stableJsonValue(value)))
    .digest('hex')
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
