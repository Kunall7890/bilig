import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { buildWorkbookBenchmarkCorpus } from '../../packages/benchmarks/src/workbook-corpus.js'
import { buildSameCorpusCaptureArtifact } from '../capture-ui-responsiveness-same-corpus.ts'
import { buildSameCorpusFingerprint } from '../ui-responsiveness-same-corpus-fingerprint.ts'
import {
  buildSameCorpusProofArchiveManifest,
  proofArchiveManifestPath,
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
