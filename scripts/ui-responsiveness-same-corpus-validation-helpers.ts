import type { NumericSummary } from '../packages/benchmarks/src/stats.js'
import {
  buildWorkbookBenchmarkCorpus,
  isWorkbookBenchmarkCorpusId,
  type WorkbookBenchmarkCorpusId,
} from '../packages/benchmarks/src/workbook-corpus.js'
import {
  buildSameCorpusFingerprint,
  sameCorpusFingerprintVersion,
  type SameCorpusCaptureCorpusFingerprint,
} from './ui-responsiveness-same-corpus-fingerprint.ts'
import type {
  SameCorpusCaptureCorpusVerification,
  UiResponsivenessSameCorpusProduct,
} from './ui-responsiveness-same-corpus-scorecard-proof.ts'

const expectedCorpusFingerprintCache = new Map<WorkbookBenchmarkCorpusId, SameCorpusCaptureCorpusFingerprint>()

export function isSha256Hex(value: string): boolean {
  return /^[a-f0-9]{64}$/u.test(value)
}

export function cloneSameCorpusVerification(verification: SameCorpusCaptureCorpusVerification): SameCorpusCaptureCorpusVerification {
  return {
    verified: verification.verified,
    method: verification.method,
    sheetName: verification.sheetName,
    materializedCells: verification.materializedCells,
    corpusFingerprint: { ...verification.corpusFingerprint, primaryViewport: { ...verification.corpusFingerprint.primaryViewport } },
    sourceWorkbookSha256: verification.sourceWorkbookSha256,
    checkedCells: verification.checkedCells.map((cell) => ({ ...cell })),
  }
}

export function validateSummary(summary: NumericSummary, label: string, minimumSampleCount: number): void {
  if (summary.samples.length < minimumSampleCount) {
    throw new Error(`UI responsiveness same-corpus scorecard has too few samples for ${label}`)
  }
  for (const value of [summary.min, summary.median, summary.p95, summary.max, summary.mean, ...summary.samples]) {
    if (!Number.isFinite(value) || value < 0) {
      throw new Error(`UI responsiveness same-corpus scorecard has invalid numeric summary for ${label}`)
    }
  }
}

export function validateSameCorpusCaptureVerification(
  verification: SameCorpusCaptureCorpusVerification,
  product: UiResponsivenessSameCorpusProduct,
  expectedMaterializedCells: number | null,
  expectedCorpusCaseId: string,
  caseId: string,
): void {
  if (!verification.verified) {
    throw new Error(`UI responsiveness same-corpus verification is not marked verified for ${caseId} ${product}`)
  }
  if (expectedMaterializedCells !== null && verification.materializedCells !== expectedMaterializedCells) {
    throw new Error(`UI responsiveness same-corpus verification materialized cell count mismatch for ${caseId} ${product}`)
  }
  validateSameCorpusCaptureCorpusFingerprint(
    verification.corpusFingerprint,
    expectedCorpusCaseId,
    expectedMaterializedCells,
    caseId,
    product,
  )
  if (verification.sourceWorkbookSha256 !== null && !isSha256Hex(verification.sourceWorkbookSha256)) {
    throw new Error(`UI responsiveness same-corpus verification source workbook fingerprint is invalid for ${caseId} ${product}`)
  }
  if (
    product === 'bilig' &&
    verification.sourceWorkbookSha256 !== null &&
    verification.sourceWorkbookSha256 !== verification.corpusFingerprint.snapshotSha256
  ) {
    throw new Error(`UI responsiveness same-corpus Bilig source workbook fingerprint is stale for ${caseId}`)
  }
  if (product === 'bilig' && verification.method !== 'bilig-benchmark-state') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'google-sheets' && verification.method !== 'google-sheets-xlsx-export') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (product === 'microsoft-excel-web' && verification.method !== 'microsoft-excel-web-source-xlsx') {
    throw new Error(`UI responsiveness same-corpus verification method mismatch for ${caseId} ${product}`)
  }
  if (verification.checkedCells.length < 3) {
    throw new Error(`UI responsiveness same-corpus verification must check at least 3 cells for ${caseId} ${product}`)
  }
  for (const cell of verification.checkedCells) {
    if (cell.address.trim().length === 0 || cell.expected !== cell.actual) {
      throw new Error(`UI responsiveness same-corpus verification cell mismatch for ${caseId} ${product}`)
    }
  }
}

function validateSameCorpusCaptureCorpusFingerprint(
  fingerprint: SameCorpusCaptureCorpusFingerprint,
  expectedCorpusCaseId: string,
  expectedMaterializedCells: number | null,
  caseId: string,
  product: UiResponsivenessSameCorpusProduct,
): void {
  if (fingerprint.version !== sameCorpusFingerprintVersion) {
    throw new Error(`UI responsiveness same-corpus verification fingerprint version is stale for ${caseId} ${product}`)
  }
  if (fingerprint.corpusCaseId !== expectedCorpusCaseId) {
    throw new Error(`UI responsiveness same-corpus verification corpus fingerprint mismatch for ${caseId} ${product}`)
  }
  if (expectedMaterializedCells !== null && fingerprint.materializedCells !== expectedMaterializedCells) {
    throw new Error(`UI responsiveness same-corpus verification fingerprint materialized cell count mismatch for ${caseId} ${product}`)
  }
  if (!isSha256Hex(fingerprint.snapshotSha256)) {
    throw new Error(`UI responsiveness same-corpus verification benchmark fingerprint is invalid for ${caseId} ${product}`)
  }
  if (!isWorkbookBenchmarkCorpusId(fingerprint.corpusCaseId)) {
    throw new Error(`UI responsiveness same-corpus verification uses unknown corpus fingerprint for ${caseId} ${product}`)
  }
  const expectedFingerprint = expectedCorpusFingerprint(fingerprint.corpusCaseId)
  if (JSON.stringify(fingerprint) !== JSON.stringify(expectedFingerprint)) {
    throw new Error(`UI responsiveness same-corpus verification benchmark fingerprint is stale for ${caseId} ${product}`)
  }
}

function expectedCorpusFingerprint(corpusId: WorkbookBenchmarkCorpusId): SameCorpusCaptureCorpusFingerprint {
  const cached = expectedCorpusFingerprintCache.get(corpusId)
  if (cached) {
    return cached
  }
  const fingerprint = buildSameCorpusFingerprint(buildWorkbookBenchmarkCorpus(corpusId)).corpusFingerprint
  expectedCorpusFingerprintCache.set(corpusId, fingerprint)
  return fingerprint
}
