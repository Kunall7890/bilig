import type { NumericSummary } from '../packages/benchmarks/src/stats.js'
import type { SameCorpusCaptureCorpusVerification } from './ui-responsiveness-same-corpus-scorecard-proof.ts'

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
