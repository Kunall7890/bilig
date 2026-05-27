import type { NumericSummary } from '../packages/benchmarks/src/stats.js'
import { sameCorpusSampleCount } from './ui-responsiveness-same-corpus-guardrails.ts'

export function validateOrderedSameCorpusTimingSamples(
  orderedSamples: readonly number[] | undefined,
  summary: NumericSummary | undefined,
  label: string,
): void {
  if (!orderedSamples || !summary) {
    return
  }
  if (orderedSamples.length < sameCorpusSampleCount) {
    throw new Error(`UI responsiveness same-corpus proof has too few ordered ${label} samples`)
  }
  if (orderedSamples.some((value) => !Number.isFinite(value) || value < 0)) {
    throw new Error(`UI responsiveness same-corpus proof has invalid ordered ${label} samples`)
  }
  const sortedSamples = [...orderedSamples].toSorted((left, right) => left - right)
  if (sortedSamples.length !== summary.samples.length || sortedSamples.some((value, index) => value !== summary.samples[index])) {
    throw new Error(`UI responsiveness same-corpus ordered ${label} samples do not match summary`)
  }
}
