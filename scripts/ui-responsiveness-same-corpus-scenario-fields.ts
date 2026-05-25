import type { SameCorpusScenarioProof } from './ui-responsiveness-same-corpus-proof.ts'
import type { SameCorpusScenarioCaseFields } from './ui-responsiveness-same-corpus-scorecard-types.ts'

export function sameCorpusScenarioCaseFields(proof: SameCorpusScenarioProof): SameCorpusScenarioCaseFields {
  return {
    biligMeanMs: proof.biligMeanMs,
    biligP95Ms: proof.biligP95Ms,
    googleMeanMs: proof.googleMeanMs,
    googleP95Ms: proof.googleP95Ms,
    ...(proof.microsoftExcelWebMeanMs !== undefined ? { microsoftExcelWebMeanMs: proof.microsoftExcelWebMeanMs } : {}),
    ...(proof.microsoftExcelWebP95Ms !== undefined ? { microsoftExcelWebP95Ms: proof.microsoftExcelWebP95Ms } : {}),
    meanRatio: proof.meanRatio,
    p95Ratio: proof.p95Ratio,
    ...(proof.microsoftExcelWebMeanRatio !== undefined ? { microsoftExcelWebMeanRatio: proof.microsoftExcelWebMeanRatio } : {}),
    ...(proof.microsoftExcelWebP95Ratio !== undefined ? { microsoftExcelWebP95Ratio: proof.microsoftExcelWebP95Ratio } : {}),
    screenshotProof: proof.screenshotProof,
    pixelGridProof: proof.pixelGridProof,
  }
}

export function sameCorpusScenarioSummaryFieldsCurrent(
  entry: SameCorpusScenarioCaseFields & { readonly scenarioProof: SameCorpusScenarioProof },
): boolean {
  return stableJsonString(sameCorpusScenarioFieldsFromCase(entry)) === stableJsonString(sameCorpusScenarioCaseFields(entry.scenarioProof))
}

export function validateSameCorpusScenarioCaseFields(
  entry: SameCorpusScenarioCaseFields & { readonly id: string; readonly scenarioProof: SameCorpusScenarioProof },
  label: 'capture' | 'scorecard',
): void {
  if (!sameCorpusScenarioSummaryFieldsCurrent(entry)) {
    throw new Error(`UI responsiveness same-corpus ${label} scenario summary fields are stale: ${entry.id}`)
  }
}

function sameCorpusScenarioFieldsFromCase(entry: SameCorpusScenarioCaseFields): SameCorpusScenarioCaseFields {
  return {
    biligMeanMs: entry.biligMeanMs,
    biligP95Ms: entry.biligP95Ms,
    googleMeanMs: entry.googleMeanMs,
    googleP95Ms: entry.googleP95Ms,
    ...(entry.microsoftExcelWebMeanMs !== undefined ? { microsoftExcelWebMeanMs: entry.microsoftExcelWebMeanMs } : {}),
    ...(entry.microsoftExcelWebP95Ms !== undefined ? { microsoftExcelWebP95Ms: entry.microsoftExcelWebP95Ms } : {}),
    meanRatio: entry.meanRatio,
    p95Ratio: entry.p95Ratio,
    ...(entry.microsoftExcelWebMeanRatio !== undefined ? { microsoftExcelWebMeanRatio: entry.microsoftExcelWebMeanRatio } : {}),
    ...(entry.microsoftExcelWebP95Ratio !== undefined ? { microsoftExcelWebP95Ratio: entry.microsoftExcelWebP95Ratio } : {}),
    screenshotProof: entry.screenshotProof,
    pixelGridProof: entry.pixelGridProof,
  }
}

function stableJsonString(value: unknown): string {
  return JSON.stringify(stableJsonValue(value))
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
