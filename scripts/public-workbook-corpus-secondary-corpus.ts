import { existsSync, readFileSync } from 'node:fs'

import { isRecord, readNonNegativeInteger } from './public-workbook-corpus-completion-audit-helpers.ts'
import type { PublicWorkbookCorpusSecondaryFormulaCorpusStatus } from './public-workbook-corpus-completion-audit-types.ts'

export const hyperFormulaSecondaryCorpusArtifact = 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json'

export function readHyperFormulaSecondaryCorpus(path: string): PublicWorkbookCorpusSecondaryFormulaCorpusStatus {
  if (!existsSync(path)) {
    return missingHyperFormulaSecondaryCorpus()
  }
  try {
    const record = JSON.parse(readFileSync(path, 'utf8')) as unknown
    if (!isRecord(record)) {
      throw new Error('artifact root is not an object')
    }
    const scorecard = isRecord(record['scorecard']) ? record['scorecard'] : {}
    const results = Array.isArray(record['results']) ? record['results'] : []
    const comparableResults = results.filter((entry) => isRecord(entry) && entry['comparable'] === true)
    const comparableVerificationEquivalentCount = comparableResults.filter((entry) => {
      const comparison = isRecord(entry) ? entry['comparison'] : null
      return isRecord(comparison) && comparison['verificationEquivalent'] === true
    }).length
    const comparableCount = comparableResults.length || readNonNegativeInteger(scorecard, 'comparableCount', 0)
    return {
      artifact: hyperFormulaSecondaryCorpusArtifact,
      artifactPresent: true,
      suite: typeof record['suite'] === 'string' ? record['suite'] : null,
      resultCount: results.length,
      comparableCount,
      workpaperWins: readNonNegativeInteger(scorecard, 'workpaperWins', 0),
      hyperformulaWins: readNonNegativeInteger(scorecard, 'hyperformulaWins', 0),
      comparableVerificationEquivalentCount,
      allComparableVerificationEquivalent: comparableCount > 0 && comparableVerificationEquivalentCount === comparableCount,
      parseError: null,
    }
  } catch (error) {
    return {
      ...missingHyperFormulaSecondaryCorpus(),
      artifactPresent: true,
      parseError: error instanceof Error ? error.message : String(error),
    }
  }
}

export function missingHyperFormulaSecondaryCorpus(): PublicWorkbookCorpusSecondaryFormulaCorpusStatus {
  return {
    artifact: hyperFormulaSecondaryCorpusArtifact,
    artifactPresent: false,
    suite: null,
    resultCount: 0,
    comparableCount: 0,
    workpaperWins: 0,
    hyperformulaWins: 0,
    comparableVerificationEquivalentCount: 0,
    allComparableVerificationEquivalent: false,
    parseError: null,
  }
}
