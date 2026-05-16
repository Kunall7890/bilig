import { existsSync, readFileSync } from 'node:fs'

import { indexPublicWorkbookCorpusCases, publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export function selectManifestArtifactsWithRecordedCases(
  manifest: PublicWorkbookManifest,
  recordedCases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookManifest {
  return {
    ...manifest,
    artifacts: selectRecordedArtifactsInManifestOrder(manifest.artifacts, recordedCases),
  }
}

export function existingScorecardGeneratedAt(scorecardPath: string): string | undefined {
  if (!existsSync(scorecardPath)) {
    return undefined
  }
  const parsed: unknown = JSON.parse(readFileSync(scorecardPath, 'utf8'))
  if (typeof parsed !== 'object' || parsed === null) {
    return undefined
  }
  const generatedAt = Reflect.get(parsed, 'generatedAt')
  return typeof generatedAt === 'string' && generatedAt.trim().length > 0 ? generatedAt : undefined
}

export function selectRecordedArtifactsInManifestOrder(
  artifacts: readonly PublicWorkbookArtifact[],
  recordedCases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookArtifact[] {
  const casesById = indexPublicWorkbookCorpusCases(recordedCases)
  return artifacts.filter((artifact) => {
    const recordedCase = casesById.get(artifact.id)
    return recordedCase?.passed === true && publicWorkbookCorpusCaseMatchesArtifact(recordedCase, artifact)
  })
}

export function selectRecordedCasesInManifestOrder(
  artifacts: readonly PublicWorkbookArtifact[],
  recordedCases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookCorpusCase[] {
  const casesById = indexPublicWorkbookCorpusCases(recordedCases)
  return artifacts.flatMap((artifact) => {
    const recordedCase = casesById.get(artifact.id)
    return recordedCase?.passed === true && publicWorkbookCorpusCaseMatchesArtifact(recordedCase, artifact) ? [recordedCase] : []
  })
}
