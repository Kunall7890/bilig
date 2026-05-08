import { describe, expect, it } from 'vitest'

import {
  listMissingPublicWorkbookArtifacts,
  listStalePublicWorkbookArtifacts,
  selectMissingPublicWorkbookArtifacts,
  selectStalePublicWorkbookArtifacts,
} from '../public-workbook-corpus-missing.ts'
import type {
  PublicWorkbookArtifact,
  PublicWorkbookCorpusCase,
  PublicWorkbookFeatureCounts,
  PublicWorkbookManifest,
  PublicWorkbookSource,
} from '../public-workbook-corpus-types.ts'

describe('public workbook corpus missing artifact selection', () => {
  it('selects the smallest missing cached artifacts first', () => {
    const small = artifact('workbook-small', 10)
    const medium = artifact('workbook-medium', 50)
    const large = artifact('workbook-large', 100)
    const manifest = manifestWithArtifacts([large, small, medium])

    expect(listMissingPublicWorkbookArtifacts({ manifest, cases: [] }).map((entry) => entry.id)).toEqual([
      'workbook-small',
      'workbook-medium',
      'workbook-large',
    ])
    expect(selectMissingPublicWorkbookArtifacts({ manifest, cases: [], limit: 2 }).map((entry) => entry.id)).toEqual([
      'workbook-small',
      'workbook-medium',
    ])
  })

  it('selects the smallest stale cached artifacts first', () => {
    const small = artifact('workbook-small', 10)
    const medium = artifact('workbook-medium', 50)
    const large = artifact('workbook-large', 100)
    const manifest = manifestWithArtifacts([large, small, medium])
    const cases = [staleCase(large), currentCase(small), staleCase(medium)]

    expect(listStalePublicWorkbookArtifacts({ manifest, cases }).map((entry) => entry.id)).toEqual(['workbook-medium', 'workbook-large'])
    expect(selectStalePublicWorkbookArtifacts({ manifest, cases, limit: 1 }).map((entry) => entry.id)).toEqual(['workbook-medium'])
  })
})

function manifestWithArtifacts(artifacts: readonly PublicWorkbookArtifact[]): PublicWorkbookManifest {
  return {
    schemaVersion: 1,
    corpus: 'public-workbook-corpus',
    targetWorkbookCount: artifacts.length,
    generatedAt: '2026-05-08T12:00:00.000Z',
    sources: artifacts.map(sourceForArtifact),
    artifacts,
  }
}

function artifact(id: string, byteSize: number): PublicWorkbookArtifact {
  return {
    id,
    sourceId: `source-${id}`,
    sourceUrl: `https://example.com/${id}.xlsx`,
    downloadUrl: `https://example.com/${id}.xlsx`,
    fileName: `${id}.xlsx`,
    cachePath: `files/${id}.xlsx`,
    sha256: id.padEnd(64, 'a').slice(0, 64),
    byteSize,
    workbookFingerprint: `${id}-fingerprint`,
    fetchedAt: '2026-05-08T12:00:00.000Z',
    license: {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    },
  }
}

function sourceForArtifact(entry: PublicWorkbookArtifact): PublicWorkbookSource {
  return {
    id: entry.sourceId,
    kind: 'direct-url',
    sourceUrl: entry.sourceUrl,
    downloadUrl: entry.downloadUrl,
    fileName: entry.fileName,
    discoveredAt: '2026-05-08T12:00:00.000Z',
    license: entry.license,
  }
}

function currentCase(entry: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    id: entry.id,
    sourceId: entry.sourceId,
    sourceUrl: entry.sourceUrl,
    fileName: entry.fileName,
    sha256: entry.sha256,
    byteSize: entry.byteSize,
    license: entry.license,
    status: 'passed',
    passed: true,
    featureCounts: featureCounts(),
    workbookMetadata: {
      workbookName: entry.id,
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: 1,
          columnCount: 1,
          nonEmptyCellCount: 1,
          usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
        },
      ],
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: [],
    evidence: [`source=${entry.sourceUrl}`, `license=${entry.license.title}`, `sha256=${entry.sha256}`],
  }
}

function staleCase(entry: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  const next = currentCase(entry)
  return {
    ...next,
    workbookMetadata: {
      ...next.workbookMetadata,
      dimensions: next.workbookMetadata.dimensions.map(({ usedRange: _usedRange, ...dimension }) => dimension),
    },
  }
}

function featureCounts(): PublicWorkbookFeatureCounts {
  return {
    sheetCount: 1,
    cellCount: 1,
    formulaCellCount: 0,
    valueCellCount: 1,
    definedNameCount: 0,
    tableCount: 0,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: 0,
    styleRangeCount: 0,
    conditionalFormatCount: 0,
    dataValidationCount: 0,
    macroPayloadCount: 0,
    warningCount: 0,
  }
}
