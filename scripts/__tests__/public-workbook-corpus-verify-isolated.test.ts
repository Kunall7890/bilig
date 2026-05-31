import { describe, expect, it } from 'vitest'

import { borrowXlsxZipByteSource as borrowLargeSimpleVerifierXlsxZipByteSource } from '../public-workbook-corpus-large-simple-compact.ts'
import {
  buildResourceLimitedFootprintVerificationCase,
  buildVerificationWorkerProcessArgs,
  disableBunSmolVerificationWorkerEnvVar,
  shouldUseBunSmolForVerificationWorker,
} from '../public-workbook-corpus-verify-isolated.ts'
import type { PublicWorkbookArtifact } from '../public-workbook-corpus-types.ts'
import { emptyFeatureCounts, type WorkbookFootprint } from '../public-workbook-corpus-workbook.ts'

describe('public workbook corpus isolated verification worker runtime', () => {
  it('uses Bun smol mode for memory-sensitive isolated verification workers', () => {
    expect(shouldUseBunSmolForVerificationWorker({ versions: { bun: '1.3.0' }, env: {} })).toBe(true)
    expect(buildVerificationWorkerProcessArgs(['worker.ts', 'verify-artifact-worker'], { versions: { bun: '1.3.0' }, env: {} })).toEqual([
      '--smol',
      'worker.ts',
      'verify-artifact-worker',
    ])
  })

  it('does not add Bun runtime flags under Node or when explicitly disabled', () => {
    expect(shouldUseBunSmolForVerificationWorker({ versions: { node: '24.0.0' }, env: {} })).toBe(false)
    expect(
      buildVerificationWorkerProcessArgs(['worker.ts'], {
        versions: { bun: '1.3.0' },
        env: { [disableBunSmolVerificationWorkerEnvVar]: 'true' },
      }),
    ).toEqual(['worker.ts'])
  })

  it('keeps readRangeInto support on compact verifier borrowed byte sources', () => {
    const source = new InstrumentedByteSource(new Uint8Array([1, 2, 3, 4]))
    const borrowed = borrowLargeSimpleVerifierXlsxZipByteSource(source)
    const scratch = new Uint8Array(2)

    expect(Array.from(borrowed.readRangeInto?.(1, 3, scratch) ?? [])).toEqual([2, 3])

    expect(source.readIntoCount).toBe(1)
    expect(source.rangeCount).toBe(0)
  })

  it('classifies large-simple resource-limited footprints before starting the heavy verifier worker', () => {
    const artifact = publicWorkbookArtifact()
    const corpusCase = buildResourceLimitedFootprintVerificationCase({
      artifact,
      footprint: workbookFootprint(
        {
          cellCount: 342_986,
          valueCellCount: 296_781,
          formulaCellCount: 46_205,
        },
        { largeSimpleXlsxImport: { eligible: true, blockers: [] } },
      ),
      baseEvidence: [`source=${artifact.sourceUrl}`],
      runStructuralSmoke: false,
      maxCellCount: 1_500_000,
    })

    expect(corpusCase).toMatchObject({
      status: 'unsupported',
      passed: true,
      featureCounts: {
        cellCount: 342_986,
        formulaCellCount: 46_205,
      },
      validation: {
        importPassed: false,
        formulaOraclePassed: true,
        roundTripPassed: true,
      },
    })
    expect(corpusCase?.unsupportedFeatureClassifications).toEqual([
      'xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>2000formulas',
      'xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells',
    ])
    expect(corpusCase?.evidence).toEqual(
      expect.arrayContaining([
        'resource-limit-classifier=2026-05-17-native-streaming-xlsx-footprint',
        'formula-oracle-formula-count=46205',
        'rss-limit-phase=round-trip',
      ]),
    )
  })
})

class InstrumentedByteSource {
  readonly byteLength: number
  rangeCount = 0
  readIntoCount = 0

  constructor(private readonly bytes: Uint8Array) {
    this.byteLength = bytes.byteLength
  }

  readRange(start: number, end: number): Uint8Array {
    this.rangeCount += 1
    return this.bytes.subarray(start, end)
  }

  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array {
    this.readIntoCount += 1
    target.set(this.bytes.subarray(start, end), 0)
    return target.subarray(0, end - start)
  }
}

function publicWorkbookArtifact(): PublicWorkbookArtifact {
  return {
    id: 'workbook-5db97e9230dbaf6b',
    sourceId: 'source',
    sourceUrl: 'https://example.com/noibyfarmsize_fr.xlsx',
    downloadUrl: 'https://example.com/noibyfarmsize_fr.xlsx',
    fileName: 'noibyfarmsize_fr.xlsx',
    sha256: '0'.repeat(64),
    byteSize: 2_000_000,
    cachePath: 'noibyfarmsize_fr.xlsx',
    workbookFingerprint: '1'.repeat(64),
    fetchedAt: '2026-05-17T00:00:00.000Z',
    license: {
      title: 'Test',
      evidenceUrl: 'https://example.com/license',
      spdxId: 'CC0-1.0',
    },
  }
}

function workbookFootprint(
  counts: Partial<WorkbookFootprint['featureCounts']>,
  options: Pick<WorkbookFootprint, 'largeSimpleXlsxImport'>,
): WorkbookFootprint {
  const featureCounts = {
    ...emptyFeatureCounts(),
    sheetCount: 10,
    ...counts,
  }
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: 'noibyfarmsize_fr.xlsx',
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: featureCounts.cellCount,
          columnCount: 1,
          nonEmptyCellCount: featureCounts.cellCount,
          usedRange: { startRow: 0, startColumn: 0, endRow: featureCounts.cellCount - 1, endColumn: 0 },
        },
      ],
    },
    externalWorkbookReferences: [],
    ...options,
  }
}
