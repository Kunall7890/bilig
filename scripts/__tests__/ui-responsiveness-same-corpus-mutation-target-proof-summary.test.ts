import { describe, expect, it } from 'vitest'

import {
  sameCorpusMutationTargetProofCaseCount,
  sameCorpusMutationTargetProofProductSummaries,
  sameCorpusMutationTargetProofSampleCount,
} from '../ui-responsiveness-same-corpus-mutation-target-proof-summary.ts'
import {
  sameCorpusMutationTargetProofProductEvidenceLines,
  sameCorpusMutationTargetProofProductGapLines,
} from '../ui-responsiveness-same-corpus-mutation-target-proof-gaps.ts'
import type { SameCorpusMutationTargetProof, SameCorpusScenarioProof } from '../ui-responsiveness-same-corpus-proof.ts'
import type { SameCorpusProductSemanticUiProof } from '../ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from '../ui-responsiveness-same-corpus-scorecard-proof.ts'

describe('same-corpus mutation target proof summary', () => {
  it('counts accepted per-product mutation samples even when the whole mutating case is incomplete', () => {
    const cases = [
      mutationTargetCase({
        biligSamples: [0, 1, 2],
        googleSamples: [],
      }),
    ]

    expect(sameCorpusMutationTargetProofCaseCount(cases, 3)).toBe(0)
    expect(sameCorpusMutationTargetProofSampleCount(cases, 3)).toBe(3)
    expect(sameCorpusMutationTargetProofProductSummaries(cases, 3)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          workload: 'edit-visible-cell',
          product: 'bilig',
          requiredSampleCount: 3,
          rawSampleCount: 3,
          acceptedSampleCount: 3,
          accepted: true,
          samples: expect.arrayContaining([
            expect.objectContaining({
              sampleIndex: 0,
              sheetName: 'WideGrid',
              sheetId: 'sheet-wide-grid',
              targetRange: 'A1',
              intendedOperation: 'edit-visible-cell',
              intendedPayload: {
                kind: 'cell-value',
                value: 'bilig-same-corpus-1',
              },
              before: expect.objectContaining({ source: 'bilig-authoritative-range', value: 'metric-1' }),
              after: expect.objectContaining({ capturedRevision: 'authoritative-readback-1', value: 'bilig-same-corpus-1' }),
              restored: expect.objectContaining({ value: 'metric-1' }),
              visibleAfter: expect.objectContaining({ value: 'bilig-same-corpus-1' }),
              visibleRestored: expect.objectContaining({ value: 'metric-1' }),
              authoritativeReadbackRevision: 'authoritative-readback-1',
              visibleRenderRevision: `bilig-visible-scene-sha256:${visibleSceneProofSha256(0)}`,
              screenshotSha256: 'b'.repeat(64),
              undoRestoreStatus: 'verified',
            }),
          ]),
        }),
        expect.objectContaining({
          workload: 'edit-visible-cell',
          product: 'google-sheets',
          requiredSampleCount: 3,
          rawSampleCount: 0,
          acceptedSampleCount: 0,
          accepted: false,
          samples: expect.arrayContaining([
            expect.objectContaining({
              sampleIndex: 0,
              present: false,
              sheetName: null,
              sheetId: null,
              targetRange: null,
              intendedOperation: null,
              before: null,
              after: null,
              restored: null,
              visibleAfter: null,
              visibleRestored: null,
              authoritativeReadbackRevision: null,
              visibleRenderRevision: null,
              undoRestoreStatus: null,
            }),
          ]),
          invalidReasons: expect.arrayContaining(['semantic UI mutation target proof for edit-visible-cell covers 0/3 samples']),
        }),
      ]),
    )
  })

  it('counts a mutating case only when both required products have claim-grade target proof for every sample', () => {
    const cases = [
      mutationTargetCase({
        biligSamples: [0, 1, 2],
        googleSamples: [0, 1, 2],
      }),
    ]

    expect(sameCorpusMutationTargetProofCaseCount(cases, 3)).toBe(1)
    expect(sameCorpusMutationTargetProofSampleCount(cases, 3)).toBe(6)
    expect(
      sameCorpusMutationTargetProofProductSummaries(cases, 3)
        .filter((entry) => entry.workload === 'edit-visible-cell')
        .map((entry) => ({
          accepted: entry.accepted,
          acceptedSampleCount: entry.acceptedSampleCount,
          product: entry.product,
        })),
    ).toEqual([
      { product: 'bilig', accepted: true, acceptedSampleCount: 3 },
      { product: 'google-sheets', accepted: true, acceptedSampleCount: 3 },
    ])
  })

  it('does not count short product proof as complete when the run manifest expects more samples', () => {
    const cases = [
      mutationTargetCase({
        biligSamples: [0],
        googleSamples: [0],
      }),
    ]

    expect(sameCorpusMutationTargetProofCaseCount(cases, 3)).toBe(0)
    expect(sameCorpusMutationTargetProofSampleCount(cases, 3)).toBe(2)
    expect(sameCorpusMutationTargetProofProductSummaries(cases, 3)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          workload: 'edit-visible-cell',
          product: 'bilig',
          requiredSampleCount: 3,
          rawSampleCount: 1,
          acceptedSampleCount: 1,
          accepted: false,
          samples: [
            expect.objectContaining({ sampleIndex: 0, present: true, accepted: true }),
            expect.objectContaining({ sampleIndex: 1, present: false, accepted: false }),
            expect.objectContaining({ sampleIndex: 2, present: false, accepted: false }),
          ],
          invalidReasons: expect.arrayContaining(['semantic UI mutation target proof for edit-visible-cell covers 1/3 samples']),
        }),
      ]),
    )
  })

  it('does not count stale target proof samples just because raw proof objects exist', () => {
    const cases = [
      mutationTargetCase({
        biligSamples: [0, 1, 2],
        googleSamples: [0, 1, 2],
        corruptGoogleProof: true,
      }),
    ]

    expect(sameCorpusMutationTargetProofCaseCount(cases, 3)).toBe(0)
    expect(sameCorpusMutationTargetProofSampleCount(cases, 3)).toBe(5)
    expect(sameCorpusMutationTargetProofProductSummaries(cases, 3)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          workload: 'edit-visible-cell',
          product: 'google-sheets',
          rawSampleCount: 3,
          acceptedSampleCount: 2,
          accepted: false,
          samples: [
            expect.objectContaining({
              sampleIndex: 0,
              accepted: false,
              invalidReasons: expect.arrayContaining([
                'semantic UI mutation target proof for edit-visible-cell did not prove the intended committed target value',
              ]),
            }),
            expect.objectContaining({ sampleIndex: 1, accepted: true, invalidReasons: [] }),
            expect.objectContaining({ sampleIndex: 2, accepted: true, invalidReasons: [] }),
          ],
          invalidReasons: expect.arrayContaining([
            'semantic UI mutation target proof for edit-visible-cell did not prove the intended committed target value',
          ]),
        }),
      ]),
    )
  })

  it('formats actionable product-level target proof gaps for the dominance audit', () => {
    const summaries = sameCorpusMutationTargetProofProductSummaries(
      [
        mutationTargetCase({
          biligSamples: [0],
          googleSamples: [],
        }),
      ],
      3,
    )

    expect(sameCorpusMutationTargetProofProductEvidenceLines(summaries)).toEqual(
      expect.arrayContaining([
        expect.stringContaining('edit-visible-cell/bilig accepted 1/3 samples (raw 1); missing samples: 1, 2; rejected samples: none'),
        expect.stringContaining(
          'edit-visible-cell/google-sheets accepted 0/3 samples (raw 0); missing samples: 0, 1, 2; rejected samples: none',
        ),
      ]),
    )
    expect(sameCorpusMutationTargetProofProductGapLines(summaries)).toEqual(
      expect.arrayContaining([
        expect.stringContaining('same-corpus mutation target proof gap: edit-visible-cell/bilig accepted 1/3 samples (raw 1)'),
        expect.stringContaining('same-corpus mutation target proof gap: edit-visible-cell/google-sheets accepted 0/3 samples (raw 0)'),
      ]),
    )
  })
})

function mutationTargetCase(args: {
  readonly biligSamples: readonly number[]
  readonly corruptGoogleProof?: boolean
  readonly googleSamples: readonly number[]
}): { readonly workload: 'edit-visible-cell'; readonly scenarioProof: SameCorpusScenarioProof } {
  return {
    workload: 'edit-visible-cell',
    scenarioProof: {
      biligMeanMs: 10,
      biligP95Ms: 11,
      googleMeanMs: 100,
      googleP95Ms: 110,
      meanRatio: 10,
      p95Ratio: 10,
      screenshotProof: {
        captured: true,
        requiredProducts: ['bilig', 'google-sheets'],
        artifactPaths: [],
        missingProducts: [],
      },
      pixelGridProof: {
        captured: true,
        requiredProducts: ['bilig', 'google-sheets'],
        products: [],
        productVerdicts: [],
        missingProducts: [],
      },
      semanticUiProof: {
        captured: false,
        requiredProducts: ['bilig', 'google-sheets'],
        products: [
          productSemanticProof('bilig', args.biligSamples),
          productSemanticProof('google-sheets', args.googleSamples, args.corruptGoogleProof ?? false),
        ],
        productVerdicts: [],
        missingProducts: ['google-sheets'],
      },
    },
  }
}

function productSemanticProof(
  product: UiResponsivenessSameCorpusProduct,
  sampleIndexes: readonly number[],
  corruptProof = false,
): SameCorpusProductSemanticUiProof {
  return {
    product,
    captured: true,
    method: product === 'bilig' ? 'bilig-visible-semantic-readback' : 'google-sheets-visible-semantic-readback',
    sheetName: 'WideGrid',
    sheetId: productSemanticSheetId(product),
    selectedRange: 'A1',
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
    ],
    authoritativeRenderRevision: product === 'bilig' ? 'authoritative-1' : null,
    visibleRenderRevision: product === 'bilig' ? 'visible-1' : null,
    screenshotSha256: 'a'.repeat(64),
    mutationTargetProofs: sampleIndexes.map((sampleIndex) => mutationTargetProof(product, sampleIndex, corruptProof && sampleIndex === 0)),
    evidence: ['semanticUiProofVersion=semantic-ui-v1', `sheetId=${productSemanticSheetId(product)}`],
  }
}

function productSemanticSheetId(product: UiResponsivenessSameCorpusProduct): string {
  if (product === 'bilig') {
    return 'sheet-wide-grid'
  }
  if (product === 'google-sheets') {
    return 'gid:160971404'
  }
  return 'excel-web-sheet-wide-grid'
}

function mutationTargetProof(
  product: UiResponsivenessSameCorpusProduct,
  sampleIndex: number,
  corruptProof: boolean,
): SameCorpusMutationTargetProof {
  const value = `${product}-same-corpus-${String(sampleIndex + 1)}`
  const afterValue = corruptProof ? 'stale-value' : value
  return {
    sampleIndex,
    workload: 'edit-visible-cell',
    intendedOperation: 'edit-visible-cell',
    intendedPayload: {
      kind: 'cell-value',
      value,
    },
    sheetName: 'WideGrid',
    sheetId: productSemanticSheetId(product),
    targetRange: 'A1',
    before: readback(product, 'metric-1', sampleIndex, 'before'),
    after: readback(product, afterValue, sampleIndex, 'after'),
    restored: readback(product, 'metric-1', sampleIndex, 'restored'),
    visibleAfter: {
      value: afterValue,
      formula: null,
      fillColor: null,
      visibleText: afterValue,
      source: 'visible-formula-bar',
    },
    visibleRestored: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'visible-formula-bar',
    },
    authoritativeReadbackRevision: authoritativeReadbackRevision(sampleIndex),
    visibleRenderRevision:
      product === 'bilig'
        ? `bilig-visible-scene-sha256:${visibleSceneProofSha256(sampleIndex)}`
        : `google-sheets-screenshot-sha256:${'b'.repeat(64)}`,
    screenshotSha256: 'b'.repeat(64),
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/${product}-sample-${String(
      sampleIndex + 1,
    )}-after.png`,
    undoRestoreStatus: 'verified',
  }
}

function readback(
  product: UiResponsivenessSameCorpusProduct,
  value: string,
  sampleIndex: number,
  phase: 'after' | 'before' | 'restored',
): SameCorpusMutationTargetProof['after'] {
  return {
    value,
    formula: null,
    fillColor: null,
    visibleText: value,
    source: product === 'bilig' ? 'bilig-authoritative-range' : 'visible-formula-bar',
    capturedRevision:
      product === 'bilig' ? (phase === 'after' ? authoritativeReadbackRevision(sampleIndex) : `${phase}-${sampleIndex}`) : null,
    visibleSceneProofSha256: product === 'bilig' && phase === 'after' ? visibleSceneProofSha256(sampleIndex) : null,
  }
}

function authoritativeReadbackRevision(sampleIndex: number): string {
  return `authoritative-readback-${String(sampleIndex + 1)}`
}

function visibleSceneProofSha256(sampleIndex: number): string {
  return String(sampleIndex + 1)
    .repeat(64)
    .slice(0, 64)
}
