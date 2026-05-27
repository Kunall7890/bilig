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
import { sameCorpusMutationTargetProofSignature } from '../ui-responsiveness-same-corpus-mutation-target-signature.ts'
import type { SameCorpusMutationTargetProof, SameCorpusScenarioProof } from '../ui-responsiveness-same-corpus-proof.ts'
import { sameCorpusMutationTargetRangeForSample } from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { SameCorpusProductSemanticUiProof } from '../ui-responsiveness-same-corpus-semantic-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from '../ui-responsiveness-same-corpus-scorecard-proof.ts'
import { sameCorpusEditVisibleCellValue } from '../ui-responsiveness-same-corpus-workload-runner.ts'

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
              product: 'bilig',
              committedTargetProofMs: 40,
              visibleTargetRenderMs: 12,
              committedStateValidationMs: 28,
              restoreValidationMs: 80,
              operationStartedAtMs: 1000,
              visibleTargetRenderCapturedAtMs: 1012,
              postMutationProofCapturedAtMs: 1040,
              restoreProofCapturedAtMs: 1120,
              sheetName: 'WideGrid',
              sheetId: 'sheet-wide-grid',
              targetRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0),
              intendedOperation: 'edit-visible-cell',
              intendedPayload: {
                kind: 'cell-value',
                value: sameCorpusEditVisibleCellValue(0),
              },
              before: expect.objectContaining({ source: 'bilig-authoritative-range', value: 'metric-1' }),
              after: expect.objectContaining({ capturedRevision: 'authoritative-readback-1', value: sameCorpusEditVisibleCellValue(0) }),
              restored: expect.objectContaining({ value: 'metric-1' }),
              visibleAfter: expect.objectContaining({ value: sameCorpusEditVisibleCellValue(0) }),
              visibleAfterSelectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0),
              visibleRestored: expect.objectContaining({ value: 'metric-1' }),
              visibleRestoredSelectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0),
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
              product: null,
              committedTargetProofMs: null,
              visibleTargetRenderMs: null,
              committedStateValidationMs: null,
              restoreValidationMs: null,
              operationStartedAtMs: null,
              visibleTargetRenderCapturedAtMs: null,
              postMutationProofCapturedAtMs: null,
              restoreProofCapturedAtMs: null,
              sheetName: null,
              sheetId: null,
              targetRange: null,
              intendedOperation: null,
              before: null,
              after: null,
              restored: null,
              visibleAfter: null,
              visibleAfterSelectedRange: null,
              visibleRestored: null,
              visibleRestoredSelectedRange: null,
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

  it('does not count screenshot-only target chrome changes as committed target proof', () => {
    const cases = [
      mutationTargetCase({
        biligSamples: [],
        googleProofMode: 'screenshot-only',
        googleSamples: [0],
      }),
    ]

    expect(sameCorpusMutationTargetProofSampleCount(cases, 1)).toBe(0)
    expect(sameCorpusMutationTargetProofProductSummaries(cases, 1)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          workload: 'edit-visible-cell',
          product: 'google-sheets',
          rawSampleCount: 1,
          acceptedSampleCount: 0,
          accepted: false,
          samples: [
            expect.objectContaining({
              sampleIndex: 0,
              accepted: false,
              invalidReasons: expect.arrayContaining([
                'semantic UI mutation target proof for edit-visible-cell did not prove a before/after target change',
                'semantic UI mutation target proof for edit-visible-cell did not prove the intended committed target value',
              ]),
            }),
          ],
        }),
      ]),
    )
  })

  it('does not count formula-bar or editor-only readbacks as grid proof', () => {
    const cases = [
      mutationTargetCase({
        biligSamples: [],
        googleProofMode: 'formula-bar-only',
        googleSamples: [0],
      }),
    ]

    expect(sameCorpusMutationTargetProofSampleCount(cases, 1)).toBe(0)
    expect(sameCorpusMutationTargetProofProductSummaries(cases, 1)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          workload: 'edit-visible-cell',
          product: 'google-sheets',
          rawSampleCount: 1,
          acceptedSampleCount: 0,
          accepted: false,
          samples: [
            expect.objectContaining({
              sampleIndex: 0,
              accepted: false,
              invalidReasons: expect.arrayContaining([
                'semantic UI mutation target proof for edit-visible-cell target readback did not come from an accepted browser-visible source',
                'semantic UI mutation target proof for edit-visible-cell visible render readback did not come from an accepted browser-visible source',
                'semantic UI mutation target proof for edit-visible-cell after screenshot semantic readback did not come from an accepted browser-visible source',
              ]),
            }),
          ],
        }),
      ]),
    )
  })

  it('does not count diagnostic captures that preserve only browser-visible state without committed workbook proof', () => {
    const cases = [
      mutationTargetCase({
        biligSamples: [],
        googleProofMode: 'missing-committed-state',
        googleSamples: [0],
      }),
    ]

    expect(sameCorpusMutationTargetProofSampleCount(cases, 1)).toBe(0)
    expect(sameCorpusMutationTargetProofProductSummaries(cases, 1)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          workload: 'edit-visible-cell',
          product: 'google-sheets',
          rawSampleCount: 1,
          acceptedSampleCount: 0,
          accepted: false,
          samples: [
            expect.objectContaining({
              sampleIndex: 0,
              accepted: false,
              committedStateProof: null,
              invalidReasons: expect.arrayContaining([
                'semantic UI mutation target proof for edit-visible-cell is missing independent Google Sheets committed-state proof',
              ]),
            }),
          ],
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
  readonly googleProofMode?: GoogleProofMode
  readonly googleSamples: readonly number[]
}): { readonly workload: 'edit-visible-cell'; readonly scenarioProof: SameCorpusScenarioProof } {
  const googleProofMode = args.googleProofMode ?? (args.corruptGoogleProof === true ? 'stale-value' : 'valid')
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
          productSemanticProof('google-sheets', args.googleSamples, googleProofMode),
        ],
        productVerdicts: [],
        missingProducts: ['google-sheets'],
      },
    },
  }
}

type GoogleProofMode = 'valid' | 'stale-value' | 'screenshot-only' | 'formula-bar-only' | 'missing-committed-state'

function productSemanticProof(
  product: UiResponsivenessSameCorpusProduct,
  sampleIndexes: readonly number[],
  googleProofMode: GoogleProofMode = 'valid',
): SameCorpusProductSemanticUiProof {
  return {
    product,
    captured: true,
    method: product === 'bilig' ? 'bilig-visible-semantic-readback' : 'google-sheets-visible-semantic-readback',
    sheetName: 'WideGrid',
    sheetId: productSemanticSheetId(product),
    selectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0),
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
    ],
    authoritativeRenderRevision: product === 'bilig' ? 'authoritative-1' : null,
    visibleRenderRevision: product === 'bilig' ? 'visible-1' : null,
    screenshotSha256: 'a'.repeat(64),
    mutationTargetProofs: sampleIndexes.map((sampleIndex) =>
      mutationTargetProof(product, sampleIndex, product === 'google-sheets' && sampleIndex === 0 ? googleProofMode : 'valid'),
    ),
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
  proofMode: GoogleProofMode,
): SameCorpusMutationTargetProof {
  const value = sameCorpusEditVisibleCellValue(sampleIndex)
  const beforeValue = 'metric-1'
  const afterValue = proofMode === 'stale-value' ? 'stale-value' : proofMode === 'screenshot-only' ? beforeValue : value
  const visibleSource = proofMode === 'formula-bar-only' ? 'visible-formula-bar' : 'visible-grid-cell'
  const committedTargetProofMs = 40 + sampleIndex
  const visibleTargetRenderMs = 12 + sampleIndex
  const committedStateValidationMs = committedTargetProofMs - visibleTargetRenderMs
  const restoreValidationMs = 80
  const operationStartedAtMs = 1000 + sampleIndex * 100
  const proof: SameCorpusMutationTargetProof = {
    product,
    sampleIndex,
    committedTargetProofMs,
    visibleTargetRenderMs,
    committedStateValidationMs,
    restoreValidationMs,
    operationStartedAtMs,
    visibleTargetRenderCapturedAtMs: operationStartedAtMs + visibleTargetRenderMs,
    postMutationProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs,
    restoreProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs + restoreValidationMs,
    workload: 'edit-visible-cell',
    intendedOperation: 'edit-visible-cell',
    intendedPayload: {
      kind: 'cell-value',
      value,
    },
    sheetName: 'WideGrid',
    sheetId: productSemanticSheetId(product),
    targetRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', sampleIndex),
    before: readback(product, beforeValue, sampleIndex, 'before'),
    after: readback(product, afterValue, sampleIndex, 'after', visibleSource),
    restored: readback(product, beforeValue, sampleIndex, 'restored'),
    visibleAfter: {
      value: afterValue,
      formula: null,
      fillColor: null,
      visibleText: afterValue,
      source: visibleSource,
      ...(product === 'bilig' ? { visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex) } : {}),
    },
    visibleAfterSelectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', sampleIndex),
    visibleRestored: {
      value: beforeValue,
      formula: null,
      fillColor: null,
      visibleText: beforeValue,
      source: 'visible-grid-cell',
      ...(product === 'bilig' ? { visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex) } : {}),
    },
    visibleRestoredSelectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', sampleIndex),
    authoritativeReadbackRevision: authoritativeReadbackRevision(sampleIndex),
    visibleRenderRevision:
      product === 'bilig'
        ? `bilig-visible-scene-sha256:${visibleSceneProofSha256(sampleIndex)}`
        : `google-sheets-screenshot-sha256:${'b'.repeat(64)}`,
    targetScreenshots: targetScreenshots(product, sampleIndex, afterValue, visibleSource),
    screenshotSha256: targetScreenshotSha256(sampleIndex, 'after'),
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/${product}-sample-${String(
      sampleIndex + 1,
    )}-after.png`,
    undoRestoreStatus: 'verified',
  }
  return signedMutationTargetProof(
    product === 'google-sheets' && proofMode !== 'missing-committed-state'
      ? { ...proof, committedStateProof: committedStateProof(proof) }
      : proof,
  )
}

function signedMutationTargetProof(
  proof: Omit<SameCorpusMutationTargetProof, 'targetProofSignature'> | SameCorpusMutationTargetProof,
): SameCorpusMutationTargetProof {
  return {
    ...proof,
    targetProofSignature: sameCorpusMutationTargetProofSignature(proof),
  }
}

function committedStateProof(proof: SameCorpusMutationTargetProof): NonNullable<SameCorpusMutationTargetProof['committedStateProof']> {
  return {
    product: 'google-sheets',
    source: 'google-sheets-xlsx-export',
    sampleIndex: proof.sampleIndex,
    workload: proof.workload,
    sheetName: proof.sheetName,
    sheetId: proof.sheetId,
    targetRange: proof.targetRange,
    before: committedStatePhaseProof(proof, 'before', proof.before),
    after: committedStatePhaseProof(proof, 'after', proof.after),
    restored: committedStatePhaseProof(proof, 'restored', proof.restored),
  }
}

function committedStatePhaseProof(
  proof: SameCorpusMutationTargetProof,
  phase: 'before' | 'after' | 'restored',
  readbackValue: SameCorpusMutationTargetProof['before'],
): NonNullable<SameCorpusMutationTargetProof['committedStateProof']>['before'] {
  const hashOffset = phase === 'before' ? 3 : phase === 'after' ? 7 : 11
  const capturedAtMs =
    phase === 'before'
      ? proof.operationStartedAtMs - 3
      : phase === 'after'
        ? proof.operationStartedAtMs + Math.max(1, proof.committedTargetProofMs / 2)
        : proof.postMutationProofCapturedAtMs + 11
  return {
    product: 'google-sheets',
    phase,
    sampleIndex: proof.sampleIndex,
    workload: proof.workload,
    sheetName: proof.sheetName,
    sheetId: proof.sheetId,
    targetRange: proof.targetRange,
    exportUrl: 'https://docs.google.com/spreadsheets/d/test-spreadsheet/export?format=xlsx',
    capturedAtMs,
    artifactPath: `tmp/same-corpus-wide-mixed-250k-edit-visible-cell/committed-state/google-sheets-sample-${String(
      proof.sampleIndex + 1,
    )}-${phase}.json`,
    artifactSha256: targetScreenshotSha256(proof.sampleIndex, phase),
    workbookByteSize: 123456 + proof.sampleIndex,
    workbookSha256: ((proof.sampleIndex + hashOffset) % 16).toString(16).repeat(64),
    readback: committedStateReadback(readbackValue),
  }
}

function committedStateReadback(sourceReadback: SameCorpusMutationTargetProof['before']): SameCorpusMutationTargetProof['before'] {
  return {
    value: sourceReadback.value,
    formula: sourceReadback.formula,
    fillColor: sourceReadback.fillColor,
    visibleText: sourceReadback.visibleText,
    source: 'google-sheets-xlsx-export',
  }
}

function targetScreenshots(
  product: UiResponsivenessSameCorpusProduct,
  sampleIndex: number,
  afterValue: string = sameCorpusEditVisibleCellValue(sampleIndex),
  afterSource: SameCorpusMutationTargetProof['after']['source'] = 'visible-grid-cell',
): SameCorpusMutationTargetProof['targetScreenshots'] {
  return {
    before: targetScreenshot(product, sampleIndex, 'before', 'metric-1', 'visible-grid-cell'),
    after: targetScreenshot(product, sampleIndex, 'after', afterValue, afterSource),
    restored: targetScreenshot(product, sampleIndex, 'restored', 'metric-1', 'visible-grid-cell'),
  }
}

function targetScreenshot(
  product: UiResponsivenessSameCorpusProduct,
  sampleIndex: number,
  phase: 'before' | 'after' | 'restored',
  value: string,
  source: SameCorpusMutationTargetProof['after']['source'],
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before'] {
  return {
    phase,
    product,
    scope: 'target-cell',
    sampleIndex,
    sheetId: productSemanticSheetId(product),
    sheetName: 'WideGrid',
    targetRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', sampleIndex),
    workload: 'edit-visible-cell',
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/${product}-sample-${String(sampleIndex + 1)}-${phase}.png`,
    screenshotSha256: targetScreenshotSha256(sampleIndex, phase),
    semanticReadback: {
      ...readback(product, value, sampleIndex, phase, source),
      source,
    },
  }
}

function targetScreenshotSha256(sampleIndex: number, phase: 'before' | 'after' | 'restored'): string {
  const hexChars = '0123456789abcdef'
  const phaseOffset = phase === 'before' ? 1 : phase === 'after' ? 11 : 2
  return hexChars[(sampleIndex * 3 + phaseOffset) % hexChars.length]?.repeat(64) ?? '0'.repeat(64)
}

function readback(
  product: UiResponsivenessSameCorpusProduct,
  value: string,
  sampleIndex: number,
  phase: 'after' | 'before' | 'restored',
  source: SameCorpusMutationTargetProof['after']['source'] = 'visible-grid-cell',
): SameCorpusMutationTargetProof['after'] {
  return {
    value,
    formula: null,
    fillColor: null,
    visibleText: value,
    source: product === 'bilig' ? 'bilig-authoritative-range' : source,
    capturedRevision:
      product === 'bilig' ? (phase === 'after' ? authoritativeReadbackRevision(sampleIndex) : `${phase}-${sampleIndex}`) : null,
    visibleSceneProofSha256: product === 'bilig' ? visibleSceneProofSha256(sampleIndex) : null,
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
