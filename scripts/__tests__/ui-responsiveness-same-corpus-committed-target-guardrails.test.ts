import { describe, expect, it } from 'vitest'

import {
  hasAcceptedCommittedTargetProofTiming,
  sameCorpusCommittedTargetProofTimingCounts,
  type SameCorpusCommittedTargetProofTimingCase,
  type SameCorpusCommittedTargetProofTimingMeasurement,
} from '../ui-responsiveness-same-corpus-guardrails.ts'
import { sameCorpusMutationTargetRangeForSample } from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { UiResponsivenessSameCorpusProduct } from '../ui-responsiveness-same-corpus-scorecard-types.ts'
import type { SameCorpusMutationTargetProof, SameCorpusProductSemanticUiProof } from '../ui-responsiveness-same-corpus-proof.ts'
import { sameCorpusEditVisibleCellValue } from '../ui-responsiveness-same-corpus-workload-runner.ts'

describe('same-corpus committed target proof guardrails', () => {
  it('counts committed mutation timing only when every expected sample is accepted', () => {
    const entry = sameCorpusEditVisibleCellCase()

    expect(hasAcceptedCommittedTargetProofTiming(entry, entry.bilig, 3)).toBe(true)
    expect(hasAcceptedCommittedTargetProofTiming(entry, entry.googleSheets, 3)).toBe(true)
    expect(sameCorpusCommittedTargetProofTimingCounts([entry], 3)).toMatchObject({
      requiredCommittedTargetProofTimingCaseCount: 3,
      committedTargetProofTimingCaseCount: 1,
      requiredCommittedTargetProofTimingSampleCount: 18,
      committedTargetProofTimingSampleCount: 6,
    })
  })

  it('counts committed mutation timing only when it is bound to accepted target proof', () => {
    const entry = sameCorpusEditVisibleCellCase({
      googleProofs: sameCorpusMutationTargetProofs('google-sheets').map((proof) =>
        proof.sampleIndex === 1 ? Object.assign({}, proof, { committedTargetProofMs: proof.committedTargetProofMs + 30 }) : proof,
      ),
    })

    expect(hasAcceptedCommittedTargetProofTiming(entry, entry.bilig, 3)).toBe(true)
    expect(hasAcceptedCommittedTargetProofTiming(entry, entry.googleSheets, 3)).toBe(false)

    expect(sameCorpusCommittedTargetProofTimingCounts([entry], 3)).toMatchObject({
      requiredCommittedTargetProofTimingCaseCount: 3,
      committedTargetProofTimingCaseCount: 0,
      requiredCommittedTargetProofTimingSampleCount: 18,
      committedTargetProofTimingSampleCount: 5,
    })
  })

  it('rejects stale target proof samples even when raw timing samples are present', () => {
    const cases = [
      sameCorpusEditVisibleCellCase({
        biligProofs: sameCorpusMutationTargetProofs('bilig').map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                visibleAfter: Object.assign({}, proof.visibleAfter, {
                  value: 'editor-only-ghost',
                  visibleText: 'editor-only-ghost',
                }),
              })
            : proof,
        ),
      }),
      sameCorpusEditVisibleCellCase({
        googleProofs: sameCorpusMutationTargetProofs('google-sheets').map((proof) =>
          proof.sampleIndex === 1
            ? Object.assign({}, proof, {
                restored: Object.assign({}, proof.after),
                undoRestoreStatus: 'failed',
              })
            : proof,
        ),
      }),
      sameCorpusEditVisibleCellCase({
        biligProofs: sameCorpusMutationTargetProofs('bilig').map((proof) =>
          proof.sampleIndex === 2
            ? Object.assign({}, proof, {
                targetRange: 'Z99',
                visibleAfterSelectedRange: 'Z99',
                visibleRestoredSelectedRange: 'Z99',
              })
            : proof,
        ),
      }),
      sameCorpusEditVisibleCellCase({
        biligProofs: sameCorpusMutationTargetProofs('bilig').map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                visibleRenderRevision: 'bilig-visible-scene-sha256:stale',
              })
            : proof,
        ),
      }),
    ]

    for (const entry of cases) {
      expect([
        hasAcceptedCommittedTargetProofTiming(entry, entry.bilig, 3),
        hasAcceptedCommittedTargetProofTiming(entry, entry.googleSheets, 3),
      ]).toContain(false)
      expect(sameCorpusCommittedTargetProofTimingCounts([entry], 3)).toMatchObject({
        committedTargetProofTimingCaseCount: 0,
        committedTargetProofTimingSampleCount: 5,
      })
    }
  })

  it('rejects timing when the semantic proof container itself is not captured', () => {
    const entry = sameCorpusEditVisibleCellCase({
      mapBiligProductProof: (proof) => ({ ...proof, captured: false }),
    })

    expect(hasAcceptedCommittedTargetProofTiming(entry, entry.bilig, 3)).toBe(false)
    expect(hasAcceptedCommittedTargetProofTiming(entry, entry.googleSheets, 3)).toBe(true)
    expect(sameCorpusCommittedTargetProofTimingCounts([entry], 3)).toMatchObject({
      committedTargetProofTimingCaseCount: 0,
      committedTargetProofTimingSampleCount: 3,
    })
  })
})

function sameCorpusEditVisibleCellCase(
  args: {
    readonly biligProofs?: readonly SameCorpusMutationTargetProof[]
    readonly googleProofs?: readonly SameCorpusMutationTargetProof[]
    readonly mapBiligProductProof?: (proof: SameCorpusProductSemanticUiProof) => SameCorpusProductSemanticUiProof
    readonly mapGoogleSheetsProductProof?: (proof: SameCorpusProductSemanticUiProof) => SameCorpusProductSemanticUiProof
  } = {},
): SameCorpusCommittedTargetProofTimingCase {
  const biligProofs = args.biligProofs ?? sameCorpusMutationTargetProofs('bilig')
  const googleProofs = args.googleProofs ?? sameCorpusMutationTargetProofs('google-sheets')
  const biligProductProof = sameCorpusProductSemanticProof('bilig', biligProofs)
  const googleSheetsProductProof = sameCorpusProductSemanticProof('google-sheets', googleProofs)
  const scenarioProof = {
    semanticUiProof: {
      products: [
        args.mapBiligProductProof ? args.mapBiligProductProof(biligProductProof) : biligProductProof,
        args.mapGoogleSheetsProductProof ? args.mapGoogleSheetsProductProof(googleSheetsProductProof) : googleSheetsProductProof,
      ],
    },
  }
  return {
    workload: 'edit-visible-cell',
    scenarioProof,
    bilig: sameCorpusMeasurement('bilig'),
    googleSheets: sameCorpusMeasurement('google-sheets'),
  }
}

function sameCorpusMeasurement(product: 'bilig' | 'google-sheets'): SameCorpusCommittedTargetProofTimingMeasurement {
  return {
    product,
    committedTargetProofMsSamples: [0, 1, 2].map((sampleIndex) => committedTargetProofMs(product, sampleIndex)),
  }
}

function sameCorpusProductSemanticProof(
  product: 'bilig' | 'google-sheets',
  mutationTargetProofs: readonly SameCorpusMutationTargetProof[],
): SameCorpusProductSemanticUiProof {
  return {
    product,
    captured: true,
    method: product === 'bilig' ? 'bilig-visible-semantic-readback' : 'google-sheets-visible-semantic-readback',
    sheetName: 'WideGrid',
    sheetId: sheetId(product),
    selectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0),
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'C1', expected: 'metric-3', actual: 'metric-3' },
    ],
    authoritativeRenderRevision: product === 'bilig' ? 'rev-1' : null,
    visibleRenderRevision: product === 'bilig' ? 'bilig-visible-scene-sha256:'.concat(visibleSceneProofSha256(0)) : null,
    screenshotSha256: 'a'.repeat(64),
    mutationTargetProofs,
    evidence: [],
  }
}

function sameCorpusMutationTargetProofs(product: 'bilig' | 'google-sheets'): SameCorpusMutationTargetProof[] {
  return [0, 1, 2].map((sampleIndex) => {
    const targetRange = sameCorpusMutationTargetRangeForSample('edit-visible-cell', sampleIndex)
    const targetValue = sameCorpusEditVisibleCellValue(sampleIndex)
    const operationStartedAtMs = 1000 + sampleIndex * 100
    const committedTargetProofMsValue = committedTargetProofMs(product, sampleIndex)
    return {
      product,
      sampleIndex,
      committedTargetProofMs: committedTargetProofMsValue,
      operationStartedAtMs,
      postMutationProofCapturedAtMs: operationStartedAtMs + committedTargetProofMsValue,
      restoreProofCapturedAtMs: operationStartedAtMs + committedTargetProofMsValue + 40,
      workload: 'edit-visible-cell',
      intendedOperation: 'edit-visible-cell',
      intendedPayload: { kind: 'cell-value', value: targetValue },
      sheetName: 'WideGrid',
      sheetId: sheetId(product),
      targetRange,
      before: sameCorpusReadback(product, sampleIndex, 'metric-1'),
      after: sameCorpusReadback(product, sampleIndex, targetValue),
      restored: sameCorpusReadback(product, sampleIndex, 'metric-1'),
      visibleAfter: sameCorpusVisibleReadback(targetValue),
      visibleAfterSelectedRange: targetRange,
      visibleRestored: sameCorpusVisibleReadback('metric-1'),
      visibleRestoredSelectedRange: targetRange,
      authoritativeReadbackRevision:
        product === 'bilig' ? `after-readback-${String(sampleIndex + 1)}` : `readback-${String(sampleIndex + 1)}`,
      visibleRenderRevision:
        product === 'bilig' ? `bilig-visible-scene-sha256:${visibleSceneProofSha256(sampleIndex)}` : `render-${String(sampleIndex + 1)}`,
      targetScreenshots: sameCorpusTargetScreenshots(product, sampleIndex),
      screenshotPath: `tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/${product}-sample-${String(sampleIndex + 1)}-after.png`,
      screenshotSha256: screenshotSha256(sampleIndex, 'after'),
      undoRestoreStatus: 'verified',
    }
  })
}

function sameCorpusReadback(product: UiResponsivenessSameCorpusProduct, sampleIndex: number, value: string) {
  const after = value === sameCorpusEditVisibleCellValue(sampleIndex)
  return {
    value,
    formula: null,
    fillColor: null,
    visibleText: value,
    source: product === 'bilig' ? ('bilig-authoritative-range' as const) : ('visible-formula-bar' as const),
    ...(product === 'bilig'
      ? {
          capturedRevision: `${after ? 'after' : 'before'}-readback-${String(sampleIndex + 1)}`,
          ...(after ? { visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex) } : {}),
        }
      : {}),
  }
}

function sameCorpusVisibleReadback(value: string) {
  return {
    value,
    formula: null,
    fillColor: null,
    visibleText: value,
    source: 'visible-formula-bar' as const,
  }
}

function sameCorpusTargetScreenshots(product: 'bilig' | 'google-sheets', sampleIndex: number) {
  return {
    before: sameCorpusTargetScreenshot(product, sampleIndex, 'before'),
    after: sameCorpusTargetScreenshot(product, sampleIndex, 'after'),
    restored: sameCorpusTargetScreenshot(product, sampleIndex, 'restored'),
  }
}

function sameCorpusTargetScreenshot(product: 'bilig' | 'google-sheets', sampleIndex: number, phase: 'before' | 'after' | 'restored') {
  return {
    phase,
    product,
    scope: 'target-cell' as const,
    sampleIndex,
    sheetId: sheetId(product),
    sheetName: 'WideGrid',
    targetRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', sampleIndex),
    workload: 'edit-visible-cell' as const,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/${product}-sample-${String(sampleIndex + 1)}-${phase}.png`,
    screenshotSha256: screenshotSha256(sampleIndex, phase),
  }
}

function committedTargetProofMs(product: 'bilig' | 'google-sheets', sampleIndex: number): number {
  return product === 'bilig' ? 40 + sampleIndex : 400 + sampleIndex * 10
}

function sheetId(product: 'bilig' | 'google-sheets'): string {
  return product === 'bilig' ? 'sheet-wide-grid' : 'gid:160971404'
}

function screenshotSha256(sampleIndex: number, phase: 'before' | 'after' | 'restored'): string {
  const seed = phase === 'before' ? 1 : phase === 'after' ? 5 : 9
  return String((sampleIndex + seed) % 10).repeat(64)
}

function visibleSceneProofSha256(sampleIndex: number): string {
  return String(sampleIndex + 1)
    .repeat(64)
    .slice(0, 64)
}
