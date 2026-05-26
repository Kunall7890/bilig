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
})

function sameCorpusEditVisibleCellCase(args: {
  readonly googleProofs: readonly SameCorpusMutationTargetProof[]
}): SameCorpusCommittedTargetProofTimingCase {
  const biligProofs = sameCorpusMutationTargetProofs('bilig')
  const scenarioProof = {
    semanticUiProof: {
      products: [sameCorpusProductSemanticProof('bilig', biligProofs), sameCorpusProductSemanticProof('google-sheets', args.googleProofs)],
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
    checkedCells: [],
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
