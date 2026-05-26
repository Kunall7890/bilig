import { describe, expect, it } from 'vitest'

import {
  validateSameCorpusProductSemanticUiProof,
  type SameCorpusMutationTargetProof,
  type SameCorpusProductSemanticUiProof,
} from '../ui-responsiveness-same-corpus-proof.ts'
import type { UiResponsivenessSameCorpusProduct } from '../ui-responsiveness-same-corpus-scorecard-proof.ts'
import type { UiResponsivenessSameCorpusWorkload } from '../ui-responsiveness-same-corpus-workloads.ts'

describe('same-corpus semantic UI mutation proof validation', () => {
  it('accepts mutation target proof tied to exact workload and sample screenshots', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(validSemanticProof(), {
      workload: 'edit-visible-cell',
      sampleCount: 3,
    })

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: true,
      invalidReasons: [],
    })
  })

  it('rejects rendered selection text that merely contains the target range', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        selectedRange: 'visible selection A1 after editor commit',
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell target range does not match the rendered selection',
      ]),
    })
  })

  it('rejects mutation screenshots that are not tied to the workload and sample', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                screenshotPath: 'tmp/mutation-target/bilig-sample-1.png',
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell screenshot artifact path is not tied to the workload',
        'semantic UI mutation target proof for edit-visible-cell screenshot artifact path is not tied to the mutation target sample',
      ]),
    })
  })

  it('rejects reused mutation target screenshots across samples', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 1
            ? Object.assign({}, proof, {
                screenshotPath: 'tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/bilig-sample-1-after.png',
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell duplicates screenshot artifact path',
        'semantic UI mutation target proof for edit-visible-cell screenshot artifact path is not tied to the mutation target sample',
      ]),
    })
  })
})

function validSemanticProof(overrides: Partial<SameCorpusProductSemanticUiProof> = {}): SameCorpusProductSemanticUiProof {
  return {
    product: 'bilig',
    captured: true,
    method: 'bilig-visible-semantic-readback',
    sheetName: 'WideGrid',
    sheetId: 'sheet-wide-grid',
    selectedRange: 'A1',
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
    ],
    authoritativeRenderRevision: 'rev-3',
    visibleRenderRevision: 'scene-7',
    screenshotSha256: 'a'.repeat(64),
    mutationTargetProofs: validMutationTargetProofs(),
    evidence: [
      'sheetName=WideGrid',
      'selectedRange=A1',
      'checkedCellCount=3',
      `screenshotSha256=${'a'.repeat(64)}`,
      'authoritativeRenderRevision=rev-3',
      'visibleRenderRevision=scene-7',
    ],
    ...overrides,
  }
}

function validMutationTargetProofs(): SameCorpusMutationTargetProof[] {
  return [0, 1, 2].map((sampleIndex) => mutationTargetProof('bilig', 'edit-visible-cell', sampleIndex))
}

function mutationTargetProof(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof {
  return {
    sampleIndex,
    workload,
    intendedOperation: 'edit-visible-cell',
    intendedPayload: {
      kind: 'cell-value',
      value: `${product}-same-corpus-${String(sampleIndex + 1)}`,
    },
    sheetName: 'WideGrid',
    targetRange: 'A1',
    before: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
    },
    after: {
      value: `${product}-same-corpus-${String(sampleIndex + 1)}`,
      formula: null,
      fillColor: null,
      visibleText: `${product}-same-corpus-${String(sampleIndex + 1)}`,
      source: 'bilig-authoritative-range',
    },
    restored: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
    },
    visibleAfter: {
      value: `${product}-same-corpus-${String(sampleIndex + 1)}`,
      formula: null,
      fillColor: null,
      visibleText: `${product}-same-corpus-${String(sampleIndex + 1)}`,
      source: 'visible-formula-bar',
    },
    visibleRestored: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'visible-formula-bar',
    },
    authoritativeReadbackRevision: `authoritative-readback-${String(sampleIndex + 1)}`,
    visibleRenderRevision: `visible-render-${String(sampleIndex + 1)}`,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${String(sampleIndex + 1)}-after.png`,
    screenshotSha256: 'a'.repeat(64),
    undoRestoreStatus: 'verified',
  }
}
