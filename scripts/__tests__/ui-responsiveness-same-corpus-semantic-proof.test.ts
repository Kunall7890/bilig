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

  it('rejects mutation target proof without committed target proof timing', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { committedTargetProofMs: Number.NaN }) : proof,
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
        'semantic UI mutation target proof for edit-visible-cell is missing committed target proof timing',
      ]),
    })
  })

  it('rejects mutation target proof without before, after, and restored target-cell screenshots', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { targetScreenshots: null }) : proof,
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
        'semantic UI mutation target proof for edit-visible-cell is missing target-cell screenshots',
      ]),
    })
  })

  it('rejects mutation target screenshots that do not visibly change after mutation', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? staleAfterMutationTargetScreenshotProof(proof) : proof,
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
        'semantic UI mutation target proof for edit-visible-cell before and after target screenshots are identical',
      ]),
    })
  })

  it('rejects mutation target screenshots that do not visibly restore after undo', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? staleRestoredMutationTargetScreenshotProof(proof) : proof,
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
        'semantic UI mutation target proof for edit-visible-cell after and restored target screenshots are identical',
      ]),
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

  it('rejects semantic UI proof without a concrete sheet id', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        sheetId: null,
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining(['semantic UI proof is missing sheet id']),
    })
  })

  it('rejects mutation target proof whose sheet id does not match the semantic proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { sheetId: 'stale-sheet-id' }) : proof,
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
        'semantic UI mutation target proof for edit-visible-cell target sheet id does not match semantic UI proof',
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

  it('rejects Bilig mutation proof whose revisions are not bound to the target readback and scene proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                after: Object.assign({}, proof.after, {
                  capturedRevision: 'stale-readback-revision',
                  visibleSceneProofSha256: null,
                }),
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
        'semantic UI mutation target proof for edit-visible-cell authoritative revision does not match target readback',
        'semantic UI mutation target proof for edit-visible-cell is missing Bilig visible scene proof',
      ]),
    })
  })

  it('rejects Bilig mutation proof whose visible render revision does not match the target scene proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                visibleRenderRevision: `bilig-visible-scene-sha256:${'f'.repeat(64)}`,
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
        'semantic UI mutation target proof for edit-visible-cell visible render revision does not match target scene proof',
      ]),
    })
  })

  it('requires Bilig fill-format visible proof to come from rendered grid cell pixels', () => {
    const accepted = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: [0, 1, 2].map((sampleIndex) => fillMutationTargetProof(sampleIndex, 'visible-grid-cell')),
      }),
      {
        workload: 'fill-format-change',
        sampleCount: 3,
      },
    )
    const rejected = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: [0, 1, 2].map((sampleIndex) => fillMutationTargetProof(sampleIndex, 'visible-formula-bar')),
      }),
      {
        workload: 'fill-format-change',
        sampleCount: 3,
      },
    )

    expect(accepted).toMatchObject({
      acceptedForCurrentScorecard: true,
      invalidReasons: [],
    })
    expect(rejected).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for fill-format-change visible render readback did not come from rendered grid cell pixels',
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
    committedTargetProofMs: 40 + sampleIndex,
    workload,
    intendedOperation: 'edit-visible-cell',
    intendedPayload: {
      kind: 'cell-value',
      value: `${product}-same-corpus-${String(sampleIndex + 1)}`,
    },
    sheetName: 'WideGrid',
    sheetId: 'sheet-wide-grid',
    targetRange: 'A1',
    before: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: `before-readback-${String(sampleIndex + 1)}`,
    },
    after: {
      value: `${product}-same-corpus-${String(sampleIndex + 1)}`,
      formula: null,
      fillColor: null,
      visibleText: `${product}-same-corpus-${String(sampleIndex + 1)}`,
      source: 'bilig-authoritative-range',
      capturedRevision: authoritativeReadbackRevision(sampleIndex),
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    },
    restored: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: `restored-readback-${String(sampleIndex + 1)}`,
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
    authoritativeReadbackRevision: authoritativeReadbackRevision(sampleIndex),
    visibleRenderRevision: visibleRenderRevision(sampleIndex),
    targetScreenshots: mutationTargetScreenshots(product, workload, sampleIndex),
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${String(sampleIndex + 1)}-after.png`,
    screenshotSha256: mutationTargetScreenshotSha256(sampleIndex, 'after'),
    undoRestoreStatus: 'verified',
  }
}

function fillMutationTargetProof(
  sampleIndex: number,
  visibleSource: 'visible-formula-bar' | 'visible-grid-cell',
): SameCorpusMutationTargetProof {
  const fillColor = fillColorForSample(sampleIndex)
  return {
    sampleIndex,
    committedTargetProofMs: 60 + sampleIndex,
    workload: 'fill-format-change',
    intendedOperation: 'fill-format-change',
    intendedPayload: {
      kind: 'fill-color',
      expectedFillColor: fillColor,
      swatchLabel: `swatch-${String(sampleIndex + 1)}`,
    },
    sheetName: 'WideGrid',
    sheetId: 'sheet-wide-grid',
    targetRange: 'A1',
    before: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: `before-readback-${String(sampleIndex + 1)}`,
    },
    after: {
      value: 'metric-1',
      formula: null,
      fillColor,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: authoritativeReadbackRevision(sampleIndex),
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    },
    restored: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: `restored-readback-${String(sampleIndex + 1)}`,
    },
    visibleAfter: {
      value: null,
      formula: null,
      fillColor,
      visibleText: null,
      source: visibleSource,
    },
    visibleRestored: {
      value: null,
      formula: null,
      fillColor: null,
      visibleText: null,
      source: visibleSource,
    },
    authoritativeReadbackRevision: authoritativeReadbackRevision(sampleIndex),
    visibleRenderRevision: visibleRenderRevision(sampleIndex),
    targetScreenshots: mutationTargetScreenshots('bilig', 'fill-format-change', sampleIndex),
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-fill-format-change/mutation-target/bilig-sample-${String(sampleIndex + 1)}-after.png`,
    screenshotSha256: mutationTargetScreenshotSha256(sampleIndex, 'after'),
    undoRestoreStatus: 'verified',
  }
}

function mutationTargetScreenshots(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['targetScreenshots'] {
  return {
    before: mutationTargetScreenshot(product, workload, sampleIndex, 'before'),
    after: mutationTargetScreenshot(product, workload, sampleIndex, 'after'),
    restored: mutationTargetScreenshot(product, workload, sampleIndex, 'restored'),
  }
}

function mutationTargetScreenshot(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusWorkload,
  sampleIndex: number,
  phase: 'before' | 'after' | 'restored',
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before'] {
  return {
    phase,
    scope: 'target-cell',
    targetRange: 'A1',
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${String(sampleIndex + 1)}-${phase}.png`,
    screenshotSha256: mutationTargetScreenshotSha256(sampleIndex, phase),
  }
}

function staleAfterMutationTargetScreenshotProof(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (!proof.targetScreenshots) {
    return proof
  }
  const beforeHash = proof.targetScreenshots.before.screenshotSha256
  return {
    ...proof,
    screenshotSha256: beforeHash,
    targetScreenshots: {
      ...proof.targetScreenshots,
      after: {
        ...proof.targetScreenshots.after,
        screenshotSha256: beforeHash,
      },
    },
  }
}

function staleRestoredMutationTargetScreenshotProof(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (!proof.targetScreenshots) {
    return proof
  }
  return {
    ...proof,
    targetScreenshots: {
      ...proof.targetScreenshots,
      restored: {
        ...proof.targetScreenshots.restored,
        screenshotSha256: proof.targetScreenshots.after.screenshotSha256,
      },
    },
  }
}

function mutationTargetScreenshotSha256(sampleIndex: number, phase: 'before' | 'after' | 'restored'): string {
  const hexChars = '0123456789abcdef'
  const phaseOffset = phase === 'before' ? 1 : phase === 'after' ? 5 : 9
  return hexChars[(sampleIndex + phaseOffset) % hexChars.length]?.repeat(64) ?? '0'.repeat(64)
}

function fillColorForSample(sampleIndex: number): string {
  const colors = ['#c9daf8', '#34a853', '#a4c2f4'] as const
  return colors[sampleIndex % colors.length]
}

function authoritativeReadbackRevision(sampleIndex: number): string {
  return `authoritative-readback-${String(sampleIndex + 1)}`
}

function visibleRenderRevision(sampleIndex: number): string {
  return `bilig-visible-scene-sha256:${visibleSceneProofSha256(sampleIndex)}`
}

function visibleSceneProofSha256(sampleIndex: number): string {
  return String(sampleIndex + 1)
    .repeat(64)
    .slice(0, 64)
}
