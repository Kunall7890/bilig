import { describe, expect, it } from 'vitest'

import type { SameCorpusMutationTargetProof, SameCorpusMutationTargetReadback } from '../ui-responsiveness-same-corpus-proof.ts'
import { sameCorpusMutationTargetScreenshotSemanticInvalidReasons } from '../ui-responsiveness-same-corpus-target-screenshot-proof.ts'

describe('same-corpus mutation target screenshot semantic proof', () => {
  it('accepts target screenshots whose semantic readback matches each phase', () => {
    expect(sameCorpusMutationTargetScreenshotSemanticInvalidReasons('bilig', 'edit-visible-cell', proof())).toEqual([])
  })

  it('rejects a target screenshot whose semantic readback is stale', () => {
    const targetProof = proof({
      targetScreenshots: {
        ...proof().targetScreenshots!,
        after: {
          ...proof().targetScreenshots!.after,
          semanticReadback: readback('stale', { source: 'visible-grid-cell' }),
        },
      },
    })

    expect(sameCorpusMutationTargetScreenshotSemanticInvalidReasons('bilig', 'edit-visible-cell', targetProof)).toEqual([
      'semantic UI mutation target proof for edit-visible-cell after screenshot semantic readback does not match target readback',
    ])
  })

  it('rejects a Bilig target screenshot whose semantic readback is tied to a stale visible scene', () => {
    const expectedScene = 'a'.repeat(64)
    const staleScene = 'f'.repeat(64)
    const targetProof = proof({
      after: readback('edited-value', { visibleSceneProofSha256: expectedScene }),
      targetScreenshots: {
        ...proof().targetScreenshots!,
        after: {
          ...proof().targetScreenshots!.after,
          semanticReadback: readback('edited-value', { source: 'visible-grid-cell', visibleSceneProofSha256: staleScene }),
        },
      },
    })

    expect(sameCorpusMutationTargetScreenshotSemanticInvalidReasons('bilig', 'edit-visible-cell', targetProof)).toEqual([
      'semantic UI mutation target proof for edit-visible-cell after screenshot semantic readback visible scene proof does not match target readback',
    ])
  })

  it('matches fill-color screenshots by normalized color rather than hash-only evidence', () => {
    const targetProof = proof({
      workload: 'fill-format-change',
      intendedOperation: 'fill-format-change',
      intendedPayload: { kind: 'fill-color', expectedFillColor: '#00ff00', swatchLabel: 'green' },
      before: readback('metric-1', { fillColor: null }),
      after: readback('metric-1', { fillColor: '#00ff00' }),
      restored: readback('metric-1', { fillColor: null }),
      visibleAfter: readback('metric-1', { fillColor: '#00ff00', source: 'visible-grid-cell' }),
      visibleRestored: readback('metric-1', { fillColor: null, source: 'visible-grid-cell' }),
      targetScreenshots: {
        before: screenshot('before', readback('metric-1', { fillColor: null }), 'fill-format-change'),
        after: screenshot('after', readback('metric-1', { fillColor: 'rgb(0, 255, 0)' }), 'fill-format-change'),
        restored: screenshot('restored', readback('metric-1', { fillColor: null }), 'fill-format-change'),
      },
    })

    expect(sameCorpusMutationTargetScreenshotSemanticInvalidReasons('bilig', 'fill-format-change', targetProof)).toEqual([])
  })
})

function proof(overrides: Partial<SameCorpusMutationTargetProof> = {}): SameCorpusMutationTargetProof {
  const before = readback('metric-1')
  const after = readback('edited-value')
  const restored = readback('metric-1')
  return {
    product: 'bilig',
    sampleIndex: 0,
    committedTargetProofMs: 42,
    operationStartedAtMs: 1000,
    postMutationProofCapturedAtMs: 1042,
    restoreProofCapturedAtMs: 1100,
    workload: 'edit-visible-cell',
    intendedOperation: 'edit-visible-cell',
    intendedPayload: { kind: 'cell-value', value: 'edited-value' },
    sheetName: 'WideGrid',
    sheetId: 'sheet-wide-grid',
    targetRange: 'C5',
    before,
    after,
    restored,
    visibleAfter: readback('edited-value', { source: 'visible-grid-cell' }),
    visibleAfterSelectedRange: 'C5',
    visibleRestored: readback('metric-1', { source: 'visible-grid-cell' }),
    visibleRestoredSelectedRange: 'C5',
    authoritativeReadbackRevision: 'readback-1',
    visibleRenderRevision: 'visible-1',
    targetScreenshots: {
      before: screenshot('before', before),
      after: screenshot('after', after),
      restored: screenshot('restored', restored),
    },
    screenshotPath: 'tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/bilig-sample-1-after.png',
    screenshotSha256: 'b'.repeat(64),
    undoRestoreStatus: 'verified',
    ...overrides,
  }
}

function screenshot(
  phase: 'before' | 'after' | 'restored',
  semanticReadback: SameCorpusMutationTargetReadback,
  workload: SameCorpusMutationTargetProof['workload'] = 'edit-visible-cell',
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before'] {
  return {
    phase,
    product: 'bilig',
    scope: 'target-cell',
    sampleIndex: 0,
    sheetId: 'sheet-wide-grid',
    sheetName: 'WideGrid',
    targetRange: 'C5',
    workload,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/bilig-sample-1-${phase}.png`,
    screenshotSha256: (phase === 'before' ? 'a' : phase === 'after' ? 'b' : 'c').repeat(64),
    semanticReadback: { ...semanticReadback, source: 'visible-grid-cell' },
  }
}

function readback(value: string, overrides: Partial<SameCorpusMutationTargetReadback> = {}): SameCorpusMutationTargetReadback {
  return {
    value,
    formula: null,
    fillColor: null,
    visibleText: value,
    source: 'bilig-authoritative-range',
    visibleSceneProofSha256: 'a'.repeat(64),
    ...overrides,
  }
}
