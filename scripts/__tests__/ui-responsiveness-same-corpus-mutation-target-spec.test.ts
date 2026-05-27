import { describe, expect, it } from 'vitest'

import {
  validateSameCorpusProductSemanticUiProof,
  type SameCorpusMutationTargetProof,
  type SameCorpusMutationTargetReadback,
  type SameCorpusMutationTargetScreenshotProof,
  type SameCorpusProductSemanticUiProof,
} from '../ui-responsiveness-same-corpus-proof.ts'
import { sameCorpusMutationTargetRangeForSample } from '../ui-responsiveness-same-corpus-mutation-target-spec.ts'
import { sameCorpusMutationTargetProofSignature } from '../ui-responsiveness-same-corpus-mutation-target-signature.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from '../ui-responsiveness-same-corpus-workloads.ts'

describe('same-corpus mutation target range contract', () => {
  it('accepts per-sample mutation proof only for the declared workload target range', () => {
    const workload = 'edit-visible-cell'

    const accepted = validateSameCorpusProductSemanticUiProof(semanticProof(workload), {
      workload,
      sampleCount: 3,
    })
    const rejected = validateSameCorpusProductSemanticUiProof(
      semanticProof(workload, (sample) => (sample.sampleIndex === 1 ? { ...sample, targetRange: 'C99' } : sample)),
      {
        workload,
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
        'semantic UI mutation target proof for edit-visible-cell target range does not match the declared sample target',
      ]),
    })
  })
})

function semanticProof(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  mapSample: (sample: SameCorpusMutationTargetProof) => SameCorpusMutationTargetProof = (sample) => sample,
): SameCorpusProductSemanticUiProof {
  return {
    product: 'bilig',
    captured: true,
    method: 'bilig-visible-semantic-readback',
    sheetName: 'WideGrid',
    sheetId: '1',
    selectedRange: sameCorpusMutationTargetRangeForSample(workload, 0),
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B2', expected: '42', actual: '42' },
      { address: 'C3', expected: 'done', actual: 'done' },
    ],
    authoritativeRenderRevision: 'bilig-render-revision-1',
    visibleRenderRevision: `bilig-visible-scene-sha256:${'d'.repeat(64)}`,
    screenshotSha256: 'a'.repeat(64),
    mutationTargetProofs: [0, 1, 2].map((sampleIndex) => mapSample(mutationTargetProof(workload, sampleIndex))),
    evidence: ['semanticUiProofVersion=semantic-ui-v1', 'sheetId=1'],
  }
}

function mutationTargetProof(workload: UiResponsivenessSameCorpusMutatingWorkload, sampleIndex: number): SameCorpusMutationTargetProof {
  const targetRange = sameCorpusMutationTargetRangeForSample(workload, sampleIndex)
  const operationStartedAtMs = 100 + sampleIndex
  const visibleTargetRenderMs = 16
  const committedTargetProofMs = 40
  const restoreValidationMs = 8
  const after = afterReadback(workload, sampleIndex)
  const proof: Omit<SameCorpusMutationTargetProof, 'targetProofSignature'> = {
    product: 'bilig',
    sampleIndex,
    committedTargetProofMs,
    visibleTargetRenderMs,
    committedStateValidationMs: committedTargetProofMs - visibleTargetRenderMs,
    restoreValidationMs,
    operationStartedAtMs,
    visibleTargetRenderCapturedAtMs: operationStartedAtMs + visibleTargetRenderMs,
    postMutationProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs,
    restoreProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs + restoreValidationMs,
    workload,
    intendedOperation: workload,
    intendedPayload: intendedPayload(workload, sampleIndex),
    sheetName: 'WideGrid',
    sheetId: '1',
    targetRange,
    before: beforeReadback(sampleIndex),
    after,
    restored: beforeReadback(sampleIndex),
    visibleAfter: { ...after, source: 'visible-grid-cell' },
    visibleRestored: { ...beforeReadback(sampleIndex), source: 'visible-grid-cell' },
    committedStateProof: null,
    visibleAfterSelectedRange: targetRange,
    visibleRestoredSelectedRange: targetRange,
    authoritativeReadbackRevision: `bilig-authoritative-readback:${sampleIndex}`,
    visibleRenderRevision: `bilig-visible-scene-sha256:${'d'.repeat(64)}`,
    targetScreenshots: {
      before: targetScreenshot(workload, sampleIndex, 'before', targetRange, beforeReadback(sampleIndex)),
      after: targetScreenshot(workload, sampleIndex, 'after', targetRange, after),
      restored: targetScreenshot(workload, sampleIndex, 'restored', targetRange, beforeReadback(sampleIndex)),
    },
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/bilig-sample-${String(sampleIndex + 1)}-after.png`,
    screenshotSha256: screenshotHash(sampleIndex, 'after'),
    undoRestoreStatus: 'verified',
  }
  return signedMutationTargetProof(proof)
}

function signedMutationTargetProof(
  proof: Omit<SameCorpusMutationTargetProof, 'targetProofSignature'> | SameCorpusMutationTargetProof,
): SameCorpusMutationTargetProof {
  return {
    ...proof,
    targetProofSignature: sameCorpusMutationTargetProofSignature(proof),
  }
}

function beforeReadback(sampleIndex: number): SameCorpusMutationTargetReadback {
  return {
    value: `before-${sampleIndex}`,
    formula: null,
    fillColor: null,
    visibleText: `before-${sampleIndex}`,
    source: 'bilig-authoritative-range',
    capturedRevision: `bilig-authoritative-readback:${sampleIndex}`,
    visibleSceneProofSha256: 'd'.repeat(64),
  }
}

function afterReadback(workload: UiResponsivenessSameCorpusMutatingWorkload, sampleIndex: number): SameCorpusMutationTargetReadback {
  if (workload === 'formula-edit') {
    return {
      value: String(sampleIndex + 2),
      formula: `=${String(sampleIndex + 1)}+1`,
      fillColor: null,
      visibleText: String(sampleIndex + 2),
      source: 'bilig-authoritative-range',
      capturedRevision: `bilig-authoritative-readback:${sampleIndex}`,
      visibleSceneProofSha256: 'd'.repeat(64),
    }
  }
  if (workload === 'fill-format-change') {
    return {
      value: `before-${sampleIndex}`,
      formula: null,
      fillColor: '#00ff00',
      visibleText: `before-${sampleIndex}`,
      source: 'bilig-authoritative-range',
      capturedRevision: `bilig-authoritative-readback:${sampleIndex}`,
      visibleSceneProofSha256: 'd'.repeat(64),
    }
  }
  return {
    value: `same-corpus-edit-${String(sampleIndex + 1)}`,
    formula: null,
    fillColor: null,
    visibleText: `same-corpus-edit-${String(sampleIndex + 1)}`,
    source: 'bilig-authoritative-range',
    capturedRevision: `bilig-authoritative-readback:${sampleIndex}`,
    visibleSceneProofSha256: 'd'.repeat(64),
  }
}

function intendedPayload(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['intendedPayload'] {
  if (workload === 'formula-edit') {
    return { kind: 'formula', formula: `=${String(sampleIndex + 1)}+1` }
  }
  if (workload === 'fill-format-change') {
    return { kind: 'fill-color', expectedFillColor: '#00ff00', swatchLabel: 'green' }
  }
  return { kind: 'cell-value', value: `same-corpus-edit-${String(sampleIndex + 1)}` }
}

function targetScreenshot(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
  phase: SameCorpusMutationTargetScreenshotProof['phase'],
  targetRange: string,
  semanticReadback: SameCorpusMutationTargetReadback,
): SameCorpusMutationTargetScreenshotProof {
  return {
    phase,
    product: 'bilig',
    scope: 'target-cell',
    sampleIndex,
    sheetId: '1',
    sheetName: 'WideGrid',
    targetRange,
    workload,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/bilig-sample-${String(sampleIndex + 1)}-${phase}.png`,
    screenshotSha256: screenshotHash(sampleIndex, phase),
    semanticReadback: { ...semanticReadback, source: 'visible-grid-cell' },
  }
}

function screenshotHash(sampleIndex: number, phase: SameCorpusMutationTargetScreenshotProof['phase']): string {
  const phaseOffset = phase === 'before' ? 0 : phase === 'after' ? 1 : 2
  return String(sampleIndex + phaseOffset + 1)
    .repeat(64)
    .slice(0, 64)
}
