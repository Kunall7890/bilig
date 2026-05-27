import { describe, expect, it } from 'vitest'

import { SameCorpusCommittedStateMismatchError } from '../ui-responsiveness-same-corpus-committed-state-proof.ts'
import {
  maybeCaptureIncompleteCommittedStatePhaseProof,
  sameCorpusCommittedStateExpectedReadback,
} from '../ui-responsiveness-same-corpus-mutation-target-capture.ts'
import type { SameCorpusMutationTargetReadback } from '../ui-responsiveness-same-corpus-proof.ts'

describe('same-corpus mutation target capture', () => {
  it('requires Google Sheets fill exports to prove the intended swatch, not stale browser readback', () => {
    const before = readback({ fillColor: null, value: 'segment-5' })
    const selectionChromeAfter = readback({ fillColor: 'rgb(11, 87, 208)', value: 'segment-5' })

    expect(
      sameCorpusCommittedStateExpectedReadback({
        before,
        phase: 'after',
        phaseReadback: selectionChromeAfter,
        sampleIndex: 0,
        workload: 'fill-format-change',
      }),
    ).toEqual({
      ...selectionChromeAfter,
      fillColor: '#c9daf8',
    })
  })

  it('requires restored committed-state exports to match the pre-mutation target', () => {
    const before = readback({ fillColor: null, value: 'segment-5' })
    const staleRestored = readback({ fillColor: '#c9daf8', value: 'segment-5' })

    expect(
      sameCorpusCommittedStateExpectedReadback({
        before,
        phase: 'restored',
        phaseReadback: staleRestored,
        sampleIndex: 0,
        workload: 'fill-format-change',
      }),
    ).toEqual(before)
  })

  it('preserves non-fill post-mutation readbacks for committed-state polling', () => {
    const before = readback({ formula: null, value: '10' })
    const after = readback({ formula: '=1+1', value: '2' })

    expect(
      sameCorpusCommittedStateExpectedReadback({
        before,
        phase: 'after',
        phaseReadback: after,
        sampleIndex: 0,
        workload: 'formula-edit',
      }),
    ).toEqual(after)
  })

  it('keeps diagnostic captures moving when only Google Sheets committed-state export mismatches', async () => {
    await expect(
      maybeCaptureIncompleteCommittedStatePhaseProof(
        { allowIncompleteCommittedStateProof: true },
        Promise.reject(committedStateMismatchError()),
      ),
    ).resolves.toBeNull()
  })

  it('keeps claim-grade captures strict for Google Sheets committed-state export mismatches', async () => {
    await expect(
      maybeCaptureIncompleteCommittedStatePhaseProof(
        { allowIncompleteCommittedStateProof: false },
        Promise.reject(committedStateMismatchError()),
      ),
    ).rejects.toMatchObject({ name: 'SameCorpusCommittedStateMismatchError' })
  })

  it('does not hide non-mismatch committed-state proof failures in diagnostic captures', async () => {
    await expect(
      maybeCaptureIncompleteCommittedStatePhaseProof(
        { allowIncompleteCommittedStateProof: true },
        Promise.reject(new Error('browser closed')),
      ),
    ).rejects.toThrow('browser closed')
  })
})

function readback(values: Partial<SameCorpusMutationTargetReadback>): SameCorpusMutationTargetReadback {
  return {
    fillColor: null,
    formula: null,
    source: 'visible-grid-cell',
    value: null,
    visibleText: values.value ?? values.formula ?? null,
    ...values,
  }
}

function committedStateMismatchError(): SameCorpusCommittedStateMismatchError {
  const expectedReadback = readback({ fillColor: '#c9daf8', source: 'visible-grid-cell', value: 'segment-5' })
  const lastReadback = readback({ fillColor: null, source: 'google-sheets-xlsx-export', value: 'segment-5' })
  return new SameCorpusCommittedStateMismatchError({
    expectedReadback,
    exportUrl: 'https://docs.google.com/spreadsheets/d/test-spreadsheet/export?format=xlsx',
    lastArtifactPath: 'proof/after.json',
    lastArtifactSha256: 'a'.repeat(64),
    lastReadback,
    lastWorkbookSha256: 'b'.repeat(64),
    phase: 'after',
    product: 'google-sheets',
    sampleIndex: 0,
    sheetId: 'gid:1',
    sheetName: 'WideGrid',
    targetRange: 'B5',
    workload: 'fill-format-change',
  })
}
