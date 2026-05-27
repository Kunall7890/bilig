import { describe, expect, it } from 'vitest'

import { sameCorpusCommittedStateExpectedReadback } from '../ui-responsiveness-same-corpus-mutation-target-capture.ts'
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
