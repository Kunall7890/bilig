import type { CellRangeRef } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import type { WorkbookAuthoritativeReadbackProof } from './workbook-agent-mutation-proof-types.js'
import { buildWorkbookSemanticReadbackProof } from './workbook-agent-mutation-semantic-readback.js'
import type { WorkbookRenderedReadbackProof } from './workbook-agent-rendered-readback.js'

const range: CellRangeRef = {
  sheetName: 'Sheet1',
  startAddress: 'B2',
  endAddress: 'B2',
}

function authoritativeReadback(overrides: Partial<WorkbookAuthoritativeReadbackProof> = {}): WorkbookAuthoritativeReadbackProof {
  return {
    requested: true,
    matched: true,
    ranges: [],
    mismatches: [],
    incompleteReason: null,
    ...overrides,
  }
}

function renderedReadback(overrides: Partial<WorkbookRenderedReadbackProof> = {}): WorkbookRenderedReadbackProof {
  return {
    requested: true,
    requestedRange: range,
    available: true,
    matched: true,
    stale: false,
    capturedRange: range,
    sourceRange: range,
    capturedAtUnixMs: 1,
    capturedRevision: 2,
    capturedBatchId: 1,
    truncated: false,
    sourceTruncated: false,
    missingCells: [],
    mismatches: [],
    incompleteReason: null,
    nextChunk: null,
    range: null,
    ...overrides,
  }
}

describe('buildWorkbookSemanticReadbackProof', () => {
  it('requires browser-rendered readback before reporting a semantic match', () => {
    const proof = buildWorkbookSemanticReadbackProof({
      authoritativeReadback: authoritativeReadback(),
      renderedReadback: renderedReadback({
        requested: false,
        matched: null,
        incompleteReason: 'No browser-rendered context was attached to this tool call.',
      }),
    })

    expect(proof).toEqual({
      requested: true,
      matched: false,
      incompleteReason: 'No browser-rendered context was attached to this tool call.',
    })
  })

  it('reports a semantic match only when authoritative and rendered readback both match', () => {
    expect(
      buildWorkbookSemanticReadbackProof({
        authoritativeReadback: authoritativeReadback(),
        renderedReadback: renderedReadback(),
      }),
    ).toEqual({
      requested: true,
      matched: true,
      incompleteReason: null,
    })
  })
})
