import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { WorkbookAgentRenderedContext, WorkbookAgentRenderedSurfaceProof } from '@bilig/contracts'
import { selectWorkbookRenderedReadback } from './workbook-agent-rendered-readback.js'

function renderedContext(input: {
  readonly capturedRevision?: number | null
  readonly batchId: number | null
  readonly omitSurfaceProof?: boolean
  readonly surfaceRevision?: number | null
  readonly surfaceProofOverrides?: Partial<WorkbookAgentRenderedSurfaceProof>
  readonly useVisibleRangeOnly?: boolean
  readonly value?: string
}): WorkbookAgentRenderedContext {
  const renderedRange = {
    range: {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'A1',
    },
    rowCount: 1,
    columnCount: 1,
    cellCount: 1,
    truncated: false,
    rows: [
      [
        {
          address: 'A1',
          input: input.value ?? 'ok',
          value: { tag: ValueTag.String, value: input.value ?? 'ok' },
          formula: null,
          displayFormat: input.value ?? 'ok',
          styleId: null,
          numberFormatId: null,
          style: null,
        },
      ],
    ],
  }
  return {
    capturedAtUnixMs: 1_000,
    capturedRevision: input.capturedRevision ?? null,
    batchId: input.batchId,
    surfaceProof: input.omitSurfaceProof
      ? null
      : ({
          mode: 'typegpu-v3',
          backendStatus: 'ready',
          frameProofStatus: 'presented',
          hasPresentedFrame: true,
          hasPresentedVisibleFrame: true,
          frameProofSignature: 'frame:proof',
          presentedFrameProofSignature: 'frame:proof',
          authoritativeRevision: input.surfaceRevision ?? input.capturedRevision ?? null,
          localRevision: null,
          projectedRevision: input.capturedRevision ?? null,
          visibleRenderRevision: input.capturedRevision ?? null,
          tileSceneRevision: input.capturedRevision ?? null,
          tileSceneCameraSeq: 1,
          currentTilePaneCount: 1,
          currentHeaderPaneCount: 1,
          presentedTilePaneCount: 1,
          presentedHeaderPaneCount: 1,
          surfaceWidth: 800,
          surfaceHeight: 600,
          surfacePixelWidth: 1600,
          surfacePixelHeight: 1200,
          devicePixelRatio: 2,
          capturedAtUnixMs: 1_000,
          ...input.surfaceProofOverrides,
        } satisfies WorkbookAgentRenderedSurfaceProof),
    selection: input.useVisibleRangeOnly ? null : renderedRange,
    visibleRange: input.useVisibleRangeOnly ? renderedRange : null,
  }
}

describe('selectWorkbookRenderedReadback', () => {
  it('does not treat renderer batch ids as authoritative workbook revisions', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ batchId: 99 }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 1,
    })

    expect(proof.capturedBatchId).toBe(99)
    expect(proof.capturedRevision).toBeNull()
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
    expect(proof.incompleteReason).toContain('older than the requested verification revision')
  })

  it('requires capturedRevision to satisfy rendered freshness even when batch id is newer', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ capturedRevision: 3, batchId: 100 }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 4,
    })

    expect(proof.capturedBatchId).toBe(100)
    expect(proof.capturedRevision).toBe(3)
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
  })

  it('records whether rendered proof came from selection or visible viewport', () => {
    const selectionProof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ capturedRevision: 3, batchId: 1 }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      minRevision: 3,
    })
    const visibleRangeProof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ capturedRevision: 3, batchId: 1, useVisibleRangeOnly: true }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      minRevision: 3,
    })

    expect(selectionProof.sourceKind).toBe('selection')
    expect(visibleRangeProof.sourceKind).toBe('visibleRange')
  })

  it('requires a browser-presented TypeGPU frame proof for revision-gated readback', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ capturedRevision: 3, batchId: 3, omitSurfaceProof: true }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      minRevision: 3,
    })

    expect(proof.available).toBe(true)
    expect(proof.matched).toBeNull()
    expect(proof.surfaceProofMatched).toBeNull()
    expect(proof.incompleteReason).toContain('No browser-presented TypeGPU frame proof')
  })

  it('rejects stale presented frame proof even when the rendered cells are fresh', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({ capturedRevision: 4, batchId: 4, surfaceRevision: 3 }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      minRevision: 4,
    })

    expect(proof.stale).toBe(true)
    expect(proof.surfaceProofMatched).toBe(false)
    expect(proof.incompleteReason).toContain('Presented TypeGPU frame proof is older')
  })

  it('rejects stale visible frame lineage even when authoritative surface revision is fresh', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({
        capturedRevision: 4,
        batchId: 4,
        surfaceRevision: 4,
        surfaceProofOverrides: {
          projectedRevision: 3,
          visibleRenderRevision: 3,
          tileSceneRevision: 3,
        },
      }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      minRevision: 4,
    })

    expect(proof.stale).toBe(true)
    expect(proof.surfaceProofMatched).toBe(false)
    expect(proof.incompleteReason).toContain('visible render revision is older')
  })

  it('rejects presented proof when current and presented frame signatures diverge', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({
        capturedRevision: 4,
        batchId: 4,
        surfaceRevision: 4,
        surfaceProofOverrides: {
          frameProofSignature: 'frame:current',
          presentedFrameProofSignature: 'frame:presented',
        },
      }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      minRevision: 4,
    })

    expect(proof.stale).toBe(false)
    expect(proof.surfaceProofMatched).toBe(false)
    expect(proof.incompleteReason).toContain('does not match the current frame lineage signature')
  })

  it('rejects presented proof without current visible tile and header panes', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({
        capturedRevision: 4,
        batchId: 4,
        surfaceRevision: 4,
        surfaceProofOverrides: {
          currentTilePaneCount: 0,
        },
      }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      minRevision: 4,
    })

    expect(proof.surfaceProofMatched).toBe(false)
    expect(proof.incompleteReason).toContain('Current TypeGPU frame proof did not include visible grid tiles and headers')
  })
})
