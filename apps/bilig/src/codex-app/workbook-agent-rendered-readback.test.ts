import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { WorkbookAgentRenderedContext } from '@bilig/contracts'
import { selectWorkbookRenderedReadback } from './workbook-agent-rendered-readback.js'

function renderedContext(input: {
  readonly capturedRevision?: number | null
  readonly batchId: number | null
  readonly sceneProof?: Partial<NonNullable<WorkbookAgentRenderedContext['visibleSceneProof']>> | null
  readonly value?: string
}): WorkbookAgentRenderedContext {
  return {
    capturedAtUnixMs: 1_000,
    capturedRevision: input.capturedRevision ?? null,
    batchId: input.batchId,
    visibleSceneProof: input.sceneProof === null ? null : visibleSceneProof(input.sceneProof ?? {}),
    selection: {
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
    },
    visibleRange: null,
  }
}

function visibleSceneProof(
  overrides: Partial<NonNullable<WorkbookAgentRenderedContext['visibleSceneProof']>>,
): NonNullable<WorkbookAgentRenderedContext['visibleSceneProof']> {
  return {
    rendererMode: 'typegpu-v3',
    frameProofStatus: 'presented',
    frameProofSignature: 'frame-1',
    presentedFrameProofSignature: 'frame-1',
    currentSceneEpochSignature: 'epoch-1',
    currentSceneOwnershipSignature: 'scene-1',
    presentedSceneEpochSignature: 'epoch-1',
    presentedSceneOwnershipSignature: 'scene-1',
    currentSceneEpoch: 'tile-1',
    presentedSceneEpoch: 'tile-1',
    currentFillHandleRevision: 'fill-1',
    presentedFillHandleRevision: 'fill-1',
    currentSelectionRevision: 'selection-1',
    presentedSelectionRevision: 'selection-1',
    currentViewportRevision: 'viewport-1',
    presentedViewportRevision: 'viewport-1',
    currentSemanticMutationRevision: '3',
    presentedSemanticMutationRevision: '3',
    currentWorkbookRevision: '3',
    presentedWorkbookRevision: '3',
    gridAuthoritativeRevision: '3',
    typeGpuAuthoritativeRevision: '3',
    visibleAuthoritativeRevision: '3',
    tileSceneRevision: 'tile-1',
    visibleRenderRevision: 'tile-1',
    hasPresentedFrame: true,
    hasPresentedVisibleFrame: true,
    frameProofMatchesPresentedFrame: true,
    visibleSceneEpochMatchesPresentedFrame: true,
    visibleSceneOwnershipMatchesPresentedFrame: true,
    visibleAuthoritativeRevisionMatchesGrid: true,
    visibleRenderRevisionMatchesTileScene: true,
    ...overrides,
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

  it('rejects stale visible-scene ownership even when rendered cells match', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({
        capturedRevision: 4,
        batchId: 4,
        sceneProof: {
          presentedSceneOwnershipSignature: 'old-scene',
          visibleSceneOwnershipMatchesPresentedFrame: false,
        },
      }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 4,
    })

    expect(proof.capturedRevision).toBe(4)
    expect(proof.visibleSceneProof.matched).toBe(false)
    expect(proof.visibleSceneProof.visibleSceneOwnershipMatchesPresentedFrame).toBe(false)
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
    expect(proof.incompleteReason).toContain('visible-scene proof')
  })

  it('rejects stale visible-scene epochs even when rendered cells match', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({
        capturedRevision: 4,
        batchId: 4,
        sceneProof: {
          presentedSceneEpochSignature: 'old-epoch',
          visibleSceneEpochMatchesPresentedFrame: false,
        },
      }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 4,
    })

    expect(proof.capturedRevision).toBe(4)
    expect(proof.visibleSceneProof.matched).toBe(false)
    expect(proof.visibleSceneProof.visibleSceneEpochMatchesPresentedFrame).toBe(false)
    expect(proof.visibleSceneProof.invalidReasons).toContain(
      'Presented visible-scene epoch does not match the current authoritative scene epoch.',
    )
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
    expect(proof.incompleteReason).toContain('visible-scene proof')
  })

  it('rejects stale semantic mutation ownership even when rendered cells match', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({
        capturedRevision: 4,
        batchId: 4,
        sceneProof: {
          presentedSemanticMutationRevision: '2',
        },
      }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 4,
    })

    expect(proof.visibleSceneProof.matched).toBe(false)
    expect(proof.visibleSceneProof.visibleSemanticMutationRevisionMatchesPresentedFrame).toBe(false)
    expect(proof.visibleSceneProof.invalidReasons).toContain(
      'Presented semantic mutation revision does not match the current authoritative scene.',
    )
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
  })

  it('rejects stale viewport, selection, and fill-handle ownership in rendered proof', () => {
    const proof = selectWorkbookRenderedReadback({
      renderedContext: renderedContext({
        capturedRevision: 4,
        batchId: 4,
        sceneProof: {
          presentedFillHandleRevision: 'fill-stale',
          presentedSelectionRevision: 'selection-stale',
          presentedViewportRevision: 'viewport-stale',
        },
      }),
      requestedRange: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      authoritativeRows: [[{ address: 'A1', input: 'ok', value: 'ok', formula: null, styleId: null, numberFormatId: null }]],
      minRevision: 4,
    })

    expect(proof.visibleSceneProof.matched).toBe(false)
    expect(proof.visibleSceneProof.visibleViewportRevisionMatchesPresentedFrame).toBe(false)
    expect(proof.visibleSceneProof.visibleSelectionRevisionMatchesPresentedFrame).toBe(false)
    expect(proof.visibleSceneProof.visibleFillHandleRevisionMatchesPresentedFrame).toBe(false)
    expect(proof.visibleSceneProof.invalidReasons).toEqual(
      expect.arrayContaining([
        'Presented viewport revision does not match the current visible scene.',
        'Presented selection revision does not match the current visible scene.',
        'Presented fill-handle revision does not match the current visible scene.',
      ]),
    )
    expect(proof.stale).toBe(true)
    expect(proof.matched).toBeNull()
  })
})
