import { describe, expect, it } from 'vitest'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import type { ViewportPatch } from '@bilig/worker-transport'
import { ProjectedViewportPatchRevisionGate } from '../projected-viewport-patch-revision-gate.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function patch(input: { authoritativeRevision?: number; batchId?: number } = {}): ViewportPatch {
  return {
    version: 1,
    full: false,
    ...(input.authoritativeRevision !== undefined ? { authoritativeRevision: input.authoritativeRevision } : {}),
    freezeRows: 0,
    freezeCols: 0,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 0,
      colStart: 0,
      colEnd: 0,
    },
    metrics: {
      ...TEST_METRICS,
      batchId: input.batchId ?? 0,
    },
    styles: [],
    cells: [
      {
        row: 0,
        col: 0,
        snapshot: {
          sheetName: 'Sheet1',
          address: 'A1',
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 0,
        },
        displayText: '',
        copyText: '',
        editorText: '',
        formatId: 0,
        styleId: 'style-0',
      },
    ],
    columns: [],
    rows: [],
  }
}

describe('ProjectedViewportPatchRevisionGate', () => {
  it('rejects older authoritative revisions after a newer patch is applied', () => {
    const gate = new ProjectedViewportPatchRevisionGate()

    const newer = patch({ authoritativeRevision: 17, batchId: 23 })
    expect(gate.shouldApplyViewportPatch(newer)).toBe(true)
    gate.noteAppliedViewportPatch(newer)

    expect(gate.shouldApplyViewportPatch(patch({ authoritativeRevision: 16, batchId: 24 }))).toBe(false)
    expect(gate.getLastAuthoritativeRevision()).toBe(17)
    expect(gate.getLastBatchId()).toBe(23)
  })

  it('accepts a higher authoritative revision even when the recalc batch id is lower', () => {
    const gate = new ProjectedViewportPatchRevisionGate()

    gate.noteAppliedViewportPatch(patch({ authoritativeRevision: 17, batchId: 23 }))

    const newerAuthoritativePatch = patch({ authoritativeRevision: 18, batchId: 21 })
    expect(gate.shouldApplyViewportPatch(newerAuthoritativePatch)).toBe(true)
    gate.noteAppliedViewportPatch(newerAuthoritativePatch)

    expect(gate.getLastAuthoritativeRevision()).toBe(18)
    expect(gate.getLastBatchId()).toBe(23)
  })

  it('rejects lower projected batch ids at the same authoritative revision', () => {
    const gate = new ProjectedViewportPatchRevisionGate()

    gate.noteAppliedViewportPatch(patch({ authoritativeRevision: 17, batchId: 23 }))

    expect(gate.shouldApplyViewportPatch(patch({ authoritativeRevision: 17, batchId: 22 }))).toBe(false)
    expect(gate.shouldApplyViewportPatch(patch({ authoritativeRevision: 17, batchId: 23 }))).toBe(true)
  })

  it('continues to accept legacy direct patches without revision metadata', () => {
    const gate = new ProjectedViewportPatchRevisionGate()

    gate.noteAppliedViewportPatch(patch({ authoritativeRevision: 17, batchId: 23 }))

    expect(gate.shouldApplyViewportPatch(patch())).toBe(true)
  })
})
