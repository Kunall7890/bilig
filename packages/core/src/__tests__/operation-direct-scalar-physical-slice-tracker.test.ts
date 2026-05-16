import { describe, expect, it } from 'vitest'

import { DirectScalarPhysicalSliceTracker } from '../engine/services/operation-direct-scalar-physical-slice-tracker.js'

describe('DirectScalarPhysicalSliceTracker', () => {
  it('trusts sorted input and formula slices on the same sheet', () => {
    const tracker = new DirectScalarPhysicalSliceTracker()

    tracker.noteCell(7, 0, 0, 'input')
    tracker.noteCell(7, 1, 0, 'input')
    tracker.noteCell(7, 0, 1, 'formula')
    tracker.noteCell(7, 1, 1, 'formula')

    expect(tracker.getTrustedSheetIdForTrackedChanges(2, 4)).toBe(7)
  })

  it('rejects cross-sheet tracked slices', () => {
    const tracker = new DirectScalarPhysicalSliceTracker()

    tracker.noteCell(7, 0, 0, 'input')
    tracker.noteCell(8, 0, 1, 'formula')

    expect(tracker.getTrustedSheetIdForTrackedChanges(1, 2)).toBeUndefined()
  })

  it('rejects out-of-order formula slices without rejecting sorted inputs separately', () => {
    const tracker = new DirectScalarPhysicalSliceTracker()

    tracker.noteCell(7, 0, 0, 'input')
    tracker.noteCell(7, 1, 0, 'input')
    tracker.noteCell(7, 2, 1, 'formula')
    tracker.noteCell(7, 1, 1, 'formula')

    expect(tracker.getTrustedSheetIdForTrackedChanges(2, 4)).toBeUndefined()
  })

  it('requires a strict split between explicit inputs and formula outputs', () => {
    const tracker = new DirectScalarPhysicalSliceTracker()

    tracker.noteCell(7, 0, 0, 'input')

    expect(tracker.getTrustedSheetIdForTrackedChanges(0, 1)).toBeUndefined()
    expect(tracker.getTrustedSheetIdForTrackedChanges(1, 1)).toBeUndefined()
  })
})
