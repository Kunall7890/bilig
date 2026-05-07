import { describe, expect, it } from 'vitest'
import { createFormulaBindingMemberCounts } from '../engine/services/formula-binding-member-counts.js'

describe('formula binding member counts', () => {
  it('tracks sheet and column formula counts independently', () => {
    const counts = createFormulaBindingMemberCounts()

    counts.increment(2, 5)
    counts.increment(2, 5)
    counts.increment(2, 7)

    expect(counts.countSheetMembers(2)).toBe(3)
    expect(counts.hasColumnMembers(2, 5)).toBe(true)
    expect(counts.hasColumnMembers(2, 7)).toBe(true)

    counts.decrement(2, 5)
    expect(counts.countSheetMembers(2)).toBe(2)
    expect(counts.hasColumnMembers(2, 5)).toBe(true)

    counts.decrement(2, 5)
    expect(counts.countSheetMembers(2)).toBe(1)
    expect(counts.hasColumnMembers(2, 5)).toBe(false)
    expect(counts.hasColumnMembers(2, 7)).toBe(true)
  })

  it('clears all tracked counts', () => {
    const counts = createFormulaBindingMemberCounts()
    counts.increment(1, 1)
    counts.clear()

    expect(counts.countSheetMembers(1)).toBe(0)
    expect(counts.hasColumnMembers(1, 1)).toBe(false)
  })
})
