import { describe, expect, it } from 'vitest'
import { createInitialFormulaCellMembership } from '../engine/services/formula-initialization-membership.js'

describe('createInitialFormulaCellMembership', () => {
  it('tracks small sparse formula batches without losing cell zero', () => {
    const membership = createInitialFormulaCellMembership({
      cellIndices: new Uint32Array([0, 500, 505]),
      maxCellIndex: 505,
    })

    expect(membership.has(0)).toBe(true)
    expect(membership.has(500)).toBe(true)
    expect(membership.has(504)).toBe(false)

    membership.delete(0)
    expect(membership.has(0)).toBe(false)
    membership.add(0)
    expect(membership.has(0)).toBe(true)
  })

  it('supports dense low-index memberships', () => {
    const membership = createInitialFormulaCellMembership({
      cellIndices: new Uint32Array([1, 3, 5]),
      maxCellIndex: 5,
      expectedCellCount: 64,
    })

    expect(membership.has(3)).toBe(true)
    membership.delete(3)
    expect(membership.has(3)).toBe(false)
    membership.add(7)
    expect(membership.has(7)).toBe(true)
  })

  it('supports large sparse memberships', () => {
    const membership = createInitialFormulaCellMembership({
      cellIndices: new Uint32Array([10, 10_000, 90_000]),
      maxCellIndex: 90_000,
      expectedCellCount: 128,
    })

    expect(membership.has(10_000)).toBe(true)
    membership.delete(10_000)
    expect(membership.has(10_000)).toBe(false)
    membership.add(75_000)
    expect(membership.has(75_000)).toBe(true)
  })
})
