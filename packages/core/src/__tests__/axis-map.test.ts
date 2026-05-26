import { describe, expect, it } from 'vitest'
import { AxisMap } from '../storage/axis-map.js'

describe('AxisMap', () => {
  it('preserves sparse insert holes and stable ids at visible positions', () => {
    const axisMap = new AxisMap()

    const removed = axisMap.splice(0, 0, 3, [
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])

    expect(removed).toEqual([])
    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])
    expect(axisMap.snapshot(0, 3)).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-c', index: 2 },
    ])
    expect(axisMap.getId(1)).toBeUndefined()
    expect(axisMap.indexOf('row-c')).toBe(2)
  })

  it('replaces, snapshots, splices, and moves axis ids by visible index', () => {
    const axisMap = new AxisMap()

    axisMap.replaceRange(0, [
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
      { id: 'row-c', index: 2 },
    ])

    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
      { id: 'row-c', index: 2 },
    ])
    expect(axisMap.snapshot(1, 2)).toEqual([
      { id: 'row-b', index: 1 },
      { id: 'row-c', index: 2 },
    ])

    const removed = axisMap.splice(1, 1, 1, [{ id: 'row-x', index: 1 }])
    expect(removed).toEqual([{ id: 'row-b', index: 1 }])
    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-x', index: 1 },
      { id: 'row-c', index: 2 },
    ])

    axisMap.move(0, 1, 3)
    expect(axisMap.list()).toEqual([
      { id: 'row-x', index: 0 },
      { id: 'row-c', index: 1 },
      { id: 'row-a', index: 2 },
    ])
  })

  it('ignores holes when snapshotting or listing', () => {
    const axisMap = new AxisMap()

    axisMap.replaceRange(2, [{ id: 'column-c', index: 2 }])

    expect(axisMap.snapshot(0, 3)).toEqual([{ id: 'column-c', index: 2 }])
    expect(axisMap.list()).toEqual([{ id: 'column-c', index: 2 }])
  })

  it('ensures stable ids for visible positions', () => {
    const axisMap = new AxisMap()

    expect(axisMap.ensureId(2, () => 'row-a')).toBe('row-a')
    expect(axisMap.getId(2)).toBe('row-a')
    expect(axisMap.ensureId(2, () => 'row-b')).toBe('row-a')
    expect(axisMap.indexOf('row-a')).toBe(2)
    expect(axisMap.indexOf('row-missing')).toBe(-1)
  })

  it('ensures dense ids from a batched allocator while preserving existing ids', () => {
    const axisMap = new AxisMap()
    let nextId = 1
    const createIds = (count: number): string[] =>
      Array.from({ length: count }, () => {
        const id = `row-${nextId}`
        nextId += 1
        return id
      })

    expect(axisMap.ensureDenseIdsFrom(0, 3, createIds)).toEqual(['row-1', 'row-2', 'row-3'])
    expect(axisMap.ensureDenseIdsFrom(1, 3, createIds)).toEqual(['row-2', 'row-3', 'row-4'])
    expect(axisMap.list()).toEqual([
      { id: 'row-1', index: 0 },
      { id: 'row-2', index: 1 },
      { id: 'row-3', index: 2 },
      { id: 'row-4', index: 3 },
    ])
    expect(axisMap.indexOf('row-4')).toBe(3)
  })

  it('reports length and no-ops for empty snapshots or moves', () => {
    const axisMap = new AxisMap()

    expect(axisMap.length).toBe(0)
    expect(axisMap.snapshot(0, 0)).toEqual([])

    axisMap.replaceRange(0, [
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
    ])

    expect(axisMap.length).toBe(2)

    axisMap.move(0, 0, 1)
    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
    ])

    axisMap.move(1, 1, 1)
    expect(axisMap.list()).toEqual([
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
    ])
  })

  it('keeps empty sparse splices as no-ops and updates shifted suffix ids incrementally', () => {
    const axisMap = new AxisMap()

    expect(axisMap.splice(50, 1, 0, [])).toEqual([])
    expect(axisMap.length).toBe(0)

    axisMap.replaceRange(2, [{ id: 'row-c', index: 2 }])

    expect(axisMap.splice(1, 1, 0, [])).toEqual([])
    expect(axisMap.list()).toEqual([{ id: 'row-c', index: 1 }])
    expect(axisMap.indexOf('row-c')).toBe(1)
  })

  it('uses the single-entry insert path while shifting suffix ids', () => {
    const axisMap = new AxisMap()
    axisMap.replaceRange(0, [
      { id: 'column-a', index: 0 },
      { id: 'column-b', index: 1 },
      { id: 'column-c', index: 2 },
    ])

    expect(axisMap.splice(1, 0, 1, [{ id: 'column-x', index: 1 }])).toEqual([])

    expect(axisMap.list()).toEqual([
      { id: 'column-a', index: 0 },
      { id: 'column-x', index: 1 },
      { id: 'column-b', index: 2 },
      { id: 'column-c', index: 3 },
    ])
    expect(axisMap.indexOf('column-b')).toBe(2)
    expect(axisMap.indexOf('column-c')).toBe(3)
  })
})
