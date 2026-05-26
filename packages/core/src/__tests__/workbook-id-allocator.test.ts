import { describe, expect, it } from 'vitest'
import { WorkbookIdAllocator } from '../workbook-id-allocator.js'

describe('WorkbookIdAllocator', () => {
  it('allocates axis and logical ids independently and resets them together', () => {
    const allocator = new WorkbookIdAllocator()

    expect(allocator.createAxisEntry('row').id).toBe('row-1')
    expect(allocator.createAxisEntry('column').id).toBe('column-1')
    expect(allocator.createLogicalAxisId('row')).toBe('lr1')
    expect(allocator.createLogicalAxisId('column')).toBe('lc1')

    allocator.reset()

    expect(allocator.createAxisEntry('row').id).toBe('row-1')
    expect(allocator.createAxisEntry('column').id).toBe('column-1')
    expect(allocator.createLogicalAxisId('row')).toBe('lr1')
    expect(allocator.createLogicalAxisId('column')).toBe('lc1')
  })

  it('allocates dense logical axis id runs from the same sequence', () => {
    const allocator = new WorkbookIdAllocator()

    expect(allocator.createLogicalAxisIds('row', 3)).toEqual(['lr1', 'lr2', 'lr3'])
    expect(allocator.createLogicalAxisId('row')).toBe('lr4')
    expect(allocator.createLogicalAxisIds('column', 2)).toEqual(['lc1', 'lc2'])
    expect(allocator.createLogicalAxisId('column')).toBe('lc3')
  })
})
