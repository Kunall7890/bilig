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

  it('bumps imported style and format ids without accepting unrelated ids', () => {
    const allocator = new WorkbookIdAllocator()

    allocator.bumpStyleId('not-a-style')
    allocator.bumpStyleId('style-12')
    allocator.bumpStyleId('style-3')
    allocator.bumpFormatId('not-a-format')
    allocator.bumpFormatId('format-7')
    allocator.bumpFormatId('format-2')

    expect(allocator.createAxisEntry('row').id).toBe('row-1')
    expect(allocator.createLogicalAxisId('column')).toBe('lc1')
    expect(allocator.createAxisEntry('column')).toBeDefined()
    expect(allocator.createLogicalAxisId('row')).toBe('lr1')
  })
})
