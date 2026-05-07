import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../workbook-store.js'

describe('workbook sheet registry', () => {
  it('updates sheet id indexes and bumps the next generated id when restoring explicit ids', () => {
    const workbook = new WorkbookStore('sheet-registry')
    const sheet = workbook.createSheet('Sheet1')

    expect(sheet.id).toBe(1)
    expect(workbook.getSheetById(1)).toBe(sheet)

    const restored = workbook.createSheet('Sheet1', 0, 10)

    expect(restored).toBe(sheet)
    expect(restored.id).toBe(10)
    expect(workbook.getSheetById(1)).toBeUndefined()
    expect(workbook.getSheetById(10)).toBe(sheet)
    expect(workbook.createSheet('Next').id).toBe(11)
  })
})
