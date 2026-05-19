import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type LiteralInput } from '@bilig/protocol'
import { residentRangeShape } from '../engine/services/formula-binding-dynamic-range-bounds.js'
import { WorkbookStore } from '../workbook-store.js'
import { StringPool } from '../string-pool.js'

function setStoredCellValue(workbook: WorkbookStore, strings: StringPool, sheetName: string, address: string, value: LiteralInput): void {
  const cellIndex = workbook.ensureCell(sheetName, address)
  workbook.cellStore.setValue(cellIndex, value, value.tag === ValueTag.String ? strings.intern(value.value) : 0)
}

describe('residentRangeShape', () => {
  it('uses resident column indexes for whole-column dynamic range shapes', () => {
    const workbook = new WorkbookStore('dynamic-range-column-bounds')
    const strings = new StringPool()
    const sheet = workbook.createSheet('Sheet1')

    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', { tag: ValueTag.Number, value: 1 })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A129', { tag: ValueTag.String, value: 'last-a' })
    setStoredCellValue(workbook, strings, 'Sheet1', 'B3', { tag: ValueTag.Boolean, value: true })
    setStoredCellValue(workbook, strings, 'Sheet1', 'C1000', { tag: ValueTag.Number, value: 999 })
    const gridScan = vi.spyOn(sheet.grid, 'forEachCellEntry').mockImplementation(() => {
      throw new Error('unexpected whole-grid scan')
    })

    expect(
      residentRangeShape({
        workbook,
        ownerSheetName: 'Sheet1',
        range: { kind: 'RangeRef', refKind: 'cols', start: 'A', end: 'B' },
      }),
    ).toEqual({ rows: 129, cols: 2 })
    expect(gridScan).not.toHaveBeenCalled()

    gridScan.mockRestore()
  })

  it('uses resident row indexes for whole-row dynamic range shapes', () => {
    const workbook = new WorkbookStore('dynamic-range-row-bounds')
    const strings = new StringPool()
    const sheet = workbook.createSheet('Sheet1')

    setStoredCellValue(workbook, strings, 'Sheet1', 'B5', { tag: ValueTag.Number, value: 1 })
    setStoredCellValue(workbook, strings, 'Sheet1', 'D5', { tag: ValueTag.String, value: 'last-row-five' })
    setStoredCellValue(workbook, strings, 'Sheet1', 'J6', { tag: ValueTag.Number, value: 999 })
    const gridScan = vi.spyOn(sheet.grid, 'forEachCellEntry').mockImplementation(() => {
      throw new Error('unexpected whole-grid scan')
    })

    expect(
      residentRangeShape({
        workbook,
        ownerSheetName: 'Sheet1',
        range: { kind: 'RangeRef', refKind: 'rows', start: '5', end: '5' },
      }),
    ).toEqual({ rows: 1, cols: 4 })
    expect(gridScan).not.toHaveBeenCalled()

    gridScan.mockRestore()
  })
})
