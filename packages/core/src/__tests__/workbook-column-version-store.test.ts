import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { WorkbookStore } from '../workbook-store.js'

describe('workbook column versions', () => {
  it('coalesces duplicate column writes inside a batched update', () => {
    const workbook = new WorkbookStore('column-versions')
    workbook.createSheet('Sheet1')
    const cellIndex = workbook.ensureCell('Sheet1', 'A1')

    workbook.cellStore.setValue(cellIndex, { tag: ValueTag.Number, value: 1 })
    expect(workbook.getSheetColumnVersion('Sheet1', 0)).toBe(1)
    expect(workbook.getSheetColumnVersion('Sheet1', 1)).toBe(0)

    workbook.withBatchedColumnVersionUpdates(() => {
      workbook.notifyColumnsWritten(1, new Uint32Array([0, 0, 1]))
      workbook.cellStore.setValue(cellIndex, { tag: ValueTag.Number, value: 2 })

      expect(workbook.getSheetColumnVersion('Sheet1', 0)).toBe(1)
      expect(workbook.getSheetColumnVersion('Sheet1', 1)).toBe(0)
    })

    expect(workbook.getSheetColumnVersion('Sheet1', 0)).toBe(2)
    expect(workbook.getSheetColumnVersion('Sheet1', 1)).toBe(1)
  })

  it('updates single and paired columns without materialized column arrays', () => {
    const workbook = new WorkbookStore('column-version-scalars')
    workbook.createSheet('Sheet1')

    workbook.notifyColumnWritten(1, 2)
    expect(workbook.getSheetColumnVersion('Sheet1', 2)).toBe(1)

    workbook.notifyColumnPairWritten(1, 2, 4)
    expect(workbook.getSheetColumnVersion('Sheet1', 2)).toBe(2)
    expect(workbook.getSheetColumnVersion('Sheet1', 4)).toBe(1)

    workbook.withBatchedColumnVersionUpdates(() => {
      workbook.notifyColumnWritten(1, 2)
      workbook.notifyColumnPairWritten(1, 2, 4)
      expect(workbook.getSheetColumnVersion('Sheet1', 2)).toBe(2)
      expect(workbook.getSheetColumnVersion('Sheet1', 4)).toBe(1)
    })

    expect(workbook.getSheetColumnVersion('Sheet1', 2)).toBe(3)
    expect(workbook.getSheetColumnVersion('Sheet1', 4)).toBe(2)
  })
})
