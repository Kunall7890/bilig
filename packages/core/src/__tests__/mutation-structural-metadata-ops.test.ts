import { describe, expect, it } from 'vitest'
import { createCellNumberFormatRecord } from '@bilig/protocol'
import { WorkbookStore } from '../workbook-store.js'
import {
  captureStructuralWorkbookMetadataOps,
  clearStructuralSheetMetadataOps,
} from '../engine/services/mutation-structural-metadata-ops.js'

function createWorkbookWithMetadata(): WorkbookStore {
  const workbook = new WorkbookStore('metadata-ops')
  workbook.createSheet('Sheet1')
  workbook.setDefinedName('Total', { kind: 'formula', formula: 'SUM(Sheet1!A1:A3)' })
  workbook.setTable({
    name: 'Sales',
    sheetName: 'Sheet1',
    startAddress: 'A1',
    endAddress: 'B3',
    columnNames: ['Amount', 'Total'],
    headerRow: true,
    totalsRow: false,
  })
  workbook.upsertCellStyle({ id: 'style-bold', font: { bold: true } })
  workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, 'style-bold')
  workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-decimal', '0.00'))
  workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' }, 'format-decimal')
  workbook.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' })
  workbook.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [{ keyAddress: 'B1', direction: 'asc' }])
  workbook.setDataValidation({
    range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' },
    rule: { kind: 'list', values: ['Draft', 'Final'] },
    allowBlank: false,
  })
  workbook.setCommentThread({
    threadId: 'thread-1',
    sheetName: 'Sheet1',
    address: 'A1',
    comments: [{ id: 'comment-1', body: 'Check this total.' }],
  })
  workbook.setNote({ sheetName: 'Sheet1', address: 'B1', text: 'Manual override' })
  return workbook
}

describe('mutation structural metadata ops', () => {
  it('captures workbook-level structural metadata as restore operations', () => {
    const ops = captureStructuralWorkbookMetadataOps(createWorkbookWithMetadata())

    expect(ops).toContainEqual({ kind: 'upsertDefinedName', name: 'Total', value: { kind: 'formula', formula: 'SUM(Sheet1!A1:A3)' } })
    expect(ops).toContainEqual({
      kind: 'upsertTable',
      table: {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B3',
        columnNames: ['Amount', 'Total'],
        headerRow: true,
        totalsRow: false,
      },
    })
  })

  it('builds clear operations for sheet metadata that survives a structural transform', () => {
    const ops = clearStructuralSheetMetadataOps(createWorkbookWithMetadata(), 'Sheet1', {
      kind: 'insert',
      axis: 'row',
      start: 0,
      count: 1,
    })

    expect(ops).toContainEqual({
      kind: 'setStyleRange',
      range: { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'B3' },
      styleId: WorkbookStore.defaultStyleId,
    })
    expect(ops).toContainEqual({
      kind: 'setFormatRange',
      range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C3' },
      formatId: WorkbookStore.defaultFormatId,
    })
    expect(ops).toContainEqual({
      kind: 'clearFilter',
      sheetName: 'Sheet1',
      range: { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'B4' },
    })
    expect(ops).toContainEqual({
      kind: 'clearSort',
      sheetName: 'Sheet1',
      range: { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'B4' },
    })
    expect(ops).toContainEqual({
      kind: 'clearDataValidation',
      sheetName: 'Sheet1',
      range: { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'A4' },
    })
    expect(ops).toContainEqual({ kind: 'deleteCommentThread', sheetName: 'Sheet1', address: 'A2' })
    expect(ops).toContainEqual({ kind: 'deleteNote', sheetName: 'Sheet1', address: 'B2' })
  })
})
