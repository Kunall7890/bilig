import { describe, expect, it } from 'vitest'
import { createCellNumberFormatRecord } from '@bilig/protocol'
import { buildMutationMetadataInverseOps } from '../engine/services/mutation-inverse-metadata-ops.js'
import { WorkbookStore } from '../workbook-store.js'

const styledRange = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } as const

function createWorkbookWithSheet(): WorkbookStore {
  const workbook = new WorkbookStore('inverse-metadata')
  workbook.createSheet('Sheet1')
  return workbook
}

describe('mutation inverse metadata ops', () => {
  it('leaves destructive structural and cell inverse ops to the mutation service', () => {
    const workbook = createWorkbookWithSheet()

    expect(buildMutationMetadataInverseOps(workbook, { kind: 'deleteSheet', name: 'Sheet1' })).toBeUndefined()
    expect(buildMutationMetadataInverseOps(workbook, { kind: 'deleteRows', sheetName: 'Sheet1', start: 1, count: 2 })).toBeUndefined()
    expect(buildMutationMetadataInverseOps(workbook, { kind: 'deleteColumns', sheetName: 'Sheet1', start: 1, count: 2 })).toBeUndefined()
    expect(
      buildMutationMetadataInverseOps(workbook, { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }),
    ).toBeUndefined()
  })

  it('builds inverse ops for sheet-level metadata from current workbook state', () => {
    const workbook = createWorkbookWithSheet()
    workbook.setFreezePane('Sheet1', 1, 2)
    workbook.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' })
    workbook.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })

    expect(buildMutationMetadataInverseOps(workbook, { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 3, cols: 1 })).toEqual([
      { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 2 },
    ])
    expect(
      buildMutationMetadataInverseOps(workbook, {
        kind: 'clearFilter',
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' },
      }),
    ).toEqual([{ kind: 'setFilter', sheetName: 'Sheet1', range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' } }])
    expect(buildMutationMetadataInverseOps(workbook, { kind: 'deleteTable', name: 'Sales' })).toEqual([
      {
        kind: 'upsertTable',
        table: {
          name: 'Sales',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B4',
          columnNames: ['Region', 'Amount'],
          headerRow: true,
          totalsRow: false,
        },
      },
    ])
  })

  it('restores style and number-format range metadata from current workbook tiles', () => {
    const workbook = createWorkbookWithSheet()
    workbook.upsertCellStyle({ id: 'style-accent', fill: { backgroundColor: '#dbeafe' }, font: { bold: true } })
    workbook.setStyleRange(styledRange, 'style-accent')
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-decimal', '0.00'))
    workbook.setFormatRange(styledRange, 'format-decimal')

    expect(buildMutationMetadataInverseOps(workbook, { kind: 'setStyleRange', range: styledRange, styleId: 'style-next' })).toEqual([
      { kind: 'upsertCellStyle', style: { id: 'style-accent', fill: { backgroundColor: '#dbeafe' }, font: { bold: true } } },
      { kind: 'setStyleRange', range: styledRange, styleId: 'style-accent' },
    ])
    expect(buildMutationMetadataInverseOps(workbook, { kind: 'setFormatRange', range: styledRange, formatId: 'format-next' })).toEqual([
      { kind: 'upsertCellNumberFormat', format: createCellNumberFormatRecord('format-decimal', '0.00') },
      { kind: 'setFormatRange', range: styledRange, formatId: 'format-decimal' },
    ])
  })
})
