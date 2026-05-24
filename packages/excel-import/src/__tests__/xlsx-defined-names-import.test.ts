import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX defined name import', () => {
  it('preserves sheet-scoped table-of-contents names without warning', () => {
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Heading']]), 'Sheet1')
    workbook.Workbook = {
      Names: [{ Name: '_Toc387132125', Sheet: 0, Ref: 'Sheet1!$A$1' }],
    }

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'toc-defined-name.xlsx')

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      {
        name: '_Toc387132125',
        scopeSheetName: 'Sheet1',
        value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' },
      },
    ])
  })

  it('preserves print titles and broken-reference names without treating them as ignored', () => {
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([['Header', '', '', '', '', 'Extent'], [], [], ['Title 1'], ['Title 2', '', '', '', '', 'Last']]),
      'Summary Info',
    )
    workbook.Workbook = {
      Names: [
        { Name: 'AgenLength', Ref: '#REF!' },
        { Name: '_xlnm.Print_Area', Sheet: 0, Ref: "'Summary Info'!$A$1:$F$5" },
        { Name: '_xlnm.Print_Titles', Sheet: 0, Ref: "'Summary Info'!$4:$5" },
      ],
    }

    const imported = importXlsx(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), 'print-defined-names.xlsx')

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: '_xlnm.Print_Area', scopeSheetName: 'Summary Info', value: { kind: 'formula', formula: "='Summary Info'!$A$1:$F$5" } },
      { name: '_xlnm.Print_Titles', scopeSheetName: 'Summary Info', value: { kind: 'formula', formula: "='Summary Info'!$4:$5" } },
      { name: 'AgenLength', value: { kind: 'formula', formula: '=#REF!' } },
    ])
  })

  it('roundtrips Excel sheet-deletion defined names as broken formulas', () => {
    const exported = exportXlsx({
      version: 1,
      workbook: {
        name: 'Sheet Deleted Defined Names',
        metadata: {
          definedNames: [
            { name: '_xlnm.Print_Area', scopeSheetName: 'Report', value: { kind: 'formula', formula: '=Report!$A$1:$A$3' } },
            { name: 'FormulaRate', value: { kind: 'formula', formula: '=#REF!' } },
            { name: 'FormulaSum', value: { kind: 'formula', formula: '=SUM(#REF!)' } },
            { name: 'RateCell', value: { kind: 'formula', formula: '=#REF!' } },
            { name: 'SalesRange', value: { kind: 'formula', formula: '=#REF!' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Report',
          order: 0,
          cells: [
            { address: 'A1', formula: 'RateCell*2' },
            { address: 'A2', formula: 'SUM(SalesRange)' },
            { address: 'A3', formula: 'FormulaRate+FormulaSum' },
          ],
        },
      ],
    })
    const workbook = XLSX.read(exported, { type: 'buffer' })

    expect(workbook.Workbook?.Names).toEqual([
      { Name: '_xlnm.Print_Area', Sheet: 0, Ref: 'Report!$A$1:$A$3' },
      { Name: 'FormulaRate', Ref: '#REF!' },
      { Name: 'FormulaSum', Ref: 'SUM(#REF!)' },
      { Name: 'RateCell', Ref: '#REF!' },
      { Name: 'SalesRange', Ref: '#REF!' },
    ])
    expect(importXlsx(exported, 'sheet-delete-defined-names.xlsx').snapshot.workbook.metadata?.definedNames).toEqual([
      { name: '_xlnm.Print_Area', scopeSheetName: 'Report', value: { kind: 'formula', formula: '=Report!$A$1:$A$3' } },
      { name: 'FormulaRate', value: { kind: 'formula', formula: '=#REF!' } },
      { name: 'FormulaSum', value: { kind: 'formula', formula: '=SUM(#REF!)' } },
      { name: 'RateCell', value: { kind: 'formula', formula: '=#REF!' } },
      { name: 'SalesRange', value: { kind: 'formula', formula: '=#REF!' } },
    ])
  })

  it('escapes structured-reference defined names with special table headers during export', () => {
    const exported = exportXlsx({
      version: 1,
      workbook: {
        name: 'Structured Defined Names',
        metadata: {
          definedNames: [{ name: 'SalesUnits', value: { kind: 'structured-ref', tableName: 'Sales', columnName: '# Units' } }],
          tables: [
            {
              name: 'Sales',
              sheetName: 'Data',
              startAddress: 'A1',
              endAddress: 'A3',
              columnNames: ['# Units'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Data',
          order: 0,
          cells: [
            { address: 'A1', value: '# Units' },
            { address: 'A2', value: 10 },
            { address: 'A3', value: 20 },
          ],
        },
      ],
    })
    const workbook = XLSX.read(exported, { type: 'buffer' })

    expect(workbook.Workbook?.Names).toEqual([{ Name: 'SalesUnits', Ref: "Sales['# Units]" }])
  })
})
