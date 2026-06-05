import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { describe, expect, it } from 'vitest'

import { exportXlsx, importXlsx } from '../index.js'
import { readXlsxTestZipText } from './xlsx-test-helpers.js'

describe('XLSX defined name import', () => {
  it('preserves sheet-scoped table-of-contents names without warning', () => {
    const imported = importXlsx(
      writeSimpleXlsxWorkbook({
        sheets: [
          {
            name: 'Sheet1',
            cells: [{ address: 'A1', row: 0, col: 0, value: 'Heading' }],
          },
        ],
        definedNames: [{ name: '_Toc387132125', localSheetIndex: 0, formula: 'Sheet1!$A$1' }],
      }),
      'toc-defined-name.xlsx',
    )

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
    const imported = importXlsx(
      writeSimpleXlsxWorkbook({
        sheets: [
          {
            name: 'Summary Info',
            cells: [
              { address: 'A1', row: 0, col: 0, value: 'Header' },
              { address: 'F1', row: 0, col: 5, value: 'Extent' },
              { address: 'A4', row: 3, col: 0, value: 'Title 1' },
              { address: 'A5', row: 4, col: 0, value: 'Title 2' },
              { address: 'F5', row: 4, col: 5, value: 'Last' },
            ],
          },
        ],
        definedNames: [
          { name: 'AgenLength', formula: '#REF!' },
          { name: '_xlnm.Print_Area', localSheetIndex: 0, formula: "'Summary Info'!$A$1:$F$5" },
          { name: '_xlnm.Print_Titles', localSheetIndex: 0, formula: "'Summary Info'!$4:$5" },
        ],
      }),
      'print-defined-names.xlsx',
    )

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
    const workbookXml = readXlsxTestZipText(exported, 'xl/workbook.xml')

    expect(workbookXml).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">Report!$A$1:$A$3</definedName>')
    expect(workbookXml).toContain('<definedName name="FormulaRate">#REF!</definedName>')
    expect(workbookXml).toContain('<definedName name="FormulaSum">SUM(#REF!)</definedName>')
    expect(workbookXml).toContain('<definedName name="RateCell">#REF!</definedName>')
    expect(workbookXml).toContain('<definedName name="SalesRange">#REF!</definedName>')
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
    const workbookXml = readXlsxTestZipText(exported, 'xl/workbook.xml')

    expect(workbookXml).toContain('<definedName name="SalesUnits">Sales[\'# Units]</definedName>')
  })
})
