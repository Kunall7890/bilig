import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { importXlsx } from '../index.js'

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
})
