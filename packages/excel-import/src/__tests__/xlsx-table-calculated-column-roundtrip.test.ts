import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('table calculated-column formula import/export', () => {
  it('preserves native calculated-column formulas in table metadata', () => {
    const source = exportXlsx(calculatedColumnWorkbook())
    const imported = importXlsx(source, 'table-calculated-column.xlsx')

    expect(imported.snapshot.workbook.metadata?.tables).toEqual([
      {
        name: 'SalesTable',
        sheetName: 'Sales',
        startAddress: 'A1',
        endAddress: 'C4',
        columnNames: ['Item', 'Qty', 'Total'],
        columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Total', calculatedColumnFormula: 'SalesTable[[#This Row],[Qty]]*10' }],
        headerRow: true,
        totalsRow: false,
      },
    ])

    const exportedTableXml = tableXml(exportXlsx(imported.snapshot))
    expect(exportedTableXml).toContain(
      '<tableColumn id="3" name="Total"><calculatedColumnFormula>SalesTable[[#This Row],[Qty]]*10</calculatedColumnFormula></tableColumn>',
    )
  })
})

function calculatedColumnWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'table-calculated-column',
      metadata: {
        tables: [
          {
            name: 'SalesTable',
            sheetName: 'Sales',
            startAddress: 'A1',
            endAddress: 'C4',
            columnNames: ['Item', 'Qty', 'Total'],
            columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Total', calculatedColumnFormula: 'SalesTable[[#This Row],[Qty]]*10' }],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Sales',
        order: 0,
        cells: [
          { address: 'A1', value: 'Item' },
          { address: 'B1', value: 'Qty' },
          { address: 'C1', value: 'Total' },
          { address: 'A2', value: 'A' },
          { address: 'B2', value: 2 },
          { address: 'C2', formula: 'B2*10', value: 20 },
          { address: 'A3', value: 'B' },
          { address: 'B3', value: 3 },
          { address: 'C3', formula: 'B3*10', value: 30 },
          { address: 'A4', value: 'C' },
          { address: 'B4', value: 4 },
          { address: 'C4', formula: 'B4*10', value: 40 },
        ],
      },
    ],
  }
}

function tableXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table\d+\.xml$/u.test(path))
  return tablePath ? strFromU8(zip[tablePath] ?? new Uint8Array()) : ''
}
