import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX table AutoFilter criteria roundtrip', () => {
  it('preserves table-scoped criteria and derives filtered row visibility', async () => {
    const sourceZip = unzipSync(exportXlsx(tableLedgerSnapshot()))
    const tablePath = onlyTablePath(sourceZip)
    const tableXml = strFromU8(sourceZip[tablePath] ?? new Uint8Array())
    sourceZip[tablePath] = strToU8(
      tableXml.replace(
        '<autoFilter ref="A1:B6"/>',
        '<autoFilter ref="A1:B6"><filterColumn colId="0"><filters blank="0"><filter val="East"/></filters></filterColumn></autoFilter>',
      ),
    )

    const imported = importXlsx(zipSync(sourceZip), 'table-autofilter-criteria.xlsx')

    expect(imported.snapshot.workbook.metadata?.tables?.[0]?.autoFilter).toEqual({
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'B6',
      criteria: [{ colId: 0, filters: { blank: false, values: ['East'] } }],
    })
    expect(imported.snapshot.sheets[0]?.metadata?.filters).toBeUndefined()
    expect(imported.snapshot.sheets[0]?.metadata?.rows).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ index: 2, hidden: true, filtered: true }),
        expect.objectContaining({ index: 4, hidden: true, filtered: true }),
      ]),
    )

    const engine = new SpreadsheetEngine({ workbookName: 'table-autofilter-criteria-roundtrip' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 90 })

    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const exportedTableXml = strFromU8(exportedZip[onlyTablePath(exportedZip)] ?? new Uint8Array())
    const exportedSheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(exportedTableXml).toContain('<autoFilter ref="A1:B6">')
    expect(exportedTableXml).toContain('<filterColumn colId="0"><filters blank="0"><filter val="East"/></filters></filterColumn>')
    expect(exportedSheetXml).not.toContain('<autoFilter')
  })
})

function tableLedgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'table-autofilter-criteria-roundtrip',
      metadata: {
        tables: [
          {
            name: 'LedgerTable',
            sheetName: 'Ledger',
            startAddress: 'A1',
            endAddress: 'B6',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Ledger',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 10 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 20 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 30 },
          { address: 'A5', value: 'West' },
          { address: 'B5', value: 40 },
          { address: 'A6', value: 'East' },
          { address: 'B6', value: 50 },
          { address: 'D1', formula: 'SUBTOTAL(9,B2:B6)', value: 150 },
          { address: 'E1', formula: 'SUBTOTAL(109,B2:B6)', value: 100 },
          { address: 'F1', formula: 'AGGREGATE(9,4,B2:B6)', value: 150 },
          { address: 'G1', formula: 'AGGREGATE(9,5,B2:B6)', value: 100 },
        ],
      },
    ],
  }
}

function onlyTablePath(zip: Record<string, Uint8Array>): string {
  const tablePaths = Object.keys(zip).filter((path) => /^xl\/tables\/table[0-9]+\.xml$/u.test(path))
  expect(tablePaths).toHaveLength(1)
  return tablePaths[0]
}
