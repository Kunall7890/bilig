import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('table totals and style import/export', () => {
  it('preserves totals-row column metadata and style settings across XLSX round trips', () => {
    const sourceBytes = buildTableTotalsStyleWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'table-totals-style.xlsx')

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.tables).toEqual([
      {
        name: 'RevenueTable',
        sheetName: 'Revenue',
        startAddress: 'A1',
        endAddress: 'C5',
        columnNames: ['Account', 'Amount', 'Margin'],
        columns: [
          { name: 'Account', totalsRowLabel: 'Total' },
          { name: 'Amount', totalsRowFunction: 'sum' },
          { name: 'Margin', totalsRowFunction: 'average' },
        ],
        headerRow: true,
        totalsRow: true,
        style: {
          name: 'TableStyleMedium9',
          showFirstColumn: false,
          showLastColumn: true,
          showRowStripes: true,
          showColumnStripes: false,
        },
      },
    ])

    const exportedXml = tableXml(exportXlsx(imported.snapshot))
    expect(exportedXml).toContain('<tableColumn id="1" name="Account" totalsRowLabel="Total"/>')
    expect(exportedXml).toContain('<tableColumn id="2" name="Amount" totalsRowFunction="sum"/>')
    expect(exportedXml).toContain('<tableColumn id="3" name="Margin" totalsRowFunction="average"/>')
    expect(exportedXml).toContain(
      '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="1" showRowStripes="1" showColumnStripes="0"/>',
    )
  })
})

function buildTableTotalsStyleWorkbookBytes(): Uint8Array {
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: 'table-totals-style',
      metadata: {
        tables: [
          {
            name: 'RevenueTable',
            sheetName: 'Revenue',
            startAddress: 'A1',
            endAddress: 'C5',
            columnNames: ['Account', 'Amount', 'Margin'],
            headerRow: true,
            totalsRow: true,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [
          { address: 'A1', value: 'Account' },
          { address: 'B1', value: 'Amount' },
          { address: 'C1', value: 'Margin' },
          { address: 'A2', value: 'Services' },
          { address: 'B2', value: 1200 },
          { address: 'C2', value: 0.42 },
          { address: 'A3', value: 'Licenses' },
          { address: 'B3', value: 900 },
          { address: 'C3', value: 0.37 },
          { address: 'A4', value: 'Support' },
          { address: 'B4', value: 300 },
          { address: 'C4', value: 0.28 },
          { address: 'A5', value: 'Total' },
        ],
      },
    ],
  }
  const zip = unzipSync(exportXlsx(snapshot))
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table\d+\.xml$/u.test(path))
  if (!tablePath) {
    throw new Error('Expected exported workbook to include a table part')
  }
  const sourceTableXml = strFromU8(zip[tablePath] ?? new Uint8Array())
  zip[tablePath] = strToU8(
    sourceTableXml
      .replace('<tableColumn id="1" name="Account"/>', '<tableColumn id="1" name="Account" totalsRowLabel="Total"/>')
      .replace('<tableColumn id="2" name="Amount"/>', '<tableColumn id="2" name="Amount" totalsRowFunction="sum"/>')
      .replace('<tableColumn id="3" name="Margin"/>', '<tableColumn id="3" name="Margin" totalsRowFunction="average"/>')
      .replace(
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
        '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="1" showRowStripes="1" showColumnStripes="0"/>',
      ),
  )
  return zipSync(zip)
}

function tableXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table\d+\.xml$/u.test(path))
  return tablePath ? strFromU8(zip[tablePath] ?? new Uint8Array()) : ''
}
