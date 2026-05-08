import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('table sort state import/export', () => {
  it('preserves table-level sortState metadata across XLSX round trips', () => {
    const sourceBytes = buildTableSortStateWorkbookBytes()
    const imported = importXlsx(sourceBytes, 'table-sort-state.xlsx')

    expect(imported.warnings).toEqual([])
    expect(tableXml(sourceBytes)).toContain('<sortState ')

    const exportedXml = tableXml(exportXlsx(imported.snapshot))
    expect(exportedXml).toContain('<autoFilter ref="A1:B4"/>')
    expect(exportedXml).toContain('<sortState ref="A2:B4">')
    expect(exportedXml).toContain('<sortCondition descending="1" ref="B2:B4"/>')
  })
})

function buildTableSortStateWorkbookBytes(): Uint8Array {
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: 'table-sort-state',
      metadata: {
        tables: [
          {
            name: 'SalesTable',
            sheetName: 'Sales',
            startAddress: 'A1',
            endAddress: 'B4',
            columnNames: ['Region', 'Revenue'],
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
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Revenue' },
          { address: 'A2', value: 'West' },
          { address: 'B2', value: 30 },
          { address: 'A3', value: 'East' },
          { address: 'B3', value: 20 },
          { address: 'A4', value: 'Central' },
          { address: 'B4', value: 10 },
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
    sourceTableXml.replace(
      /<tableColumns\b/u,
      '<sortState ref="A2:B4"><sortCondition descending="1" ref="B2:B4"/></sortState><tableColumns',
    ),
  )
  return zipSync(zip)
}

function tableXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table\d+\.xml$/u.test(path))
  return tablePath ? strFromU8(zip[tablePath] ?? new Uint8Array()) : ''
}
