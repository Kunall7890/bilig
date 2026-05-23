import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { SpreadsheetEngine } from '@bilig/core'
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

  it('exports engine-produced table sorts as table metadata that reimports with sorted formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-table-sort-state' })
    await engine.ready()
    engine.importSnapshot(tableLedgerSnapshot())

    expect(engine.sortTable('Ledger', 'Sales', [{ keyAddress: 'B1', direction: 'desc' }])).toBe(true)

    const exported = exportXlsx(engine.exportSnapshot())
    const exportedXml = tableXml(exported)
    expect(exportedXml).toContain('<sortState ref="A2:D6">')
    expect(exportedXml).toContain('<sortCondition descending="1" ref="B2:B6"/>')

    const imported = importXlsx(exported, 'engine-table-sort-state.xlsx')
    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
      name: 'Sales',
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'D6',
      sortState: '<sortState ref="A2:D6"><sortCondition descending="1" ref="B2:B6"/></sortState>',
    })
    const rehydrated = new SpreadsheetEngine({ workbookName: 'engine-table-sort-state-reimport' })
    await rehydrated.ready()
    rehydrated.importSnapshot(imported.snapshot)
    expect(engineRows(rehydrated)).toEqual([
      ['East', 50, 'invoice-005', 100],
      ['West', 40, 'invoice-002', 80],
      ['East', 30, 'invoice-003', 60],
      ['West', 20, 'invoice-004', 40],
      ['East', 10, 'invoice-001', 20],
    ])
    expect(snapshotFormulas(imported.snapshot)).toEqual([
      { address: 'D2', formula: 'B2*2' },
      { address: 'D3', formula: 'B3*2' },
      { address: 'D4', formula: 'B4*2' },
      { address: 'D5', formula: 'B5*2' },
      { address: 'D6', formula: 'B6*2' },
    ])
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

function tableLedgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Engine table sort state',
      metadata: {
        tables: [
          {
            name: 'Sales',
            sheetName: 'Ledger',
            startAddress: 'A1',
            endAddress: 'D6',
            columnNames: ['Region', 'Amount', 'Invoice', 'Double'],
            columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Invoice' }, { name: 'Double' }],
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
          { address: 'C1', value: 'Invoice' },
          { address: 'D1', value: 'Double' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 10 },
          { address: 'C2', value: 'invoice-001' },
          { address: 'D2', formula: 'B2*2', value: 20 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 40 },
          { address: 'C3', value: 'invoice-002' },
          { address: 'D3', formula: 'B3*2', value: 80 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 30 },
          { address: 'C4', value: 'invoice-003' },
          { address: 'D4', formula: 'B4*2', value: 60 },
          { address: 'A5', value: 'West' },
          { address: 'B5', value: 20 },
          { address: 'C5', value: 'invoice-004' },
          { address: 'D5', formula: 'B5*2', value: 40 },
          { address: 'A6', value: 'East' },
          { address: 'B6', value: 50 },
          { address: 'C6', value: 'invoice-005' },
          { address: 'D6', formula: 'B6*2', value: 100 },
        ],
      },
    ],
  }
}

function engineRows(engine: SpreadsheetEngine): unknown[][] {
  return engine.getRangeValues({ sheetName: 'Ledger', startAddress: 'A2', endAddress: 'D6' }).map((row) =>
    row.map((value) => {
      if ('value' in value) {
        return value.value
      }
      return null
    }),
  )
}

function snapshotFormulas(snapshot: WorkbookSnapshot): readonly { readonly address: string; readonly formula: string }[] {
  const formulas = new Map(snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell.formula]) ?? [])
  return ['D2', 'D3', 'D4', 'D5', 'D6'].map((address) => {
    const formula = formulas.get(address)
    if (!formula) {
      throw new Error(`Missing formula at ${address}`)
    }
    return { address, formula }
  })
}
