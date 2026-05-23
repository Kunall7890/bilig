import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('table formula metadata import/export', () => {
  it('preserves calculated-column and totals-row table formulas across XLSX round trips', () => {
    const imported = importXlsx(buildTableFormulaWorkbookBytes(), 'table-formulas.xlsx')
    const table = imported.snapshot.workbook.metadata?.tables?.[0]

    expect(imported.warnings).toEqual([])
    expect(table).toEqual({
      name: 'CommissionTable',
      sheetName: 'Commissions',
      startAddress: 'A1',
      endAddress: 'D5',
      columnNames: ['Rep', 'Amount', 'Rate', 'Commission'],
      columns: [
        { name: 'Rep', totalsRowLabel: 'Total' },
        { name: 'Amount', totalsRowFunction: 'sum' },
        { name: 'Rate' },
        {
          name: 'Commission',
          calculatedColumnFormula: '[@Amount]*[@Rate]',
          totalsRowFunction: 'custom',
          totalsRowFormula: 'SUM(CommissionTable[Commission])',
        },
      ],
      headerRow: true,
      totalsRow: true,
    })

    const exportedXml = tableXml(exportXlsx(imported.snapshot))
    expect(exportedXml).toContain('<tableColumn id="1" name="Rep" totalsRowLabel="Total"/>')
    expect(exportedXml).toContain('<tableColumn id="2" name="Amount" totalsRowFunction="sum"/>')
    expect(exportedXml).toContain('<tableColumn id="3" name="Rate"/>')
    expect(exportedXml).toContain(
      '<tableColumn id="4" name="Commission" totalsRowFunction="custom"><calculatedColumnFormula>[@Amount]*[@Rate]</calculatedColumnFormula><totalsRowFormula>SUM(CommissionTable[Commission])</totalsRowFormula></tableColumn>',
    )
  })

  it('does not infer a totals row from calculated-column formulas alone', () => {
    const imported = importXlsx(buildCalculatedColumnOnlyWorkbookBytes(), 'calculated-column-only.xlsx')
    const table = imported.snapshot.workbook.metadata?.tables?.[0]

    expect(imported.warnings).toEqual([])
    expect(table).toMatchObject({
      name: 'MarginTable',
      sheetName: 'Margins',
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Amount', 'Cost', 'Margin'],
      columns: [{ name: 'Amount' }, { name: 'Cost' }, { name: 'Margin', calculatedColumnFormula: '[@Amount]-[@Cost]' }],
      headerRow: true,
      totalsRow: false,
    })

    expect(tableXml(exportXlsx(imported.snapshot))).toContain(
      '<tableColumn id="3" name="Margin"><calculatedColumnFormula>[@Amount]-[@Cost]</calculatedColumnFormula></tableColumn>',
    )
  })
})

function buildTableFormulaWorkbookBytes(): Uint8Array {
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: 'table-formulas',
      metadata: {
        tables: [
          {
            name: 'CommissionTable',
            sheetName: 'Commissions',
            startAddress: 'A1',
            endAddress: 'D5',
            columnNames: ['Rep', 'Amount', 'Rate', 'Commission'],
            headerRow: true,
            totalsRow: true,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Commissions',
        order: 0,
        cells: [
          { address: 'A1', value: 'Rep' },
          { address: 'B1', value: 'Amount' },
          { address: 'C1', value: 'Rate' },
          { address: 'D1', value: 'Commission' },
          { address: 'A2', value: 'Ada' },
          { address: 'B2', value: 1000 },
          { address: 'C2', value: 0.12 },
          { address: 'D2', formula: 'B2*C2' },
          { address: 'A3', value: 'Grace' },
          { address: 'B3', value: 1500 },
          { address: 'C3', value: 0.1 },
          { address: 'D3', formula: 'B3*C3' },
          { address: 'A4', value: 'Katherine' },
          { address: 'B4', value: 500 },
          { address: 'C4', value: 0.15 },
          { address: 'D4', formula: 'B4*C4' },
          { address: 'A5', value: 'Total' },
          { address: 'B5', formula: 'SUM(B2:B4)' },
          { address: 'D5', formula: 'SUM(D2:D4)' },
        ],
      },
    ],
  }

  return withTableColumnXml(exportXlsx(snapshot), (xml) =>
    xml
      .replace('<tableColumn id="1" name="Rep"/>', '<tableColumn id="1" name="Rep" totalsRowLabel="Total"/>')
      .replace('<tableColumn id="2" name="Amount"/>', '<tableColumn id="2" name="Amount" totalsRowFunction="sum"/>')
      .replace(
        '<tableColumn id="4" name="Commission"/>',
        '<tableColumn id="4" name="Commission" totalsRowFunction="custom"><calculatedColumnFormula>[@Amount]*[@Rate]</calculatedColumnFormula><totalsRowFormula>SUM(CommissionTable[Commission])</totalsRowFormula></tableColumn>',
      ),
  )
}

function buildCalculatedColumnOnlyWorkbookBytes(): Uint8Array {
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: 'calculated-column-only',
      metadata: {
        tables: [
          {
            name: 'MarginTable',
            sheetName: 'Margins',
            startAddress: 'A1',
            endAddress: 'C3',
            columnNames: ['Amount', 'Cost', 'Margin'],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Margins',
        order: 0,
        cells: [
          { address: 'A1', value: 'Amount' },
          { address: 'B1', value: 'Cost' },
          { address: 'C1', value: 'Margin' },
          { address: 'A2', value: 1000 },
          { address: 'B2', value: 600 },
          { address: 'C2', formula: 'A2-B2' },
          { address: 'A3', value: 800 },
          { address: 'B3', value: 450 },
          { address: 'C3', formula: 'A3-B3' },
        ],
      },
    ],
  }

  return withTableColumnXml(exportXlsx(snapshot), (xml) =>
    xml.replace(
      '<tableColumn id="3" name="Margin"/>',
      '<tableColumn id="3" name="Margin"><calculatedColumnFormula>[@Amount]-[@Cost]</calculatedColumnFormula></tableColumn>',
    ),
  )
}

function withTableColumnXml(bytes: Uint8Array, update: (xml: string) => string): Uint8Array {
  const zip = unzipSync(bytes)
  const tablePath = onlyTablePath(zip)
  zip[tablePath] = strToU8(update(strFromU8(zip[tablePath] ?? new Uint8Array())))
  return zipSync(zip)
}

function tableXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  return strFromU8(zip[onlyTablePath(zip)] ?? new Uint8Array())
}

function onlyTablePath(zip: Record<string, Uint8Array>): string {
  const tablePaths = Object.keys(zip).filter((path) => /^xl\/tables\/table[0-9]+\.xml$/u.test(path))
  expect(tablePaths).toHaveLength(1)
  return tablePaths[0]
}
