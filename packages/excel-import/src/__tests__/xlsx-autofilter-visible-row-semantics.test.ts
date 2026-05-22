import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX AutoFilter visible-row semantics', () => {
  it('derives filter-hidden row state from imported AutoFilter criteria', async () => {
    const zip = unzipSync(exportXlsx(ledgerSnapshot()))
    const sheetPath = 'xl/worksheets/sheet1.xml'
    const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
    zip[sheetPath] = strToU8(
      sheetXml.replace(
        '<autoFilter ref="A1:B6"/>',
        '<autoFilter ref="A1:B6"><filterColumn colId="0"><filters blank="0"><filter val="East"/></filters></filterColumn></autoFilter>',
      ),
    )

    const imported = importXlsx(zipSync(zip), 'autofilter-visible-row-semantics.xlsx')
    const rows = imported.snapshot.sheets[0]?.metadata?.rows

    expect(rows).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ index: 2, hidden: true, filtered: true }),
        expect.objectContaining({ index: 4, hidden: true, filtered: true }),
      ]),
    )

    const engine = new SpreadsheetEngine({ workbookName: 'autofilter-visible-row-semantics' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 90 })
  })

  it('derives numeric custom-filter visibility with Excel-style boundary rows', async () => {
    const zip = unzipSync(exportXlsx(customFilterSnapshot()))
    const sheetPath = 'xl/worksheets/sheet1.xml'
    const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
    zip[sheetPath] = strToU8(
      sheetXml.replace(
        '<autoFilter ref="A1:B7"/>',
        '<autoFilter ref="A1:B7"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="25"/></customFilters></filterColumn></autoFilter>',
      ),
    )

    const imported = importXlsx(zipSync(zip), 'autofilter-custom-filter-visible-row-semantics.xlsx')
    const rows = imported.snapshot.sheets[0]?.metadata?.rows

    expect(rows).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ index: 1, hidden: true, filtered: true }),
        expect.objectContaining({ index: 2, hidden: true, filtered: true }),
        expect.objectContaining({ index: 6, hidden: true, filtered: true }),
      ]),
    )

    const engine = new SpreadsheetEngine({ workbookName: 'autofilter-custom-filter-visible-row-semantics' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
  })
})

function ledgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'autofilter-visible-row-semantics' },
    sheets: [
      {
        id: 1,
        name: 'Ledger',
        order: 0,
        metadata: { filters: [{ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B6' }] },
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

function customFilterSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'autofilter-custom-filter-visible-row-semantics' },
    sheets: [
      {
        id: 1,
        name: 'Ledger',
        order: 0,
        metadata: { filters: [{ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B7' }] },
        cells: [
          { address: 'A1', value: 'Category' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'Low' },
          { address: 'B2', value: 10 },
          { address: 'A3', value: 'Boundary' },
          { address: 'B3', value: 25 },
          { address: 'A4', value: 'Keep 1' },
          { address: 'B4', value: 30 },
          { address: 'A5', value: 'Keep 2' },
          { address: 'B5', value: 40 },
          { address: 'A6', value: 'Text number' },
          { address: 'B6', value: '150' },
          { address: 'A7', value: 'Blank amount' },
          { address: 'B7', value: '' },
          { address: 'D1', formula: 'SUBTOTAL(9,B2:B7)', value: 105 },
          { address: 'E1', formula: 'SUBTOTAL(103,A2:A7)', value: 6 },
        ],
      },
    ],
  }
}
