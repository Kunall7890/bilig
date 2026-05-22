import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }
const expectedTableMetadata = [
  {
    name: 'SalesTable',
    sheetName: 'Sales',
    startAddress: 'A1',
    endAddress: 'C5',
    columnNames: ['Item', 'Qty', 'Total'],
    columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Total', calculatedColumnFormula: 'SalesTable[[#This Row],[Qty]]*10' }],
    headerRow: true,
    totalsRow: false,
  },
]

describe('macOS Desktop Excel table calculated-column oracle', () => {
  it('uses native calculated-column metadata to fill inserted table rows in headless', () => {
    const imported = importXlsx(exportXlsx(calculatedColumnWorkbook()), 'table-calculated-column-source.xlsx')
    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      const sheetId = workbook.getSheetId('Sales')
      if (sheetId === undefined) {
        throw new Error('Sales sheet is missing')
      }
      workbook.addRows(sheetId, 2, 1)
      workbook.setCellContents(cell(2, 0), 'New')
      workbook.setCellContents(cell(2, 1), 5)

      expect(workbook.getCellFormula(addressToCell('C3'))).toBe("='Sales'!B3*10")
      expect(workbook.getCellValue(addressToCell('C3'))).toMatchObject({ value: 50 })

      const roundTrip = importXlsx(exportXlsx(workbook.exportSnapshot()), 'table-calculated-column-roundtrip.xlsx')
      expect(roundTrip.snapshot.workbook.metadata?.tables).toEqual(expectedTableMetadata)
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel calculated-column values and metadata after inserting a table row',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-calculated-column-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-table-calculated-column-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(calculatedColumnWorkbook()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Sales',
          operations: [
            { kind: 'insertRows', range: '3:3' },
            { kind: 'setCellValue', address: 'A3', value: 'New' },
            { kind: 'setCellValue', address: 'B3', value: 5 },
          ],
          inspectCells: ['C2', 'C3', 'C4', 'C5'],
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'C2', value: { kind: 'number', value: 20 } },
          { address: 'C3', value: { kind: 'number', value: 50 } },
          { address: 'C4', value: { kind: 'number', value: 30 } },
          { address: 'C5', value: { kind: 'number', value: 40 } },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-table-calculated-column-oracle.xlsx')
        expect(excelTruth.snapshot.workbook.metadata?.tables).toEqual(expectedTableMetadata)

        const workbook = WorkPaper.buildFromSnapshot(
          importXlsx(exportXlsx(calculatedColumnWorkbook()), 'headless-source.xlsx').snapshot,
          workbookConfig,
        )
        try {
          const sheetId = workbook.getSheetId('Sales')
          if (sheetId === undefined) {
            throw new Error('Sales sheet is missing')
          }
          workbook.addRows(sheetId, 2, 1)
          workbook.setCellContents(cell(2, 0), 'New')
          workbook.setCellContents(cell(2, 1), 5)

          const headlessWorkbookPath = join(tempDir, 'headless-table-calculated-column-oracle.xlsx')
          writeFileSync(headlessWorkbookPath, exportXlsx(workbook.exportSnapshot()))
          const headlessExcel = runMacosExcelStructuralOperationOracle({
            workbookPath: headlessWorkbookPath,
            worksheetName: 'Sales',
            operations: [{ kind: 'setCellValue', address: 'B3', value: 5 }],
            inspectCells: ['C2', 'C3', 'C4', 'C5'],
            saveWorkbook: true,
          })
          expect(headlessExcel.cells.map(({ address, value }) => ({ address, value }))).toEqual(
            excelResult.cells.map(({ address, value }) => ({ address, value })),
          )

          const headlessImport = importXlsx(
            new Uint8Array(readFileSync(headlessWorkbookPath)),
            'headless-table-calculated-column-oracle.xlsx',
          )
          expect(headlessImport.snapshot.workbook.metadata?.tables).toEqual(expectedTableMetadata)
        } finally {
          workbook.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    60_000,
  )
})

function calculatedColumnWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Table calculated column',
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

function cell(row: number, col: number): WorkPaperCellAddress {
  return { sheet: 1, row, col }
}

function addressToCell(address: string): WorkPaperCellAddress {
  const match = /^([A-Z]+)([1-9][0-9]*)$/u.exec(address)
  if (!match) {
    throw new Error(`Unexpected oracle address: ${address}`)
  }
  let col = 0
  for (const char of match[1]) {
    col = col * 26 + char.charCodeAt(0) - 64
  }
  return cell(Number(match[2]) - 1, col - 1)
}
