import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { WorkPaper } from '../work-paper.js'

describe('macOS Desktop Excel data validation oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'imports Excel-saved validations and blocks invalid headless writes',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-data-validation-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-data-validation-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(validationWorkbook()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Order',
          operations: [
            { kind: 'setCellValue', address: 'B2', value: 'East' },
            { kind: 'setCellValue', address: 'C2', value: 0.25 },
          ],
          inspectCells: ['B2', 'C2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'B2', value: { kind: 'string', value: 'East' } },
          { address: 'C2', value: { kind: 'number', value: 0.25 } },
        ])

        const imported = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-data-validation-oracle.xlsx')
        expect(imported.snapshot.sheets.find((sheet) => sheet.name === 'Order')?.metadata?.validations).toEqual([
          expect.objectContaining({
            range: { sheetName: 'Order', startAddress: 'B2', endAddress: 'B2' },
            rule: {
              kind: 'list',
              source: {
                kind: 'named-range',
                name: 'Regions',
              },
            },
          }),
          expect.objectContaining({
            range: { sheetName: 'Order', startAddress: 'C2', endAddress: 'C2' },
            rule: { kind: 'decimal', operator: 'between', values: [0, 1] },
          }),
        ])

        const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, { maxRows: 32, maxColumns: 8, useColumnIndex: true })
        const orderSheetId = workbook.getSheetId('Order')!

        workbook.setCellContents({ sheet: orderSheetId, row: 1, col: 1 }, 'West')
        workbook.setCellContents({ sheet: orderSheetId, row: 1, col: 2 }, 0.5)
        expect(workbook.getCellValue({ sheet: orderSheetId, row: 1, col: 1 })).toEqual({
          tag: ValueTag.String,
          value: 'West',
          stringId: expect.any(Number),
        })
        expect(workbook.getCellValue({ sheet: orderSheetId, row: 1, col: 2 })).toEqual({ tag: ValueTag.Number, value: 0.5 })

        expect(() => workbook.setCellContents({ sheet: orderSheetId, row: 1, col: 1 }, 'Bogus')).toThrow(/Excel data validation/)
        expect(() => workbook.setCellContents({ sheet: orderSheetId, row: 1, col: 2 }, 1.5)).toThrow(/Excel data validation/)
        expect(workbook.getCellValue({ sheet: orderSheetId, row: 1, col: 1 })).toEqual({
          tag: ValueTag.String,
          value: 'West',
          stringId: expect.any(Number),
        })
        expect(workbook.getCellValue({ sheet: orderSheetId, row: 1, col: 2 })).toEqual({ tag: ValueTag.Number, value: 0.5 })

        const headlessWorkbookPath = join(tempDir, 'headless-data-validation-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(workbook.exportSnapshot()))
        const headlessExcel = runMacosExcelStructuralOperationOracle({
          workbookPath: headlessWorkbookPath,
          worksheetName: 'Order',
          operations: [{ kind: 'setCellValue', address: 'B2', value: 'North' }],
          inspectCells: ['B2', 'C2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(headlessExcel.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'B2', value: { kind: 'string', value: 'North' } },
          { address: 'C2', value: { kind: 'number', value: 0.5 } },
        ])

        const roundTrip = importXlsx(new Uint8Array(readFileSync(headlessWorkbookPath)), 'headless-data-validation-oracle.xlsx')
        expect(roundTrip.snapshot.sheets.find((sheet) => sheet.name === 'Order')?.metadata?.validations).toEqual(
          imported.snapshot.sheets.find((sheet) => sheet.name === 'Order')?.metadata?.validations,
        )
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    120_000,
  )
})

function validationWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Data validation oracle',
      metadata: {
        definedNames: [
          {
            name: 'Regions',
            value: {
              kind: 'range-ref',
              sheetName: 'Choices',
              startAddress: 'A2',
              endAddress: 'A4',
            },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Order',
        order: 0,
        metadata: {
          validations: [
            {
              range: { sheetName: 'Order', startAddress: 'B2', endAddress: 'B2' },
              rule: {
                kind: 'list',
                source: {
                  kind: 'named-range',
                  name: 'Regions',
                },
              },
              allowBlank: false,
              showDropdown: true,
              errorStyle: 'stop',
              errorTitle: 'Region required',
              errorMessage: 'Pick a known sales region.',
            },
            {
              range: { sheetName: 'Order', startAddress: 'C2', endAddress: 'C2' },
              rule: { kind: 'decimal', operator: 'between', values: [0, 1] },
              allowBlank: false,
              errorStyle: 'stop',
              errorTitle: 'Discount out of range',
              errorMessage: 'Discount must be between 0 and 1.',
            },
          ],
        },
        cells: [
          { address: 'A1', value: 'Order' },
          { address: 'B1', value: 'Region' },
          { address: 'C1', value: 'Discount' },
          { address: 'B2', value: 'East' },
          { address: 'C2', value: 0.25 },
        ],
      },
      {
        id: 2,
        name: 'Choices',
        order: 1,
        cells: [
          { address: 'A1', value: 'Regions' },
          { address: 'A2', value: 'East' },
          { address: 'A3', value: 'West' },
          { address: 'A4', value: 'North' },
        ],
      },
    ],
  }
}
