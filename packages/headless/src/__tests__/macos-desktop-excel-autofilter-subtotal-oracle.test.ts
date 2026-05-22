import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const inspectedCells = ['D1', 'E1', 'F1', 'G1'] as const

describe('macOS Desktop Excel AutoFilter SUBTOTAL oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel filtered-row semantics after applying AutoFilter',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-autofilter-subtotal-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-autofilter-subtotal-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(ledgerSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Ledger',
          operations: [{ kind: 'applyAutoFilter', range: 'A1:B6', field: 1, criteria1: 'East' }],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'D1', value: { kind: 'number', value: 90 } },
          { address: 'E1', value: { kind: 'number', value: 90 } },
          { address: 'F1', value: { kind: 'number', value: 150 } },
          { address: 'G1', value: { kind: 'number', value: 90 } },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-autofilter-subtotal-oracle.xlsx')
        expect(excelTruth.snapshot.sheets[0]?.metadata?.rows).toEqual(
          expect.arrayContaining([
            expect.objectContaining({ index: 2, hidden: true, filtered: true }),
            expect.objectContaining({ index: 4, hidden: true, filtered: true }),
          ]),
        )

        const engine = new SpreadsheetEngine({ workbookName: 'headless-autofilter-subtotal-oracle' })
        await engine.ready()
        engine.importSnapshot(excelTruth.snapshot)
        engine.recalculateNow()

        expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 90 })
        expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 90 })
        expect(engine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
        expect(engine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 90 })

        const localEngine = new SpreadsheetEngine({ workbookName: 'headless-local-autofilter-subtotal-oracle' })
        await localEngine.ready()
        localEngine.importSnapshot(ledgerSnapshot())
        localEngine.setFilter('Ledger', {
          sheetName: 'Ledger',
          startAddress: 'A1',
          endAddress: 'B6',
          criteria: [
            {
              colId: 0,
              filters: { values: ['East'] },
            },
          ],
        })

        expectRowFiltered(localEngine.getRowMetadata('Ledger'), 2)
        expectRowFiltered(localEngine.getRowMetadata('Ledger'), 4)
        expect(localEngine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 90 })
        expect(localEngine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 90 })
        expect(localEngine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
        expect(localEngine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 90 })

        localEngine.clearFilter('Ledger', { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B6' })
        expect(localEngine.getRowMetadata('Ledger').some((record) => record.filtered === true)).toBe(false)
        expect(localEngine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 150 })
        expect(localEngine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 150 })
        expect(localEngine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
        expect(localEngine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 150 })

        localEngine.setFilter('Ledger', {
          sheetName: 'Ledger',
          startAddress: 'A1',
          endAddress: 'B6',
          criteria: [
            {
              colId: 0,
              filters: { values: ['East'] },
            },
          ],
        })
        const headlessWorkbookPath = join(tempDir, 'headless-autofilter-subtotal-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(localEngine.exportSnapshot()))
        const headlessExcel = runMacosExcelStructuralOperationOracle({
          workbookPath: headlessWorkbookPath,
          worksheetName: 'Ledger',
          operations: [{ kind: 'setCellValue', address: 'A1', value: 'Region' }],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(headlessExcel.cells.map(({ address, value }) => ({ address, value }))).toEqual(
          excelResult.cells.map(({ address, value }) => ({ address, value })),
        )
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function ledgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'AutoFilter subtotal oracle' },
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

function expectRowFiltered(
  records: ReadonlyArray<{ readonly start: number; readonly count: number; readonly filtered?: boolean | null }>,
  row: number,
): void {
  expect(records.some((record) => record.filtered === true && row >= record.start && row < record.start + record.count)).toBe(true)
}
