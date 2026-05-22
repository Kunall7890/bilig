import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

const inspectedCells = ['D1', 'E1', 'F1', 'G1'] as const

describe('macOS Desktop Excel table AutoFilter oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves table-scoped criteria and visible-row formula semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-autofilter-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-table-autofilter-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(tableLedgerSnapshot()))

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
        expect(onlyTableXml(new Uint8Array(readFileSync(excelWorkbookPath)))).toContain(
          '<filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn>',
        )

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-table-autofilter-oracle.xlsx')
        expect(excelTruth.snapshot.workbook.metadata?.tables?.[0]?.autoFilter).toMatchObject({
          sheetName: 'Ledger',
          startAddress: 'A1',
          endAddress: 'B6',
          criteria: [{ colId: 0, filters: { values: ['East'] } }],
        })
        expectRowFiltered(excelTruth.snapshot.sheets[0]?.metadata?.rows, 2)
        expectRowFiltered(excelTruth.snapshot.sheets[0]?.metadata?.rows, 4)

        const engine = new SpreadsheetEngine({ workbookName: 'headless-table-autofilter-oracle' })
        await engine.ready()
        engine.importSnapshot(excelTruth.snapshot)
        engine.recalculateNow()

        expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 90 })
        expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 90 })
        expect(engine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
        expect(engine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 90 })

        const headlessWorkbookPath = join(tempDir, 'headless-table-autofilter-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(engine.exportSnapshot()))
        expect(onlyTableXml(new Uint8Array(readFileSync(headlessWorkbookPath)))).toContain(
          '<filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn>',
        )

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
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    120_000,
  )
})

function tableLedgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Table AutoFilter oracle',
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

function onlyTableXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  const tablePaths = Object.keys(zip).filter((path) => /^xl\/tables\/table[0-9]+\.xml$/u.test(path))
  expect(tablePaths).toHaveLength(1)
  return strFromU8(zip[tablePaths[0]] ?? new Uint8Array())
}

function expectRowFiltered(
  rows: ReadonlyArray<{ readonly index: number; readonly filtered?: boolean | null }> | undefined,
  row: number,
): void {
  expect(rows).toEqual(expect.arrayContaining([expect.objectContaining({ index: row, filtered: true })]))
}
