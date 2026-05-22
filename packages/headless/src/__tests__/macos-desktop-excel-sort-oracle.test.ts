import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

const inspectedCells = ['A2', 'B2', 'C2', 'A3', 'B3', 'C3', 'A4', 'B4', 'C4', 'A5', 'B5', 'C5', 'A6', 'B6', 'C6', 'E1'] as const

const expectedSortedRows = [
  ['East', 50, 'invoice-005'],
  ['West', 40, 'invoice-002'],
  ['East', 30, 'invoice-003'],
  ['West', 20, 'invoice-004'],
  ['East', 10, 'invoice-001'],
] as const

describe('macOS Desktop Excel sort oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel structural range sort semantics before preserving sort metadata',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-sort-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-sort-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(ledgerSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Ledger',
          operations: [
            {
              kind: 'applySort',
              range: 'A1:C6',
              keys: [{ key: 'B1', order: 'descending' }],
              header: 'yes',
              orientation: 'rows',
            },
          ],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelRows(excelResult.cells)).toEqual(expectedSortedRows)
        expect(excelResult.cells.at(-1)?.value).toEqual({ kind: 'number', value: 50 })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-sort-oracle.xlsx')
        expect(importedRows(excelTruth.snapshot)).toEqual(expectedSortedRows)
        expect(excelTruth.snapshot.sheets[0]?.metadata?.sorts?.[0]?.keys.length).toBeGreaterThan(0)

        const headless = new SpreadsheetEngine({ workbookName: 'headless-sort-oracle' })
        await headless.ready()
        headless.importSnapshot(ledgerSnapshot())
        expect(
          headless.sortRange(
            'Ledger',
            { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'C6' },
            [{ keyAddress: 'B1', direction: 'desc' }],
            { header: true },
          ),
        ).toBe(true)

        expect(engineRows(headless)).toEqual(expectedSortedRows)
        expect(headless.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 50 })

        const headlessWorkbookPath = join(tempDir, 'headless-sort-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(headless.exportSnapshot()))
        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: headlessWorkbookPath,
          worksheetName: 'Ledger',
          formulaCells: [],
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

function ledgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Sort oracle' },
    sheets: [
      {
        id: 1,
        name: 'Ledger',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'C1', value: 'Invoice' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 10 },
          { address: 'C2', value: 'invoice-001' },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 40 },
          { address: 'C3', value: 'invoice-002' },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 30 },
          { address: 'C4', value: 'invoice-003' },
          { address: 'A5', value: 'West' },
          { address: 'B5', value: 20 },
          { address: 'C5', value: 'invoice-004' },
          { address: 'A6', value: 'East' },
          { address: 'B6', value: 50 },
          { address: 'C6', value: 'invoice-005' },
          { address: 'E1', formula: 'B2', value: 10 },
        ],
      },
    ],
  }
}

function excelRows(cells: readonly { readonly value: { readonly value?: unknown } }[]): unknown[][] {
  return [
    [cells[0]?.value.value, cells[1]?.value.value, cells[2]?.value.value],
    [cells[3]?.value.value, cells[4]?.value.value, cells[5]?.value.value],
    [cells[6]?.value.value, cells[7]?.value.value, cells[8]?.value.value],
    [cells[9]?.value.value, cells[10]?.value.value, cells[11]?.value.value],
    [cells[12]?.value.value, cells[13]?.value.value, cells[14]?.value.value],
  ]
}

function importedRows(snapshot: WorkbookSnapshot): unknown[][] {
  const sheet = snapshot.sheets[0]
  const values = new Map(sheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])
  return [
    [values.get('A2'), values.get('B2'), values.get('C2')],
    [values.get('A3'), values.get('B3'), values.get('C3')],
    [values.get('A4'), values.get('B4'), values.get('C4')],
    [values.get('A5'), values.get('B5'), values.get('C5')],
    [values.get('A6'), values.get('B6'), values.get('C6')],
  ]
}

function engineRows(engine: SpreadsheetEngine): unknown[][] {
  return engine.getRangeValues({ sheetName: 'Ledger', startAddress: 'A2', endAddress: 'C6' }).map((row) => row.map(cellValue))
}

function cellValue(value: CellValue): unknown {
  switch (value.tag) {
    case ValueTag.Number:
    case ValueTag.String:
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Empty:
    case ValueTag.Error:
      return null
    default:
      return null
  }
}
