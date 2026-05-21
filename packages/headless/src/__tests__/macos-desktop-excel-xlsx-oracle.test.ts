import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelRecalculationOracle, type NormalizedFormulaValue } from '@bilig/excel-fixtures'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }

describe('macOS Desktop Excel XLSX oracle for WorkPaper', () => {
  it('exports and reimports the oracle fixture through the headless workbook path', () => {
    const workbook = buildOracleWorkbook()
    try {
      expect(normalizedCellValue(workbook.getCellValue(cell(0, 2)))).toEqual({ kind: 'number', value: 16 })
      expect(normalizedCellValue(workbook.getCellValue(cell(1, 2)))).toEqual({ kind: 'number', value: 22 })

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(normalizedCellValue(reimported.getCellValue(cell(0, 2)))).toEqual({ kind: 'number', value: 16 })
        expect(normalizedCellValue(reimported.getCellValue(cell(1, 2)))).toEqual({ kind: 'number', value: 22 })
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips fresh Desktop Excel recalculation caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-oracle.xlsx')
        const workbook = buildOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelRecalculationOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          valueCells: ['C1', 'C2'],
          saveWorkbook: true,
        })

        expect(excelResult.values).toEqual([
          { kind: 'number', value: 16 },
          { kind: 'number', value: 22 },
        ])

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(normalizedCellValue(reimported.getCellValue(cell(0, 2)))).toEqual(excelResult.values[0])
          expect(normalizedCellValue(reimported.getCellValue(cell(1, 2)))).toEqual(excelResult.values[1])
        } finally {
          reimported.dispose()
        }
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )
})

function buildOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [
        [10, 3, '=A1+B1*2'],
        [5, 4, '=SUM(A1:B2)'],
      ],
    },
    workbookConfig,
  )
}

function cell(row: number, col: number): WorkPaperCellAddress {
  return { sheet: 1, row, col }
}

function normalizedCellValue(value: CellValue): NormalizedFormulaValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return { kind: 'blank' }
    case ValueTag.Boolean:
      return { kind: 'boolean', value: value.value }
    case ValueTag.Error:
      return { kind: 'error', value: String(value.code) }
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.String:
      return { kind: 'string', value: value.value }
  }
}
