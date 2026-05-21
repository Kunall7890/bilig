import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import {
  buildFormulaCellComparison,
  buildReportSummary,
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  type FormulaCellComparison,
  type NormalizedFormulaValue,
} from '@bilig/excel-fixtures'
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

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: ['C1', 'C2'],
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual([
          { address: 'C1', formula: '=A1+B1*2', rawValue: 'number\t16.0', value: { kind: 'number', value: 16 } },
          { address: 'C2', formula: '=SUM(A1:B2)', rawValue: 'number\t22.0', value: { kind: 'number', value: 22 } },
        ])

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          const comparisons = buildHeadlessExcelComparisons(reimported, excelResult.cells)
          const summary = buildReportSummary({
            workbooks: [
              {
                id: 'headless-oracle',
                workbook: 'headless-oracle.xlsx',
                elapsedMs: 0,
                formulaCells: comparisons.length,
                status: 'ok',
                comparisons,
              },
            ],
          })

          expect(comparisons.map((comparison) => comparison.classification)).toEqual(['bilig_matches_excel', 'bilig_matches_excel'])
          expect(summary).toMatchObject({
            biligVsFreshExcelMatchRate: 1,
            comparableFormulaCells: 2,
            realBiligMismatches: 0,
            totalFormulaCells: 2,
          })
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

function buildHeadlessExcelComparisons(
  workbook: WorkPaper,
  excelCells: readonly { readonly address: string; readonly formula?: string; readonly value: NormalizedFormulaValue }[],
): FormulaCellComparison[] {
  return excelCells.map((excelCell) => {
    const address = addressToCell(excelCell.address)
    const formula = workbook.getCellFormula(address)
    if (!formula) {
      throw new Error(`Missing imported formula at ${excelCell.address}`)
    }
    return buildFormulaCellComparison({
      workbookId: 'headless-oracle',
      sheet: 'Cases',
      address: excelCell.address,
      formula,
      ...(excelCell.formula !== undefined ? { excelOracleFormula: excelCell.formula } : {}),
      excelOracleValue: excelCell.value,
      actualBiligValue: normalizedCellValue(workbook.getCellValue(address)),
    })
  })
}

function addressToCell(address: string): WorkPaperCellAddress {
  switch (address) {
    case 'C1':
      return cell(0, 2)
    case 'C2':
      return cell(1, 2)
    default:
      throw new Error(`Unexpected oracle address: ${address}`)
  }
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
