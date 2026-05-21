import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { readRuntimeImage } from '@bilig/core'
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
const oracleFormulaAddresses = ['C1', 'D1', 'E1', 'F1', 'G1', 'H1'] as const
const expectedOracleCells = [
  { address: 'C1', formula: '=COUNTBLANK(A1:A5)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'D1', formula: '=COUNTIF(A1:A5,"")', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'E1', formula: '=COUNTIF(A1:A5,"<>")', rawValue: 'number\t4.0', value: { kind: 'number', value: 4 } },
  { address: 'F1', formula: '=SUMIF(A1:A5,"",B1:B5)', rawValue: 'number\t50.0', value: { kind: 'number', value: 50 } },
  { address: 'G1', formula: '=SUMIF(A1:A5,"<>",B1:B5)', rawValue: 'number\t130.0', value: { kind: 'number', value: 130 } },
  { address: 'H1', formula: '=SUMIFS(B1:B5,A1:A5,"<>")', rawValue: 'number\t130.0', value: { kind: 'number', value: 130 } },
] as const
const futureFunctionOracleAddresses = ['D2', 'E2', 'F2'] as const
const expectedFutureFunctionOracleCells = [
  { address: 'D2', formula: '=TEXTJOIN("-",TRUE,A2:A4)', rawValue: 'string\ta-c', value: { kind: 'string', value: 'a-c' } },
  { address: 'E2', formula: '=XLOOKUP("b",B2:B4,C2:C4)', rawValue: 'number\t20.0', value: { kind: 'number', value: 20 } },
  { address: 'F2', formula: '=XMATCH("b",B2:B4,0)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
] as const
const dynamicSpillOracleAddresses = ['B1', 'B2', 'B3'] as const
const expectedDynamicSpillOracleValues = [
  { address: 'B1', value: { kind: 'number', value: 2 } },
  { address: 'B2', value: { kind: 'number', value: 4 } },
  { address: 'B3', value: { kind: 'number', value: 6 } },
] as const

describe('macOS Desktop Excel XLSX oracle for WorkPaper', () => {
  it('exports and reimports the oracle fixture through the headless workbook path', () => {
    const workbook = buildOracleWorkbook()
    try {
      expect(oracleFormulaAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(oracleFormulaAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address))))).toEqual(
          expectedOracleCells.map((expected) => expected.value),
        )
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports Excel future functions in a Desktop Excel-compatible XLSX shape', () => {
    const workbook = buildFutureFunctionOracleWorkbook()
    try {
      expect(futureFunctionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedFutureFunctionOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-future-function-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          futureFunctionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedFutureFunctionOracleCells.map((expected) => expected.value))
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports native dynamic-array spill caches through the headless XLSX path', () => {
    const workbook = buildDynamicSpillOracleWorkbook()
    try {
      expect(dynamicSpillOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedDynamicSpillOracleValues.map((expected) => expected.value),
      )

      const snapshot = workbook.exportSnapshot()
      expect(
        readRuntimeImage(snapshot)
          ?.cellValues?.filter((cellValue) => cellValue.sheetName === 'Cases' && cellValue.col === 1)
          .map(({ row, value }) => ({ row, value })),
      ).toEqual([
        { row: 0, value: { tag: ValueTag.Number, value: 2 } },
        { row: 1, value: { tag: ValueTag.Number, value: 4 } },
        { row: 2, value: { tag: ValueTag.Number, value: 6 } },
      ])

      const imported = importXlsx(exportXlsx(snapshot), 'headless-dynamic-spill-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(dynamicSpillOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address))))).toEqual(
          expectedDynamicSpillOracleValues.map((expected) => expected.value),
        )
        expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
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
          inspectCells: oracleFormulaAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedOracleCells)

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

          expect(comparisons.map((comparison) => comparison.classification)).toEqual(
            oracleFormulaAddresses.map(() => 'bilig_matches_excel'),
          )
          expect(summary).toMatchObject({
            biligVsFreshExcelMatchRate: 1,
            comparableFormulaCells: oracleFormulaAddresses.length,
            realBiligMismatches: 0,
            totalFormulaCells: oracleFormulaAddresses.length,
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

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips fresh Desktop Excel future-function recalculation caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-future-function-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-future-function-oracle.xlsx')
        const workbook = buildFutureFunctionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: futureFunctionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedFutureFunctionOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-future-function-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          const comparisons = buildHeadlessExcelComparisons(reimported, excelResult.cells, 'headless-future-function-oracle')
          const summary = buildReportSummary({
            workbooks: [
              {
                id: 'headless-future-function-oracle',
                workbook: 'headless-future-function-oracle.xlsx',
                elapsedMs: 0,
                formulaCells: comparisons.length,
                status: 'ok',
                comparisons,
              },
            ],
          })

          expect(comparisons.map((comparison) => comparison.classification)).toEqual(
            futureFunctionOracleAddresses.map(() => 'bilig_matches_excel'),
          )
          expect(summary).toMatchObject({
            biligVsFreshExcelMatchRate: 1,
            comparableFormulaCells: futureFunctionOracleAddresses.length,
            realBiligMismatches: 0,
            totalFormulaCells: futureFunctionOracleAddresses.length,
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

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips Desktop Excel native dynamic-array spill caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-dynamic-spill-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-dynamic-spill-oracle.xlsx')
        const workbook = buildDynamicSpillOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: dynamicSpillOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedDynamicSpillOracleValues)
        expect(excelResult.cells[0]?.formula).toBe('=MAP(A1:A3,LAMBDA(x,x*2))')

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-dynamic-spill-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            dynamicSpillOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedDynamicSpillOracleValues.map((expected) => expected.value))
          expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
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
        [
          'North',
          10,
          '=COUNTBLANK(A1:A5)',
          '=COUNTIF(A1:A5,"")',
          '=COUNTIF(A1:A5,"<>")',
          '=SUMIF(A1:A5,"",B1:B5)',
          '=SUMIF(A1:A5,"<>",B1:B5)',
          '=SUMIFS(B1:B5,A1:A5,"<>")',
        ],
        [null, 20],
        ['=IF(TRUE,"","x")', 30],
        [' ', 40],
        ['South', 50],
      ],
    },
    workbookConfig,
  )
}

function buildFutureFunctionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [
        ['Label', 'Key', 'Value', 'Joined', 'Lookup', 'Match'],
        ['a', 'a', 10, '=TEXTJOIN("-",TRUE,A2:A4)', '=XLOOKUP("b",B2:B4,C2:C4)', '=XMATCH("b",B2:B4,0)'],
        [null, 'b', 20],
        ['c', 'c', 30],
      ],
    },
    workbookConfig,
  )
}

function buildDynamicSpillOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[1, '=MAP(A1:A3,LAMBDA(x,x*2))'], [2], [3]],
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
  workbookId = 'headless-oracle',
): FormulaCellComparison[] {
  return excelCells.map((excelCell) => {
    const address = addressToCell(excelCell.address)
    const formula = workbook.getCellFormula(address)
    if (!formula) {
      throw new Error(`Missing imported formula at ${excelCell.address}`)
    }
    return buildFormulaCellComparison({
      workbookId,
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
