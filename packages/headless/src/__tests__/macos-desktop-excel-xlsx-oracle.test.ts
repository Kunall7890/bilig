import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { readRuntimeImage, SpreadsheetEngine } from '@bilig/core'
import {
  buildFormulaCellComparison,
  buildReportSummary,
  isMacosExcelInstalled,
  runMacosExcelInspectionOracle,
  runMacosExcelStructuralOperationOracle,
  type FormulaCellComparison,
  type NormalizedFormulaValue,
} from '@bilig/excel-fixtures'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
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
const structuralMoveColumnFormulaOracleCell = {
  address: 'F1',
  formula: '=SUM(B1:B1)',
  rawValue: 'number\t3.0',
  value: { kind: 'number', value: 3 },
} as const
const tableColumnInsertOracleCells = [
  { address: 'B1', formula: 'Column1', rawValue: 'string\tColumn1', value: { kind: 'string', value: 'Column1' } },
  { address: 'F1', formula: '=SUM(Sales[Margin])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
const tableColumnDeleteOracleFormulaCells = [
  { address: 'D1', formula: '=SUM(#REF!)' },
  { address: 'E1', formula: '=SUM(Sales[Margin])' },
] as const
const tableColumnDeleteDefinedNameOracleFormulaCells = [
  { address: 'D1', formula: '=SUM(SalesAmount)' },
  { address: 'E1', formula: '=SUM(SalesAmountFormula)' },
] as const
const tableHeaderRenameOracleCells = [
  { address: 'B1', formula: 'Revenue', rawValue: 'string\tRevenue', value: { kind: 'string', value: 'Revenue' } },
  { address: 'E1', formula: '=SUM(Sales[Revenue])', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'F1', formula: '=SUM(Sales[Margin])', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
] as const
const tableHeaderRenameDefinedNameOracleCells = [
  { address: 'B1', formula: 'Revenue', rawValue: 'string\tRevenue', value: { kind: 'string', value: 'Revenue' } },
  { address: 'E1', formula: '=SUM(SalesAmount)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
  { address: 'F1', formula: '=SUM(SalesAmountFormula)', rawValue: 'number\t30.0', value: { kind: 'number', value: 30 } },
] as const
const tableEmptyBodyOracleCell = {
  address: 'D1',
  formula: '=SUM(Sales[Amount])',
  rawValue: 'number\t0.0',
  value: { kind: 'number', value: 0 },
} as const

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

  it('rewrites moved-out column ranges like Desktop Excel', () => {
    const workbook = buildStructuralMoveColumnOracleWorkbook()
    const sheetId = workbook.getSheetId('Cases')!
    try {
      workbook.moveColumns(sheetId, 1, 1, 4)
      expect(workbook.getCellFormula(addressToCell('F1'))).toBe('=SUM(B1:B1)')
      expect(normalizedCellValue(workbook.getCellValue(addressToCell('F1')))).toEqual(structuralMoveColumnFormulaOracleCell.value)

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-structural-move-column-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=SUM(B1:B1)')
        expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual(structuralMoveColumnFormulaOracleCell.value)
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('keeps dynamic-array spill metadata valid after structural row inserts', () => {
    const workbook = buildDynamicSpillOracleWorkbook()
    const sheetId = workbook.getSheetId('Cases')!
    try {
      workbook.addRows(sheetId, 0, 1)
      expect(workbook.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B2', rows: 3, cols: 1 }])
      expect(['B2', 'B3', 'B4'].map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedDynamicSpillOracleValues.map((expected) => expected.value),
      )
      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-structural-dynamic-spill-oracle.xlsx')
      expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B2', rows: 3, cols: 1 }])
    } finally {
      workbook.dispose()
    }
  })

  it('generates Excel-compatible table headers when inserting columns inside tables', async () => {
    const engine = await buildTableColumnInsertOracleEngine()

    engine.insertColumns('Data', 1, 1)

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'D3',
      columnNames: ['Region', 'Column1', 'Revenue', 'Margin'],
    })
    expect(engine.getCellValue('Data', 'B1')).toMatchObject({ tag: ValueTag.String, value: 'Column1' })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(Sales[Margin])')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-column-insert-oracle.xlsx')
    expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'D3',
      columnNames: ['Region', 'Column1', 'Revenue', 'Margin'],
    })
  })

  it('rewrites deleted table-column structured references before XLSX export', async () => {
    const engine = await buildTableColumnDeleteOracleEngine()

    engine.deleteColumns('Data', 1, 1)

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Margin'],
    })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(#REF!)')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Margin])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-column-delete-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B3',
        columnNames: ['Region', 'Margin'],
      })
      expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(#REF!)')
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual({
        kind: 'error',
        value: String(ErrorCode.Ref),
      })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 5 })
    } finally {
      reimported.dispose()
    }
  })

  it('renames table headers and structured references before XLSX export', async () => {
    const engine = await buildTableHeaderRenameOracleEngine()

    engine.setCellValue('Data', 'B1', 'Revenue')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Revenue', 'Margin'],
    })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Revenue])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(Sales[Margin])')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-header-rename-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'C3',
        columnNames: ['Region', 'Revenue', 'Margin'],
      })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual({ kind: 'number', value: 5 })
    } finally {
      reimported.dispose()
    }
  })

  it('rewrites table-column defined names to #REF! before XLSX export', async () => {
    const engine = await buildTableColumnDeleteDefinedNameOracleEngine()

    engine.deleteColumns('Data', 1, 1)

    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'formula', formula: '=#REF!' },
    })
    expect(engine.getDefinedName('SalesAmountFormula')).toEqual({
      name: 'SalesAmountFormula',
      value: { kind: 'formula', formula: '=#REF!' },
    })
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-column-delete-defined-name-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
        { name: 'SalesAmount', value: { kind: 'formula', formula: '=#REF!' } },
        { name: 'SalesAmountFormula', value: { kind: 'formula', formula: '=#REF!' } },
      ])
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual({ kind: 'error', value: String(ErrorCode.Ref) })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'error', value: String(ErrorCode.Ref) })
    } finally {
      reimported.dispose()
    }
  })

  it('renames table-column defined names before XLSX export', async () => {
    const engine = await buildTableHeaderRenameDefinedNameOracleEngine()

    engine.setCellValue('Data', 'B1', 'Revenue')

    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'structured-ref', tableName: 'Sales', columnName: 'Revenue' },
    })
    expect(engine.getDefinedName('SalesAmountFormula')).toEqual({
      name: 'SalesAmountFormula',
      value: { kind: 'formula', formula: '=Sales[Revenue]' },
    })
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 30 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-header-rename-defined-name-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
        { name: 'SalesAmount', value: { kind: 'formula', formula: '=Sales[Revenue]' } },
        { name: 'SalesAmountFormula', value: { kind: 'formula', formula: '=Sales[Revenue]' } },
      ])
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual({ kind: 'number', value: 30 })
    } finally {
      reimported.dispose()
    }
  })

  it('routes WorkPaper table header edits through structured reference rewrites', async () => {
    const engine = await buildTableHeaderRenameOracleEngine()
    const workbook = WorkPaper.buildFromSnapshot(engine.exportSnapshot(), workbookConfig)
    try {
      workbook.setCellContents(addressToCell('B1'), 'Revenue')

      expect(workbook.getCellFormula(addressToCell('E1'))).toBe('=SUM(Sales[Revenue])')
      expect(normalizedCellValue(workbook.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
      expect(workbook.exportSnapshot().workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'C3',
        columnNames: ['Region', 'Revenue', 'Margin'],
      })
    } finally {
      workbook.dispose()
    }
  })

  it('keeps table structured-reference aggregates valid when deleting the only data row', async () => {
    const engine = await buildTableEmptyBodyOracleEngine()

    engine.deleteRows('Data', 1, 1)

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Region', 'Amount'],
    })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(Sales[Amount])')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Number, value: 0 })

    const imported = importXlsx(exportXlsx(engine.exportSnapshot()), 'headless-table-empty-body-oracle.xlsx')
    const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B2',
        columnNames: ['Region', 'Amount'],
      })
      expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual(tableEmptyBodyOracleCell.value)
    } finally {
      reimported.dispose()
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

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel column-move formula rewrite semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-structural-move-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-structural-move-column-oracle.xlsx')
        const workbook = buildStructuralMoveColumnOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Cases',
          operations: [{ kind: 'moveColumns', sourceRange: 'B:B', destinationRange: 'F:F' }],
          inspectCells: ['F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual([structuralMoveColumnFormulaOracleCell])

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-structural-move-column-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=SUM(B1:B1)')
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual(structuralMoveColumnFormulaOracleCell.value)
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
    'matches Desktop Excel table column-insert structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-column-insert-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-column-insert-oracle.xlsx')
        const engine = await buildTableColumnInsertOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertColumns', range: 'B:B' }],
          inspectCells: ['B1', 'F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(tableColumnInsertOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-column-insert-oracle-recalculated.xlsx')
        expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
          name: 'Sales',
          sheetName: 'Data',
          startAddress: 'A1',
          endAddress: 'D3',
          columnNames: ['Region', 'Column1', 'Revenue', 'Margin'],
        })
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel table header-rename structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-header-rename-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-header-rename-oracle.xlsx')
        const engine = await buildTableHeaderRenameOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'setCellValue', address: 'B1', value: 'Revenue' }],
          inspectCells: ['B1', 'E1', 'F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(tableHeaderRenameOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-header-rename-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
            name: 'Sales',
            sheetName: 'Data',
            startAddress: 'A1',
            endAddress: 'C3',
            columnNames: ['Region', 'Revenue', 'Margin'],
          })
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual({ kind: 'number', value: 5 })
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
    'matches Desktop Excel table-header defined-name rename semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-header-defined-name-rename-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-header-defined-name-rename-oracle.xlsx')
        const engine = await buildTableHeaderRenameDefinedNameOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'setCellValue', address: 'B1', value: 'Revenue' }],
          inspectCells: ['B1', 'E1', 'F1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual(tableHeaderRenameDefinedNameOracleCells)

        const imported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-table-header-defined-name-rename-oracle-recalculated.xlsx',
        )
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
            { name: 'SalesAmount', value: { kind: 'formula', formula: '=Sales[Revenue]' } },
            { name: 'SalesAmountFormula', value: { kind: 'formula', formula: '=Sales[Revenue]' } },
          ])
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 30 })
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('F1')))).toEqual({ kind: 'number', value: 30 })
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
    'matches Desktop Excel table column-delete structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-column-delete-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-column-delete-oracle.xlsx')
        const engine = await buildTableColumnDeleteOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteColumns', range: 'B:B' }],
          inspectCells: ['D1', 'E1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, formula }) => ({ address, formula }))).toEqual(tableColumnDeleteOracleFormulaCells)
        expect(excelResult.cells[1]).toMatchObject({
          address: 'E1',
          rawValue: 'number\t5.0',
          value: { kind: 'number', value: 5 },
        })

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-column-delete-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
            name: 'Sales',
            sheetName: 'Data',
            startAddress: 'A1',
            endAddress: 'B3',
            columnNames: ['Region', 'Margin'],
          })
          expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(#REF!)')
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual({
            kind: 'error',
            value: String(ErrorCode.Ref),
          })
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({ kind: 'number', value: 5 })
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
    'matches Desktop Excel table-column defined-name deletion semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-column-defined-name-delete-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-column-defined-name-delete-oracle.xlsx')
        const engine = await buildTableColumnDeleteDefinedNameOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteColumns', range: 'B:B' }],
          inspectCells: ['D1', 'E1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, formula }) => ({ address, formula }))).toEqual(
          tableColumnDeleteDefinedNameOracleFormulaCells,
        )

        const imported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-table-column-defined-name-delete-oracle-recalculated.xlsx',
        )
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
            { name: 'SalesAmount', value: { kind: 'formula', formula: '=#REF!' } },
            { name: 'SalesAmountFormula', value: { kind: 'formula', formula: '=#REF!' } },
          ])
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual({
            kind: 'error',
            value: String(ErrorCode.Ref),
          })
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('E1')))).toEqual({
            kind: 'error',
            value: String(ErrorCode.Ref),
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
    'matches Desktop Excel empty-table-body structured-reference semantics',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-table-empty-body-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-table-empty-body-oracle.xlsx')
        const engine = await buildTableEmptyBodyOracleEngine()
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteRows', range: '2:2' }],
          inspectCells: ['D1'],
          saveWorkbook: true,
        })
        expect(excelResult.cells).toEqual([tableEmptyBodyOracleCell])

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-table-empty-body-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(imported.snapshot.workbook.metadata?.tables?.[0]).toMatchObject({
            name: 'Sales',
            sheetName: 'Data',
            startAddress: 'A1',
            endAddress: 'B2',
            columnNames: ['Region', 'Amount'],
          })
          expect(normalizedCellValue(reimported.getCellValue(addressToCell('D1')))).toEqual(tableEmptyBodyOracleCell.value)
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

function buildStructuralMoveColumnOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[1, 2, 3, 4, 5, '=SUM(B1:C1)']],
    },
    workbookConfig,
  )
}

async function buildTableColumnInsertOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-column-insert-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Revenue', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Revenue', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Margin])')
  return engine
}

async function buildTableColumnDeleteOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-column-delete-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Amount', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Amount', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Amount])')
  engine.setCellFormula('Data', 'F1', 'SUM(Sales[Margin])')
  return engine
}

async function buildTableColumnDeleteDefinedNameOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-column-delete-defined-name-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Amount', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Amount', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('SalesAmount', { kind: 'structured-ref', tableName: 'Sales', columnName: 'Amount' })
  engine.setDefinedName('SalesAmountFormula', { kind: 'formula', formula: '=Sales[Amount]' })
  engine.setCellFormula('Data', 'E1', 'SUM(SalesAmount)')
  engine.setCellFormula('Data', 'F1', 'SUM(SalesAmountFormula)')
  return engine
}

async function buildTableHeaderRenameDefinedNameOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-header-rename-defined-name-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Amount', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Amount', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('SalesAmount', { kind: 'structured-ref', tableName: 'Sales', columnName: 'Amount' })
  engine.setDefinedName('SalesAmountFormula', { kind: 'formula', formula: '=Sales[Amount]' })
  engine.setCellFormula('Data', 'E1', 'SUM(SalesAmount)')
  engine.setCellFormula('Data', 'F1', 'SUM(SalesAmountFormula)')
  return engine
}

async function buildTableHeaderRenameOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-header-rename-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Amount', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Amount', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Amount])')
  engine.setCellFormula('Data', 'F1', 'SUM(Sales[Margin])')
  return engine
}

async function buildTableEmptyBodyOracleEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'table-empty-body-oracle' })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }, [
    ['Region', 'Amount'],
    ['East', 10],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'B2',
    columnNames: ['Region', 'Amount'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setCellFormula('Data', 'D1', 'SUM(Sales[Amount])')
  return engine
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
