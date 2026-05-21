import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { dataTableFormulasWarning, exportXlsx, importXlsx } from '@bilig/excel-import'
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
import { ErrorCode, ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }
const indexImplicitIntersectionConfig = { maxRows: 10, maxColumns: 10, useColumnIndex: true }
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
const singleImplicitIntersectionOracleAddresses = ['C1', 'C2', 'C3', 'D1'] as const
const expectedSingleImplicitIntersectionOracleValues = [
  { address: 'C1', value: { kind: 'number', value: 1 } },
  { address: 'C2', value: { kind: 'number', value: 2 } },
  { address: 'C3', value: { kind: 'number', value: 3 } },
  { address: 'D1', value: { kind: 'number', value: 1 } },
] as const
const expectedDesktopExcelSingleImplicitIntersectionOracleCells = [
  { address: 'C1', formula: '=A1:A3', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  { address: 'C2', formula: '=A1:A3', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'C3', formula: '=A1:A3', rawValue: 'number\t3.0', value: { kind: 'number', value: 3 } },
  { address: 'D1', formula: '=SUM(@A1:A3)', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
] as const
const dynamicSpillOracleAddresses = ['B1', 'B2', 'B3'] as const
const expectedDynamicSpillOracleValues = [
  { address: 'B1', value: { kind: 'number', value: 2 } },
  { address: 'B2', value: { kind: 'number', value: 4 } },
  { address: 'B3', value: { kind: 'number', value: 6 } },
] as const
const spillReferenceOracleAddresses = ['B1', 'B2', 'B3', 'D1', 'E1', 'F1'] as const
const expectedSpillReferenceOracleCells = [
  { address: 'B1', formula: '=SEQUENCE(3,1,1,1)', rawValue: 'number\t1.0', value: { kind: 'number', value: 1 } },
  { address: 'B2', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'B3', rawValue: 'number\t3.0', value: { kind: 'number', value: 3 } },
  { address: 'D1', formula: '=SUM(B1#)', rawValue: 'number\t6.0', value: { kind: 'number', value: 6 } },
  { address: 'E1', formula: '=ROWS(B1#)', rawValue: 'number\t3.0', value: { kind: 'number', value: 3 } },
  { address: 'F1', formula: '=INDEX(B1#,2)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
] as const
const textsplitErrorOracleAddresses = ['C1', 'D1', 'C2', 'D2'] as const
const expectedTextsplitErrorOracleCells = [
  { address: 'C1', formula: '=TEXTSPLIT(A1,",","|")', rawValue: 'string\tred', value: { kind: 'string', value: 'red' } },
  { address: 'D1', rawValue: 'string\tblue', value: { kind: 'string', value: 'blue' } },
  { address: 'C2', rawValue: 'string\tgreen', value: { kind: 'string', value: 'green' } },
  { address: 'D2', rawValue: 'error\t#N/A', value: { kind: 'error', value: String(ErrorCode.NA) } },
] as const
const chooseArrayIndexOracleAddresses = ['E1', 'F1', 'E2', 'F2', 'E3', 'F3', 'H1', 'H2', 'H3', 'H4', 'H6', 'H7'] as const
const expectedChooseArrayIndexOracleValues = [
  { address: 'E1', value: { kind: 'string', value: 'a' } },
  { address: 'F1', value: { kind: 'number', value: 10 } },
  { address: 'E2', value: { kind: 'string', value: 'b' } },
  { address: 'F2', value: { kind: 'number', value: 20 } },
  { address: 'E3', value: { kind: 'string', value: 'c' } },
  { address: 'F3', value: { kind: 'number', value: 30 } },
  { address: 'H1', value: { kind: 'number', value: 660 } },
  { address: 'H2', value: { kind: 'number', value: 100 } },
  { address: 'H3', value: { kind: 'number', value: 200 } },
  { address: 'H4', value: { kind: 'number', value: 300 } },
  { address: 'H6', value: { kind: 'number', value: 600 } },
  { address: 'H7', value: { kind: 'number', value: 20 } },
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
const indexImplicitIntersectionOracleAddresses = ['E1', 'E2', 'E3', 'E4', 'A5', 'B5', 'C5', 'D5', 'E5', 'G1', 'H1', 'I1'] as const
const expectedIndexImplicitIntersectionOracleCells = [
  { address: 'E1', formula: '=INDEX(A1:C3,0,2)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'E2', formula: '=INDEX(A1:C3,0,2)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'E3', formula: '=INDEX(A1:C3,0,2)', rawValue: 'number\t8.0', value: { kind: 'number', value: 8 } },
  { address: 'E4', formula: '=INDEX(A1:C3,0,2)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'A5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'number\t4.0', value: { kind: 'number', value: 4 } },
  { address: 'B5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'C5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'number\t6.0', value: { kind: 'number', value: 6 } },
  { address: 'D5', formula: '=INDEX(A1:C3,2,0)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'E5', formula: '=INDEX(A1:C3,0,0)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'G1', formula: '=SUM(INDEX(A1:C3,0,2))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'H1', formula: '=SUM(INDEX(A1:C3,2,0))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'I1', formula: '=SUM(INDEX(A1:C3,0,0))', rawValue: 'number\t45.0', value: { kind: 'number', value: 45 } },
] as const
const offsetImplicitIntersectionOracleAddresses = ['E1', 'E2', 'E3', 'E4', 'A5', 'B5', 'C5', 'D5', 'E5', 'G1', 'H1', 'I1'] as const
const expectedOffsetImplicitIntersectionOracleCells = [
  { address: 'E1', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'number\t2.0', value: { kind: 'number', value: 2 } },
  { address: 'E2', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'E3', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'number\t8.0', value: { kind: 'number', value: 8 } },
  { address: 'E4', formula: '=OFFSET(A1,0,1,3,1)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'A5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'number\t4.0', value: { kind: 'number', value: 4 } },
  { address: 'B5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'number\t5.0', value: { kind: 'number', value: 5 } },
  { address: 'C5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'number\t6.0', value: { kind: 'number', value: 6 } },
  { address: 'D5', formula: '=OFFSET(A2,0,0,1,3)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'E5', formula: '=OFFSET(A1,0,0,3,3)', rawValue: 'error\t#VALUE!', value: { kind: 'error', value: String(ErrorCode.Value) } },
  { address: 'G1', formula: '=SUM(OFFSET(A1,0,1,3,1))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'H1', formula: '=SUM(OFFSET(A2,0,0,1,3))', rawValue: 'number\t15.0', value: { kind: 'number', value: 15 } },
  { address: 'I1', formula: '=SUM(OFFSET(A1,0,0,3,3))', rawValue: 'number\t45.0', value: { kind: 'number', value: 45 } },
] as const
const dataTableOracleAddresses = ['C3', 'D3', 'C4', 'D4'] as const
const expectedDataTableOracleValues = [
  { address: 'C3', value: { kind: 'number', value: 40 } },
  { address: 'D3', value: { kind: 'number', value: 60 } },
  { address: 'C4', value: { kind: 'number', value: 60 } },
  { address: 'D4', value: { kind: 'number', value: 90 } },
] as const
const expectedDataTableImportedFormulaByAddress = new Map([
  ['C3', '=MULTIPLE.OPERATIONS(B2,A1,C2,A2,B3)'],
  ['D3', '=MULTIPLE.OPERATIONS(B2,A1,D2,A2,B3)'],
  ['C4', '=MULTIPLE.OPERATIONS(B2,A1,C2,A2,B4)'],
  ['D4', '=MULTIPLE.OPERATIONS(B2,A1,D2,A2,B4)'],
] as const)
const oneVariableDataTableOracleAddresses = ['C2', 'D2', 'B6', 'B7', 'B8'] as const
const expectedOneVariableDataTableOracleValues = [
  { address: 'C2', value: { kind: 'number', value: 30 } },
  { address: 'D2', value: { kind: 'number', value: 40 } },
  { address: 'B6', value: { kind: 'number', value: 20 } },
  { address: 'B7', value: { kind: 'number', value: 30 } },
  { address: 'B8', value: { kind: 'number', value: 40 } },
] as const
const expectedOneVariableDataTableImportedFormulaByAddress = new Map([
  ['C2', '=MULTIPLE.OPERATIONS(B2,A1,C1)'],
  ['D2', '=MULTIPLE.OPERATIONS(B2,A1,D1)'],
  ['B6', '=MULTIPLE.OPERATIONS(B5,A1,A6)'],
  ['B7', '=MULTIPLE.OPERATIONS(B5,A1,A7)'],
  ['B8', '=MULTIPLE.OPERATIONS(B5,A1,A8)'],
] as const)
const aggregateOptionsOracleAddresses = ['B1', 'B2', 'B3', 'B4', 'C1'] as const
const expectedAggregateOptionsOracleCells = [
  { address: 'B1', formula: '=AGGREGATE(9,3,A1:A5)', rawValue: 'number\t40.0', value: { kind: 'number', value: 40 } },
  { address: 'B2', formula: '=AGGREGATE(9,6,A1:A5)', rawValue: 'number\t120.0', value: { kind: 'number', value: 120 } },
  { address: 'B3', formula: '=AGGREGATE(9,4,A1:A5)', rawValue: 'error\t#DIV/0!', value: { kind: 'error', value: String(ErrorCode.Div0) } },
  { address: 'B4', formula: '=AGGREGATE(9,7,A1:A5)', rawValue: 'number\t100.0', value: { kind: 'number', value: 100 } },
  { address: 'C1', formula: '=SUBTOTAL(109,A1:A4)', rawValue: 'number\t40.0', value: { kind: 'number', value: 40 } },
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

  it('exports and reimports AGGREGATE option semantics through the headless workbook path', () => {
    const workbook = buildAggregateOptionsOracleWorkbook()
    try {
      expect(aggregateOptionsOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedAggregateOptionsOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-aggregate-options-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          aggregateOptionsOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedAggregateOptionsOracleCells.map((expected) => expected.value))
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports and reimports standalone INDEX implicit-intersection formulas without spill metadata', () => {
    const workbook = buildIndexImplicitIntersectionOracleWorkbook()
    try {
      expect(
        indexImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedIndexImplicitIntersectionOracleCells.map((expected) => expected.value))
      expect(workbook.engine.getSpillRanges()).toEqual([])

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-index-implicit-intersection-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, indexImplicitIntersectionConfig)
      try {
        expect(
          indexImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedIndexImplicitIntersectionOracleCells.map((expected) => expected.value))
        expect(reimported.engine.getSpillRanges()).toEqual([])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports and reimports standalone OFFSET implicit-intersection formulas without spill metadata', () => {
    const workbook = buildOffsetImplicitIntersectionOracleWorkbook()
    try {
      expect(
        offsetImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedOffsetImplicitIntersectionOracleCells.map((expected) => expected.value))
      expect(workbook.engine.getSpillRanges()).toEqual([])

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-offset-implicit-intersection-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, indexImplicitIntersectionConfig)
      try {
        expect(
          offsetImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedOffsetImplicitIntersectionOracleCells.map((expected) => expected.value))
        expect(reimported.engine.getSpillRanges()).toEqual([])
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

  it('exports SINGLE implicit-intersection formulas through Desktop Excel-compatible XLSX', () => {
    const workbook = buildSingleImplicitIntersectionOracleWorkbook()
    try {
      expect(
        singleImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedSingleImplicitIntersectionOracleValues.map((expected) => expected.value))

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-single-implicit-intersection-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          singleImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedSingleImplicitIntersectionOracleValues.map((expected) => expected.value))
        expect(reimported.getCellFormula(addressToCell('C1'))).toBe('=SINGLE(A1:A3)')
        expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(SINGLE(A1:A3))')
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

  it('exports spill-reference consumer formulas through Desktop Excel-compatible XLSX', () => {
    const workbook = buildSpillReferenceOracleWorkbook()
    try {
      expect(spillReferenceOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedSpillReferenceOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-spill-reference-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          spillReferenceOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedSpillReferenceOracleCells.map((expected) => expected.value))
        expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(B1#)')
        expect(reimported.getCellFormula(addressToCell('E1'))).toBe('=ROWS(B1#)')
        expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=INDEX(B1#,2)')
        expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'B1', rows: 3, cols: 1 }])
      } finally {
        reimported.dispose()
      }
    } finally {
      workbook.dispose()
    }
  })

  it('exports and reimports TEXTSPLIT error spill children through the headless XLSX path', () => {
    const workbook = buildTextsplitErrorOracleWorkbook()
    try {
      expect(textsplitErrorOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedTextsplitErrorOracleCells.map((expected) => expected.value),
      )

      const imported = importXlsx(exportXlsx(workbook.exportSnapshot()), 'headless-textsplit-error-oracle.xlsx')
      const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
      try {
        expect(
          textsplitErrorOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
        ).toEqual(expectedTextsplitErrorOracleCells.map((expected) => expected.value))
        expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'C1', rows: 2, cols: 2 }])
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

  it('imports native two-variable data-table outputs into headless calculable formulas', () => {
    const imported = importXlsx(buildNativeDataTableXlsx(), 'headless-native-data-table-oracle.xlsx')
    expect(imported.warnings).not.toContain(dataTableFormulasWarning)
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas).toEqual([
      {
        address: 'C3',
        formulaXml: '<f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2"/>',
      },
    ])

    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(dataTableOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address))))).toEqual(
        expectedDataTableOracleValues.map((expected) => expected.value),
      )
      for (const [address, formula] of expectedDataTableImportedFormulaByAddress) {
        expect(workbook.getCellFormula(addressToCell(address))).toBe(formula)
      }
    } finally {
      workbook.dispose()
    }
  })

  it('imports native one-variable data-table outputs into headless calculable formulas', () => {
    const imported = importXlsx(buildNativeOneVariableDataTableXlsx(), 'headless-native-one-variable-data-table-oracle.xlsx')
    expect(imported.warnings).not.toContain(dataTableFormulasWarning)
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas).toEqual([
      {
        address: 'C2',
        formulaXml: '<f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1"/>',
      },
      {
        address: 'B6',
        formulaXml: '<f t="dataTable" ref="B6:B8" dt2D="0" dtr="0" r1="A1"/>',
      },
    ])

    const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
    try {
      expect(
        oneVariableDataTableOracleAddresses.map((address) => normalizedCellValue(workbook.getCellValue(addressToCell(address)))),
      ).toEqual(expectedOneVariableDataTableOracleValues.map((expected) => expected.value))
      for (const [address, formula] of expectedOneVariableDataTableImportedFormulaByAddress) {
        expect(workbook.getCellFormula(addressToCell(address))).toBe(formula)
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
    'round-trips Desktop Excel AGGREGATE option semantics into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-aggregate-options-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-aggregate-options-oracle.xlsx')
        const workbook = buildAggregateOptionsOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: aggregateOptionsOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedAggregateOptionsOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-aggregate-options-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          const comparisons = buildHeadlessExcelComparisons(reimported, excelResult.cells, 'headless-aggregate-options-oracle')
          const summary = buildReportSummary({
            workbooks: [
              {
                id: 'headless-aggregate-options-oracle',
                workbook: 'headless-aggregate-options-oracle.xlsx',
                elapsedMs: 0,
                formulaCells: comparisons.length,
                status: 'ok',
                comparisons,
              },
            ],
          })

          expect(comparisons.map((comparison) => comparison.classification)).toEqual(
            aggregateOptionsOracleAddresses.map(() => 'bilig_matches_excel'),
          )
          expect(summary).toMatchObject({
            biligVsFreshExcelMatchRate: 1,
            comparableFormulaCells: aggregateOptionsOracleAddresses.length,
            realBiligMismatches: 0,
            totalFormulaCells: aggregateOptionsOracleAddresses.length,
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
    'matches Desktop Excel standalone INDEX implicit-intersection semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-index-implicit-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-index-implicit-intersection-oracle.xlsx')
        const workbook = buildIndexImplicitIntersectionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Sheet1',
          formulaCells: [],
          inspectCells: indexImplicitIntersectionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedIndexImplicitIntersectionOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-index-implicit-intersection-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, indexImplicitIntersectionConfig)
        try {
          expect(
            indexImplicitIntersectionOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedIndexImplicitIntersectionOracleCells.map((expected) => expected.value))
          expect(reimported.engine.getSpillRanges()).toEqual([])
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
    'matches Desktop Excel standalone OFFSET implicit-intersection semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-offset-implicit-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-offset-implicit-intersection-oracle.xlsx')
        const workbook = buildOffsetImplicitIntersectionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Sheet1',
          formulaCells: [],
          inspectCells: offsetImplicitIntersectionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedOffsetImplicitIntersectionOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-offset-implicit-intersection-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, indexImplicitIntersectionConfig)
        try {
          expect(
            offsetImplicitIntersectionOracleAddresses.map((address) =>
              normalizedCellValue(reimported.getCellValue(addressToCell(address))),
            ),
          ).toEqual(expectedOffsetImplicitIntersectionOracleCells.map((expected) => expected.value))
          expect(reimported.engine.getSpillRanges()).toEqual([])
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
    'matches Desktop Excel SINGLE implicit-intersection semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-single-implicit-intersection-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-single-implicit-intersection-oracle.xlsx')
        const workbook = buildSingleImplicitIntersectionOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: singleImplicitIntersectionOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedDesktopExcelSingleImplicitIntersectionOracleCells)

        const imported = importXlsx(
          new Uint8Array(readFileSync(workbookPath)),
          'headless-single-implicit-intersection-oracle-recalculated.xlsx',
        )
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            singleImplicitIntersectionOracleAddresses.map((address) =>
              normalizedCellValue(reimported.getCellValue(addressToCell(address))),
            ),
          ).toEqual(expectedSingleImplicitIntersectionOracleValues.map((expected) => expected.value))
          expect(reimported.getCellFormula(addressToCell('C1'))).toBe('=SINGLE(A1:A3)')
          expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(SINGLE(A1:A3))')
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
    'matches Desktop Excel spill-reference formula consumer semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-spill-reference-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-spill-reference-oracle.xlsx')
        const workbook = buildSpillReferenceOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: spillReferenceOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedSpillReferenceOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-spill-reference-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            spillReferenceOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedSpillReferenceOracleCells.map((expected) => expected.value))
          expect(reimported.getCellFormula(addressToCell('D1'))).toBe('=SUM(B1#)')
          expect(reimported.getCellFormula(addressToCell('E1'))).toBe('=ROWS(B1#)')
          expect(reimported.getCellFormula(addressToCell('F1'))).toBe('=INDEX(B1#,2)')
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
    'round-trips Desktop Excel TEXTSPLIT error spill-child caches back into headless import',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-textsplit-error-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-textsplit-error-oracle.xlsx')
        const workbook = buildTextsplitErrorOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [],
          inspectCells: textsplitErrorOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells).toEqual(expectedTextsplitErrorOracleCells)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-textsplit-error-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            textsplitErrorOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedTextsplitErrorOracleCells.map((expected) => expected.value))
          expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Cases', address: 'C1', rows: 2, cols: 2 }])
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
    'matches Desktop Excel CHOOSE array-index virtual table semantics',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-choose-array-index-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-choose-array-index-oracle.xlsx')
        const workbook = buildChooseArrayIndexOracleWorkbook()
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'ChooseRef',
          formulaCells: [],
          inspectCells: chooseArrayIndexOracleAddresses,
          saveWorkbook: true,
        })

        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedChooseArrayIndexOracleValues)
        expect(excelResult.cells[0]?.formula).toBe('=CHOOSE({1,2},A1:A3,B1:B3)')

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-choose-array-index-oracle-recalculated.xlsx')
        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            chooseArrayIndexOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedChooseArrayIndexOracleValues.map((expected) => expected.value))
          expect(imported.snapshot.workbook.metadata?.spills).toEqual([
            { sheetName: 'ChooseRef', address: 'E1', rows: 3, cols: 2 },
            { sheetName: 'ChooseRef', address: 'H2', rows: 3, cols: 1 },
          ])
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

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'imports Desktop Excel native data-table outputs into headless formulas',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-data-table-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-data-table-oracle.xlsx')
        const workbook = buildDataTableOracleWorkbook(false)
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'DataTable',
          operations: [{ kind: 'createDataTable', range: 'B2:D4', rowInput: 'A1', columnInput: 'A2' }],
          inspectCells: dataTableOracleAddresses,
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedDataTableOracleValues)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-data-table-oracle-recalculated.xlsx')
        expect(imported.warnings).not.toContain(dataTableFormulasWarning)
        expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas[0]?.address).toBe('C3')
        expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas[0]?.formulaXml).toContain('t="dataTable"')

        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(dataTableOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address))))).toEqual(
            expectedDataTableOracleValues.map((expected) => expected.value),
          )
          for (const [address, formula] of expectedDataTableImportedFormulaByAddress) {
            expect(reimported.getCellFormula(addressToCell(address))).toBe(formula)
          }
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
    'imports Desktop Excel native one-variable data-table outputs into headless formulas',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-one-variable-data-table-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-one-variable-data-table-oracle.xlsx')
        const workbook = buildOneVariableDataTableOracleWorkbook(false)
        try {
          writeFileSync(workbookPath, exportXlsx(workbook.exportSnapshot()))
        } finally {
          workbook.dispose()
        }

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'DataTable',
          operations: [
            { kind: 'createDataTable', range: 'B1:D2', rowInput: 'A1' },
            { kind: 'createDataTable', range: 'A5:B8', columnInput: 'A1' },
          ],
          inspectCells: oneVariableDataTableOracleAddresses,
          saveWorkbook: true,
        })
        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual(expectedOneVariableDataTableOracleValues)

        const imported = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-one-variable-data-table-oracle-recalculated.xlsx')
        expect(imported.warnings).not.toContain(dataTableFormulasWarning)
        expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas.map(({ address }) => address)).toEqual(['C2', 'B6'])

        const reimported = WorkPaper.buildFromSnapshot(imported.snapshot, workbookConfig)
        try {
          expect(
            oneVariableDataTableOracleAddresses.map((address) => normalizedCellValue(reimported.getCellValue(addressToCell(address)))),
          ).toEqual(expectedOneVariableDataTableOracleValues.map((expected) => expected.value))
          for (const [address, formula] of expectedOneVariableDataTableImportedFormulaByAddress) {
            expect(reimported.getCellFormula(addressToCell(address))).toBe(formula)
          }
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

function buildAggregateOptionsOracleWorkbook(): WorkPaper {
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: { name: 'headless-aggregate-options-oracle' },
    sheets: [
      {
        id: 1,
        name: 'Cases',
        order: 0,
        metadata: {
          rows: [{ id: 'row:1', index: 1, hidden: true }],
        },
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
          { address: 'A4', formula: 'SUBTOTAL(9,A1:A3)' },
          { address: 'A5', formula: '1/0' },
          { address: 'B1', formula: 'AGGREGATE(9,3,A1:A5)' },
          { address: 'B2', formula: 'AGGREGATE(9,6,A1:A5)' },
          { address: 'B3', formula: 'AGGREGATE(9,4,A1:A5)' },
          { address: 'B4', formula: 'AGGREGATE(9,7,A1:A5)' },
          { address: 'C1', formula: 'SUBTOTAL(109,A1:A4)' },
        ],
      },
    ],
  }
  return WorkPaper.buildFromSnapshot(snapshot, workbookConfig)
}

function buildIndexImplicitIntersectionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Sheet1: [
        [1, 2, 3, null, '=INDEX(A1:C3,0,2)', null, '=SUM(INDEX(A1:C3,0,2))', '=SUM(INDEX(A1:C3,2,0))', '=SUM(INDEX(A1:C3,0,0))'],
        [4, 5, 6, null, '=INDEX(A1:C3,0,2)'],
        [7, 8, 9, null, '=INDEX(A1:C3,0,2)'],
        [null, null, null, null, '=INDEX(A1:C3,0,2)'],
        ['=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,2,0)', '=INDEX(A1:C3,0,0)'],
      ],
    },
    indexImplicitIntersectionConfig,
  )
}

function buildOffsetImplicitIntersectionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Sheet1: [
        [1, 2, 3, null, '=OFFSET(A1,0,1,3,1)', null, '=SUM(OFFSET(A1,0,1,3,1))', '=SUM(OFFSET(A2,0,0,1,3))', '=SUM(OFFSET(A1,0,0,3,3))'],
        [4, 5, 6, null, '=OFFSET(A1,0,1,3,1)'],
        [7, 8, 9, null, '=OFFSET(A1,0,1,3,1)'],
        [null, null, null, null, '=OFFSET(A1,0,1,3,1)'],
        ['=OFFSET(A2,0,0,1,3)', '=OFFSET(A2,0,0,1,3)', '=OFFSET(A2,0,0,1,3)', '=OFFSET(A2,0,0,1,3)', '=OFFSET(A1,0,0,3,3)'],
      ],
    },
    indexImplicitIntersectionConfig,
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

function buildSingleImplicitIntersectionOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [
        [1, null, '=SINGLE(A1:A3)', '=SUM(SINGLE(A1:A3))'],
        [2, null, '=SINGLE(A1:A3)'],
        [3, null, '=SINGLE(A1:A3)'],
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

function buildSpillReferenceOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [[1, '=SEQUENCE(3,1,1,1)', null, '=SUM(B1#)', '=ROWS(B1#)', '=INDEX(B1#,2)'], [2], [3]],
    },
    workbookConfig,
  )
}

function buildTextsplitErrorOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Cases: [['red,blue|green', null, '=TEXTSPLIT(A1,",","|")']],
    },
    workbookConfig,
  )
}

function buildChooseArrayIndexOracleWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      ChooseRef: [
        ['a', 10, 100, null, '=CHOOSE({1,2},A1:A3,B1:B3)', null, null, '=SUM(CHOOSE({1,2},B1:B3,C1:C3))'],
        ['b', 20, 200, null, null, null, null, '=CHOOSE(2,A1:A3,C1:C3)'],
        ['c', 30, 300],
        [],
        [],
        [null, null, null, null, null, null, null, '=SUM(CHOOSE(2,B1:B3,C1:C3))'],
        [null, null, null, null, null, null, null, '=XLOOKUP("b",CHOOSE(1,A1:A3,C1:C3),CHOOSE(1,B1:B3,C1:C3),"missing",0)'],
      ],
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

function buildDataTableOracleWorkbook(includeOutputs: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      DataTable: [
        [1],
        [10, '=A3', 2, 3],
        ['=A1*A2', 20, includeOutputs ? 40 : null, includeOutputs ? 60 : null],
        [null, 30, includeOutputs ? 60 : null, includeOutputs ? 90 : null],
      ],
    },
    workbookConfig,
  )
}

function buildOneVariableDataTableOracleWorkbook(includeOutputs: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      DataTable: [
        [1, 2, 3, 4],
        ['=A1*10', '=A2', includeOutputs ? 30 : null, includeOutputs ? 40 : null],
        [],
        [],
        [1, '=A1*10'],
        [2, includeOutputs ? 20 : null],
        [3, includeOutputs ? 30 : null],
        [4, includeOutputs ? 40 : null],
      ],
    },
    workbookConfig,
  )
}

function buildNativeDataTableXlsx(): Uint8Array {
  const workbook = buildDataTableOracleWorkbook(true)
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      sheetXml.replace(
        /<c\b[^>]*\br=(["'])C3\1[^>]*>[\s\S]*?<\/c>/u,
        '<c r="C3"><f t="dataTable" ref="C3:D4" dt2D="1" dtr="1" r1="A1" r2="A2"/><v>40</v></c>',
      ),
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function buildNativeOneVariableDataTableXlsx(): Uint8Array {
  const workbook = buildOneVariableDataTableOracleWorkbook(true)
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      sheetXml
        .replace(
          /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
          '<c r="C2"><f t="dataTable" ref="C2:D2" dt2D="0" dtr="1" r1="A1"/><v>30</v></c>',
        )
        .replace(
          /<c\b[^>]*\br=(["'])B6\1[^>]*>[\s\S]*?<\/c>/u,
          '<c r="B6"><f t="dataTable" ref="B6:B8" dt2D="0" dtr="0" r1="A1"/><v>20</v></c>',
        ),
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
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
