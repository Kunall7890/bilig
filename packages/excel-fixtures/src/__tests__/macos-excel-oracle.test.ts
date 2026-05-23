import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { ErrorCode } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import {
  createMacosExcelRecalculationAppleScript,
  createMacosExcelInspectionAppleScript,
  createMacosExcelStructuralOperationAppleScript,
  isMacosExcelInstalled,
  parseMacosExcelInspectionOutput,
  parseMacosExcelRecalculationOutput,
  runMacosExcelInspectionOracle,
  runMacosExcelRecalculationOracle,
  runMacosExcelStructuralOperationOracle,
} from '../macos-excel-oracle.js'

describe('macOS Desktop Excel oracle harness', () => {
  it('builds an AppleScript runner that opens, recalculates, reads, and closes a workbook', () => {
    const script = createMacosExcelRecalculationAppleScript({
      worksheetName: 'Cases',
      formulaCells: [{ address: 'C1', formula: '=A1+B1' }],
      valueCells: ['C1'],
    })

    expect(script).toContain('tell application "Microsoft Excel"')
    expect(script).toContain('open workbook workbook file name workbookPath update links do not update links')
    expect(script).toContain('set formula of range "C1"')
    expect(script).toContain('calculate full rebuild')
    expect(script).toContain('my typedCellValue(value of range "C1"')
    expect(script).toContain('string value of range "C1"')
    expect(script).toContain('close targetWorkbook saving no')
    expect(script).not.toContain('set display alerts')
    expect(script).not.toContain('active workbook')
  })

  it('can save the opened workbook when a caller needs fresh Excel caches persisted', () => {
    const script = createMacosExcelRecalculationAppleScript({
      worksheetName: 'Cases',
      formulaCells: [],
      valueCells: ['C1'],
      saveWorkbook: true,
    })

    expect(script).toContain('close targetWorkbook saving yes')
  })

  it('can open companion workbooks and ask Excel to update external links', () => {
    const script = createMacosExcelInspectionAppleScript({
      worksheetName: 'Cases',
      formulaCells: [],
      inspectCells: ['C1'],
      updateLinks: 'external',
    })

    expect(script).toContain('repeat with companionIndex from 2 to count of argv')
    expect(script).toContain('open workbook workbook file name companionPath update links do not update links')
    expect(script).toContain('open workbook workbook file name workbookPath update links update external links only')
    expect(script).toContain('repeat with companionWorkbook in companionWorkbooks')
  })

  it('builds an inspection runner that reads formulas and values from the opened workbook', () => {
    const script = createMacosExcelInspectionAppleScript({
      worksheetName: 'Cases',
      formulaCells: [{ address: 'C1', formula: '=A1+B1' }],
      inspectCells: ['C1'],
      saveWorkbook: true,
    })

    expect(script).toContain('set inspectedRange to range "C1"')
    expect(script).toContain('open workbook workbook file name workbookPath update links do not update links')
    expect(script).toContain('my formulaText(formula of inspectedRange)')
    expect(script).toContain('my typedCellValue(value of inspectedRange, string value of inspectedRange)')
    expect(script).toContain('close targetWorkbook saving yes')
    expect(script).not.toContain('set display alerts')
    expect(script).not.toContain('active workbook')
  })

  it('builds a structural operation runner that can cut-insert columns before inspection', () => {
    const script = createMacosExcelStructuralOperationAppleScript({
      worksheetName: 'Cases',
      operations: [
        { kind: 'setCellValue', address: 'B1', value: 9 },
        { kind: 'clearCell', address: 'B1' },
        { kind: 'moveColumns', sourceRange: 'B:B', destinationRange: 'F:F' },
      ],
      inspectCells: ['F1'],
      saveWorkbook: true,
    })

    expect(script).toContain('set targetWorksheet to worksheet "Cases"')
    expect(script).toContain('open workbook workbook file name workbookPath update links do not update links')
    expect(script).toContain('set value of range "B1" of targetWorksheet to 9')
    expect(script).toContain('clear contents range "B1" of targetWorksheet')
    expect(script).toContain('cut range (range "B:B" of targetWorksheet)')
    expect(script).toContain('insert into range (range "F:F" of targetWorksheet) shift shift to right')
    expect(script).toContain('set inspectedRange to range "F1" of targetWorksheet')
    expect(script).toContain('my typedCellValue(value of inspectedRange, string value of inspectedRange)')
    expect(script).toContain('close targetWorkbook saving yes')
    expect(script).not.toContain('set display alerts')
    expect(script).not.toContain('active workbook')
  })

  it('builds a structural operation runner that can apply Desktop Excel range sorts', () => {
    const script = createMacosExcelStructuralOperationAppleScript({
      worksheetName: 'Cases',
      operations: [
        {
          kind: 'applySort',
          range: 'A1:D7',
          keys: [
            { key: 'B1', order: 'descending' },
            { key: 'A1', order: 'ascending' },
          ],
          header: 'yes',
          orientation: 'rows',
        },
      ],
      inspectCells: ['A2'],
      saveWorkbook: true,
    })

    expect(script).toContain('sort (range "A1:D7" of targetWorksheet)')
    expect(script).toContain('key1 (range "B1" of targetWorksheet)')
    expect(script).toContain('order1 sort descending')
    expect(script).toContain('key2 (range "A1" of targetWorksheet)')
    expect(script).toContain('order2 sort ascending')
    expect(script).toContain('header header yes')
    expect(script).toContain('orientation sort columns')
  })

  it('builds a structural operation runner that can apply Desktop Excel table sorts', () => {
    const script = createMacosExcelStructuralOperationAppleScript({
      worksheetName: 'Cases',
      operations: [
        {
          kind: 'applyTableSort',
          tableName: 'Sales',
          keys: [{ key: 'B2:B6', order: 'descending' }],
          header: 'yes',
          orientation: 'rows',
        },
      ],
      inspectCells: ['A2'],
      saveWorkbook: true,
    })

    expect(script).toContain('set tableSort to sort object of list object "Sales" of targetWorksheet')
    expect(script).toContain('clear sortfieldset (sortfieldset of tableSort)')
    expect(script).toContain('add sortfield (sortfieldset of tableSort) key (range "B2:B6" of targetWorksheet) order sort descending')
    expect(script).toContain('set sort header of tableSort to header yes')
    expect(script).toContain('set sort orientation of tableSort to sort columns')
    expect(script).toContain('apply sort tableSort')
  })

  it('builds one-variable and two-variable data-table structural operations', () => {
    const script = createMacosExcelStructuralOperationAppleScript({
      worksheetName: 'Cases',
      operations: [
        { kind: 'createDataTable', range: 'B1:D2', rowInput: 'A1' },
        { kind: 'createDataTable', range: 'A5:B8', columnInput: 'A1' },
        { kind: 'createDataTable', range: 'B2:D4', rowInput: 'A1', columnInput: 'A2' },
      ],
      inspectCells: ['C2'],
    })

    expect(script).toContain('data table (range "B1:D2" of targetWorksheet) row input (range "A1" of targetWorksheet)')
    expect(script).toContain('data table (range "A5:B8" of targetWorksheet) column input (range "A1" of targetWorksheet)')
    expect(script).toContain(
      'data table (range "B2:D4" of targetWorksheet) row input (range "A1" of targetWorksheet) column input (range "A2" of targetWorksheet)',
    )
  })

  it('parses typed Excel oracle values into normalized formula values', () => {
    expect(
      parseMacosExcelRecalculationOutput(
        ['version=16.96', 'number\t42', 'boolean\ttrue', 'string\tBilig', 'blank\t', 'error\t#N/A'].join('\n'),
        5,
      ),
    ).toEqual({
      excelVersion: '16.96',
      rawValues: ['number\t42', 'boolean\ttrue', 'string\tBilig', 'blank\t', 'error\t#N/A'],
      values: [
        { kind: 'number', value: 42 },
        { kind: 'boolean', value: true },
        { kind: 'string', value: 'Bilig' },
        { kind: 'blank' },
        { kind: 'error', value: String(ErrorCode.NA) },
      ],
    })
  })

  it('parses inspected formula cells into address-keyed normalized values', () => {
    expect(
      parseMacosExcelInspectionOutput(
        ['version=16.109', 'C1\tA1+B1*2\tnumber\t16.0', 'D1\t\tstring\tready', 'E1\tNA()\terror\t#N/A'].join('\n'),
        ['C1', 'D1', 'E1'],
      ),
    ).toEqual({
      excelVersion: '16.109',
      cells: [
        {
          address: 'C1',
          formula: 'A1+B1*2',
          rawValue: 'number\t16.0',
          value: { kind: 'number', value: 16 },
        },
        {
          address: 'D1',
          rawValue: 'string\tready',
          value: { kind: 'string', value: 'ready' },
        },
        {
          address: 'E1',
          formula: 'NA()',
          rawValue: 'error\t#N/A',
          value: { kind: 'error', value: String(ErrorCode.NA) },
        },
      ],
    })
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'uses real Microsoft Excel for Mac as the recalculation oracle',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-excel-fixtures-live-'))
      try {
        const workbookPath = join(tempDir, 'oracle.xlsx')
        const worksheet = XLSX.utils.aoa_to_sheet([[10, 3, null]])
        const workbook = XLSX.utils.book_new()
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Cases')
        writeFileSync(workbookPath, XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }))

        const result = runMacosExcelRecalculationOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [{ address: 'C1', formula: '=A1+B1*2' }],
          valueCells: ['C1'],
        })

        expect(result.excelVersion).toMatch(/^\d+\./u)
        expect(result.values).toEqual([{ kind: 'number', value: 16 }])

        const inspection = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Cases',
          formulaCells: [{ address: 'C1', formula: '=A1+B1*2' }],
          inspectCells: ['C1'],
        })
        expect(inspection.cells).toEqual([
          { address: 'C1', formula: '=A1+B1*2', rawValue: 'number\t16.0', value: { kind: 'number', value: 16 } },
        ])

        const structuralWorkbookPath = join(tempDir, 'structural-oracle.xlsx')
        const structuralWorksheet = XLSX.utils.aoa_to_sheet([[1, 2, 3, 4, 5, { f: 'SUM(B1:C1)' }]])
        const structuralWorkbook = XLSX.utils.book_new()
        XLSX.utils.book_append_sheet(structuralWorkbook, structuralWorksheet, 'Cases')
        writeFileSync(structuralWorkbookPath, XLSX.write(structuralWorkbook, { type: 'buffer', bookType: 'xlsx' }))

        const structural = runMacosExcelStructuralOperationOracle({
          workbookPath: structuralWorkbookPath,
          worksheetName: 'Cases',
          operations: [{ kind: 'moveColumns', sourceRange: 'B:B', destinationRange: 'F:F' }],
          inspectCells: ['F1'],
        })
        expect(structural.cells).toEqual([
          { address: 'F1', formula: '=SUM(B1:B1)', rawValue: 'number\t3.0', value: { kind: 'number', value: 3 } },
        ])
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )
})
