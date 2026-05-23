import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, type NormalizedFormulaValue } from '@bilig/excel-fixtures'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const inspectedCells = ['A1', 'B1', 'C1', 'D1'] as const
const expectedValues: readonly NormalizedFormulaValue[] = [
  { kind: 'number', value: 10 },
  { kind: 'number', value: 20 },
  { kind: 'number', value: 30 },
  { kind: 'string', value: '' },
] as const

describe('macOS Desktop Excel conditional format artifact oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Desktop Excel advanced visual conditional-format rules after a headless edit',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-cf-artifacts-oracle-')
      try {
        const sourcePath = join(tempDir, 'excel-advanced-conditional-format-oracle.xlsx')
        writeFileSync(sourcePath, buildAdvancedConditionalFormattingWorkbook())

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Dashboard',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual(expectedValues)

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-advanced-conditional-format-oracle.xlsx')
        expectConditionalFormatArtifacts(excelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)

        const workpaper = WorkPaper.buildFromSnapshot(excelTruth.snapshot)
        const sheet = workpaper.getSheetId('Dashboard')
        if (sheet === undefined) {
          throw new Error('Expected Dashboard sheet to be available')
        }
        workpaper.setCellContents({ sheet, row: 0, col: 3 }, 'headless edit')
        const headlessSnapshot = workpaper.exportSnapshot()
        expectConditionalFormatArtifacts(headlessSnapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)

        const headlessPath = join(tempDir, 'headless-advanced-conditional-format-oracle.xlsx')
        writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: headlessPath,
          worksheetName: 'Dashboard',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(headlessExcel.cells.map((cell) => cell.value)).toEqual([
          { kind: 'number', value: 10 },
          { kind: 'number', value: 20 },
          { kind: 'number', value: 30 },
          { kind: 'string', value: 'headless edit' },
        ])

        const headlessExcelTruth = importXlsx(
          new Uint8Array(readFileSync(headlessPath)),
          'headless-advanced-conditional-format-oracle.xlsx',
        )
        expectConditionalFormatArtifacts(headlessExcelTruth.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml)
        expect(readWorksheetXml(headlessPath).match(/<conditionalFormatting\b/gu)).toHaveLength(3)
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function expectConditionalFormatArtifacts(xml: string | undefined): void {
  expect(xml).toContain('type="dataBar"')
  expect(xml).toContain('type="colorScale"')
  expect(xml).toContain('type="iconSet"')
}

function buildAdvancedConditionalFormattingWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    [10, 20, 30, null],
    [20, 40, 60, null],
    [30, 60, 90, null],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Dashboard')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(advancedConditionalFormattingWorksheetXml)
  return zipSync(zip)
}

function readWorksheetXml(path: string): string {
  return strFromU8(unzipSync(new Uint8Array(readFileSync(path)))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
}

const advancedConditionalFormattingWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<dimension ref="A1:D3"/>',
  '<sheetData>',
  '<row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>20</v></c><c r="C1"><v>30</v></c></row>',
  '<row r="2"><c r="A2"><v>20</v></c><c r="B2"><v>40</v></c><c r="C2"><v>60</v></c></row>',
  '<row r="3"><c r="A3"><v>30</v></c><c r="B3"><v>60</v></c><c r="C3"><v>90</v></c></row>',
  '</sheetData>',
  '<conditionalFormatting sqref="A1:A3">',
  '<cfRule type="dataBar" priority="1">',
  '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF63C384"/></dataBar>',
  '</cfRule>',
  '</conditionalFormatting>',
  '<conditionalFormatting sqref="B1:B3">',
  '<cfRule type="colorScale" priority="2">',
  '<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>',
  '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>',
  '</cfRule>',
  '</conditionalFormatting>',
  '<conditionalFormatting sqref="C1:C3">',
  '<cfRule type="iconSet" priority="3">',
  '<iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/>',
  '<cfvo type="percent" val="67"/></iconSet>',
  '</cfRule>',
  '</conditionalFormatting>',
  '</worksheet>',
].join('')
