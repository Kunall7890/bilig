import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const workbookConfig = { maxRows: 16, maxColumns: 8, useColumnIndex: true }
const expectedShiftedHyperlinks = [
  {
    sheetName: 'Links',
    address: 'A2',
    target: 'https://example.com/report',
    tooltip: 'Open report',
    display: 'Open report',
  },
  {
    sheetName: 'Links',
    address: 'B3',
    target: '#Summary!A1',
    tooltip: 'Jump to summary',
    display: 'Summary',
  },
] as const

describe('macOS Desktop Excel hyperlink structural oracle', () => {
  it('retargets imported hyperlink metadata through headless structural inserts and XLSX export', () => {
    const workbook = WorkPaper.buildFromSnapshot(
      importXlsx(buildHyperlinkWorkbookBytes(), 'hyperlink-structural.xlsx').snapshot,
      workbookConfig,
    )
    try {
      const sheetId = workbook.getSheetId('Links')!
      workbook.addRows(sheetId, 0, 1)

      const exportedSnapshot = workbook.exportSnapshot()
      expect(exportedSnapshot.sheets[0]?.metadata?.hyperlinks).toEqual(expectedShiftedHyperlinks)

      const roundTrip = importXlsx(exportXlsx(exportedSnapshot), 'headless-hyperlink-structural-roundtrip.xlsx')
      expect(roundTrip.snapshot.sheets[0]?.metadata?.hyperlinks).toEqual(expectedShiftedHyperlinks)
    } finally {
      workbook.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel hyperlink metadata after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-hyperlink-structural-oracle-')
      try {
        const excelTruthPath = join(tempDir, 'excel-hyperlink-structural.xlsx')
        writeFileSync(excelTruthPath, buildHyperlinkWorkbookBytes())
        const excelTruth = runMacosExcelStructuralOperationOracle({
          workbookPath: excelTruthPath,
          worksheetName: 'Links',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A2', 'B3'],
          saveWorkbook: true,
        })
        const excelTruthImport = importXlsx(new Uint8Array(readFileSync(excelTruthPath)), 'excel-hyperlink-structural.xlsx')
        expect(excelTruthImport.snapshot.sheets[0]?.metadata?.hyperlinks).toEqual(expectedShiftedHyperlinks)

        const headlessWorkbook = WorkPaper.buildFromSnapshot(
          importXlsx(buildHyperlinkWorkbookBytes(), 'headless-hyperlink-structural.xlsx').snapshot,
          workbookConfig,
        )
        try {
          const sheetId = headlessWorkbook.getSheetId('Links')!
          headlessWorkbook.addRows(sheetId, 0, 1)
          const headlessPath = join(tempDir, 'headless-hyperlink-structural.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessWorkbook.exportSnapshot()))

          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Links',
            formulaCells: [],
            inspectCells: ['A2', 'B3'],
            saveWorkbook: true,
          })
          expect(headlessExcel.cells).toEqual(excelTruth.cells)

          const headlessImport = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-hyperlink-structural.xlsx')
          expect(headlessImport.snapshot.sheets[0]?.metadata?.hyperlinks).toEqual(expectedShiftedHyperlinks)
        } finally {
          headlessWorkbook.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    60_000,
  )
})

function buildHyperlinkWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const links = XLSX.utils.aoa_to_sheet([['Open report'], ['', 'Summary']])
  links['A1'] = {
    ...links['A1'],
    l: {
      Target: 'https://example.com/report',
      Tooltip: 'Open report',
    },
  }
  links['B2'] = {
    ...links['B2'],
    l: {
      Target: '#Summary!A1',
      Tooltip: 'Jump to summary',
    },
  }
  XLSX.utils.book_append_sheet(workbook, links, 'Links')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Destination']]), 'Summary')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}
