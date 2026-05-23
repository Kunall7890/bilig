import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

describe('macOS Desktop Excel sheet move metadata topology oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel worksheet metadata ownership after moving a sheet tab',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-sheet-move-metadata-oracle-')
      try {
        const sourceBytes = exportXlsx(sheetMoveMetadataTopologySnapshot())
        const importedSource = importXlsx(sourceBytes, 'sheet-move-metadata-source.xlsx').snapshot
        expect(metadataCodeNames(importedSource)).toEqual(['Data:DataCode', 'Inputs:InputsCode', 'Report:ReportCode'])

        const excelWorkbookPath = join(tempDir, 'excel-sheet-move-metadata-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'moveSheet', name: 'Report', before: 'Inputs' }],
          inspectCells: ['A1', 'A2'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: 'Report' },
          { kind: 'number', value: 33 },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-sheet-move-metadata-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
        expect(metadataCodeNames(excelTruth.snapshot)).toEqual(['Data:DataCode', 'Report:ReportCode', 'Inputs:InputsCode'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const reportSheet = workpaper.getSheetId('Report')
          if (reportSheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.moveSheet(reportSheet, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
          expect(metadataCodeNames(headlessSnapshot)).toEqual(metadataCodeNames(excelTruth.snapshot))
          expect(metadataTabColors(headlessSnapshot)).toEqual(metadataTabColors(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'headless-sheet-move-metadata.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1', 'A2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-sheet-move-metadata-truth.xlsx')
          expect(metadataCodeNames(headlessTruth.snapshot)).toEqual(metadataCodeNames(excelTruth.snapshot))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )
})

function sheetMoveMetadataTopologySnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel sheet move metadata topology oracle',
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        metadata: {
          sheetPr: { xml: '<sheetPr codeName="DataCode"><outlinePr summaryBelow="0"/></sheetPr>' },
          tabColor: { rgb: 'FF808080' },
        },
        cells: [
          { address: 'A1', value: 'Data' },
          { address: 'A2', value: 11 },
        ],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        metadata: {
          sheetPr: { xml: '<sheetPr codeName="InputsCode"><outlinePr summaryRight="0"/></sheetPr>' },
          tabColor: { rgb: 'FFFF0000' },
        },
        cells: [
          { address: 'A1', value: 'Inputs' },
          { address: 'A2', value: 22 },
        ],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        metadata: {
          sheetPr: { xml: '<sheetPr codeName="ReportCode"><pageSetUpPr fitToPage="1"/></sheetPr>' },
          tabColor: { rgb: 'FF00AA00' },
        },
        cells: [
          { address: 'A1', value: 'Report' },
          { address: 'A2', value: 33 },
        ],
      },
    ],
  }
}

function metadataCodeNames(snapshot: WorkbookSnapshot): string[] {
  return snapshot.sheets.map((sheet) => `${sheet.name}:${codeName(sheet.metadata?.sheetPr?.xml) ?? ''}`)
}

function metadataTabColors(snapshot: WorkbookSnapshot): string[] {
  return snapshot.sheets.map((sheet) => `${sheet.name}:${JSON.stringify(sheet.metadata?.tabColor ?? null)}`)
}

function codeName(sheetPrXml: string | undefined): string | undefined {
  return /\bcodeName="([^"]+)"/u.exec(sheetPrXml ?? '')?.[1]
}
