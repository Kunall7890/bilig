import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

describe('macOS Desktop Excel preserved package metadata oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'preserves Desktop Excel workbook and sheet view state after a headless edit',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-package-metadata-oracle-')
      try {
        const sourcePath = join(tempDir, 'view-state-source.xlsx')
        writeFileSync(sourcePath, exportXlsx(viewStateSnapshot()))

        const excelInitial = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Report',
          formulaCells: [],
          inspectCells: ['A1', 'B1'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelInitial.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: 'view-state' },
          { kind: 'string', value: '' },
        ])

        const excelTruth = importXlsx(new Uint8Array(readFileSync(sourcePath)), 'excel-saved-view-state-source.xlsx')
        expect(viewStateSummary(excelTruth.snapshot)).toEqual({
          workbookView: true,
          sheetView: true,
        })

        const workpaper = WorkPaper.buildFromSnapshot(excelTruth.snapshot)
        try {
          const sheet = workpaper.getSheetId('Report')
          if (sheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.setCellContents({ sheet, row: 0, col: 1 }, 'headless edit')

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(viewStateXml(headlessSnapshot)).toEqual(viewStateXml(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'view-state-headless-edit.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const excelHeadless = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1', 'B1'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(excelHeadless.cells.map((cell) => cell.value)).toEqual([
            { kind: 'string', value: 'view-state' },
            { kind: 'string', value: 'headless edit' },
          ])

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'excel-saved-view-state-headless.xlsx')
          expect(viewStateSummary(headlessTruth.snapshot)).toEqual(viewStateSummary(excelTruth.snapshot))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function viewStateSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel package metadata oracle',
      metadata: {
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>',
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        metadata: {
          viewState: {
            sheetViewsXml: '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>',
          },
        },
        cells: [{ address: 'A1', value: 'view-state' }],
      },
    ],
  }
}

function viewStateXml(snapshot: WorkbookSnapshot): { readonly workbook: string | undefined; readonly sheet: string | undefined } {
  return {
    workbook: snapshot.workbook.metadata?.viewState?.bookViewsXml,
    sheet: snapshot.sheets[0]?.metadata?.viewState?.sheetViewsXml,
  }
}

function viewStateSummary(snapshot: WorkbookSnapshot): { readonly workbookView: boolean; readonly sheetView: boolean } {
  const xml = viewStateXml(snapshot)
  return {
    workbookView: xml.workbook?.includes('<workbookView') ?? false,
    sheetView: xml.sheet?.includes('<sheetView') ?? false,
  }
}
