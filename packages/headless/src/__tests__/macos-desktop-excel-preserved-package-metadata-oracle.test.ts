import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
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

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel preserved style artifacts and view refs after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-package-metadata-structure-oracle-')
      try {
        const excelWorkbookPath = join(tempDir, 'excel-preserved-package-structure-source.xlsx')
        const sourceBytes = exportXlsx(structuralPreservedMetadataSnapshot())
        writeFileSync(excelWorkbookPath, sourceBytes)

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Report',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'B3', 'C4'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-preserved-package-structure-truth.xlsx')
        expect(styleArtifactRefs(excelTruth.snapshot)).toEqual({
          styled: ['B3', 'C4'],
          blank: [],
        })
        expect(sheetViewRefs(excelTruth.snapshot)).toEqual({
          topLeftCell: 'B2',
          activeCell: 'C3',
          sqref: 'C3 D4:E5',
        })

        const workpaper = WorkPaper.buildFromSnapshot(importXlsx(sourceBytes, 'headless-preserved-package-structure-source.xlsx').snapshot)
        try {
          const sheet = workpaper.getSheetId('Report')
          if (sheet === undefined) {
            throw new Error('Expected Report sheet')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(styleArtifactRefs(headlessSnapshot)).toEqual(styleArtifactRefs(excelTruth.snapshot))
          expect(sheetViewRefs(headlessSnapshot)).toEqual(sheetViewRefs(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'headless-preserved-package-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Report',
            formulaCells: [],
            inspectCells: ['A1', 'B3', 'C4'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-preserved-package-structure-truth.xlsx')
          expect(styleArtifactRefs(headlessTruth.snapshot)).toEqual(styleArtifactRefs(excelTruth.snapshot))
          expect(sheetViewRefs(headlessTruth.snapshot)).toEqual(sheetViewRefs(excelTruth.snapshot))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel preserved pivot package refs after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-package-structure-oracle-')
      try {
        const sourceBytes = exportXlsx(pivotPackageStructuralSourceSnapshot())
        const importedSource = importXlsx(sourceBytes, 'pivot-package-structure-source.xlsx').snapshot
        expect(pivotLocationRefs(importedSource)).toEqual(['B2:C5'])

        const excelWorkbookPath = join(tempDir, 'excel-pivot-package-structure-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Pivot',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-pivot-package-structure-truth.xlsx')
        expect(pivotLocationRefs(excelTruth.snapshot)).toEqual(['B3:C6'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const pivotSheet = workpaper.getSheetId('Pivot')
          if (pivotSheet === undefined) {
            throw new Error('Expected Pivot sheet')
          }
          workpaper.addRows(pivotSheet, 0, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(pivotLocationRefs(headlessSnapshot)).toEqual(pivotLocationRefs(excelTruth.snapshot))

          const headlessPath = join(tempDir, 'headless-pivot-package-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Pivot',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-pivot-package-structure-truth.xlsx')
          expect(pivotLocationRefs(headlessTruth.snapshot)).toEqual(pivotLocationRefs(excelTruth.snapshot))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel preserved pivot cache source refs after structural source row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-cache-source-structure-oracle-')
      try {
        const sourceBytes = exportXlsx(pivotPackageStructuralSourceSnapshot())
        const importedSource = importXlsx(sourceBytes, 'pivot-cache-source-structure-source.xlsx').snapshot
        expect(pivotCacheSourceRefs(importedSource)).toEqual(['A1:D4'])

        const excelWorkbookPath = join(tempDir, 'excel-pivot-cache-source-structure-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-pivot-cache-source-structure-truth.xlsx')
        const excelSourceRefs = pivotCacheSourceRefs(excelTruth.snapshot)
        expect(excelSourceRefs).not.toEqual(['A1:D4'])

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.addRows(dataSheet, 0, 1)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(pivotCacheSourceRefs(headlessSnapshot)).toEqual(excelSourceRefs)

          const headlessPath = join(tempDir, 'headless-pivot-cache-source-structure.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-pivot-cache-source-structure-truth.xlsx')
          expect(pivotCacheSourceRefs(headlessTruth.snapshot)).toEqual(excelSourceRefs)
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

function structuralPreservedMetadataSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel structural package metadata oracle',
      metadata: {
        styleArtifacts: {
          stylesXml: headerStyleReferenceStylesXml,
        },
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
          styleArtifacts: {
            cellStyleIndexes: [
              { address: 'B2', styleIndex: 1 },
              { address: 'C3', styleIndex: 1 },
            ],
            blankCellAddresses: ['C3'],
          },
          viewState: {
            sheetViewsXml:
              '<sheetViews><sheetView workbookViewId="0" topLeftCell="B2" tabSelected="1"><selection activeCell="C3" sqref="C3 D4:E5"/></sheetView></sheetViews>',
          },
        },
        cells: [
          { address: 'A1', value: 'preserved-structure' },
          { address: 'B2', value: 'styled' },
        ],
      },
    ],
  }
}

function pivotPackageStructuralSourceSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel pivot package structural oracle',
      metadata: {
        pivots: [
          {
            name: 'SalesByRegion',
            sheetName: 'Pivot',
            address: 'B2',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
            sourceKind: 'worksheet',
            groupBy: ['Region'],
            values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
            cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
            cachedRecords: [
              ['East', 'priority', 'Widget', 20],
              ['West', 'priority', 'Widget', 7],
              ['East', 'priority', 'Gizmo', 5],
            ],
            rows: 4,
            cols: 2,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Notes' },
          { address: 'C1', value: 'Product' },
          { address: 'D1', value: 'Sales' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 'priority' },
          { address: 'C2', value: 'Widget' },
          { address: 'D2', value: 20 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 'priority' },
          { address: 'C3', value: 'Widget' },
          { address: 'D3', value: 7 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 'priority' },
          { address: 'C4', value: 'Gizmo' },
          { address: 'D4', value: 5 },
        ],
      },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        cells: [],
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

function styleArtifactRefs(snapshot: WorkbookSnapshot): { readonly styled: string[]; readonly blank: string[] } {
  const artifacts = snapshot.sheets[0]?.metadata?.styleArtifacts
  return {
    styled: (artifacts?.cellStyleIndexes ?? []).map((entry) => entry.address).toSorted(),
    blank: (artifacts?.blankCellAddresses ?? []).toSorted(),
  }
}

function sheetViewRefs(snapshot: WorkbookSnapshot): {
  readonly topLeftCell: string | undefined
  readonly activeCell: string | undefined
  readonly sqref: string | undefined
} {
  const xml = snapshot.sheets[0]?.metadata?.viewState?.sheetViewsXml ?? ''
  return {
    topLeftCell: readXmlAttribute(xml, 'topLeftCell'),
    activeCell: readXmlAttribute(xml, 'activeCell'),
    sqref: readXmlAttribute(xml, 'sqref'),
  }
}

function pivotLocationRefs(snapshot: WorkbookSnapshot): string[] {
  return (snapshot.workbook.metadata?.pivotArtifacts?.parts ?? [])
    .filter((part) => part.path.startsWith('xl/pivotTables/'))
    .map((part) => readXmlAttribute(part.xml, 'ref'))
    .filter((ref): ref is string => ref !== undefined)
    .toSorted()
}

function pivotCacheSourceRefs(snapshot: WorkbookSnapshot): string[] {
  return (snapshot.workbook.metadata?.pivotArtifacts?.parts ?? [])
    .filter((part) => part.path.startsWith('xl/pivotCache/pivotCacheDefinition'))
    .map((part) => readXmlAttribute(part.xml, 'ref'))
    .filter((ref): ref is string => ref !== undefined)
    .toSorted()
}

function readXmlAttribute(xml: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(xml)?.[1]
}

const headerStyleReferenceStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')
