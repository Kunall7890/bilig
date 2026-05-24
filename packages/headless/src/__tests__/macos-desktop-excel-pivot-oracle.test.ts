import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const expectedEastSales = { tag: ValueTag.Number, value: 25 }
const expectedExcelEastSales = { kind: 'number', value: 25 }

describe('macOS Desktop Excel pivot oracle', () => {
  it('refreshes semantic source-backed pivots and exports Desktop Excel-refreshable caches', async () => {
    const engine = await buildEditedPivotEngine()
    const exportedSnapshot = engine.exportSnapshot()

    expect(engine.getCellValue('Pivot', 'E2')).toEqual(expectedEastSales)
    expect(exportedSnapshot.workbook.metadata?.pivots?.[0]).toMatchObject({
      cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
      cachedRecords: [
        ['East', 'priority', 'Widget', 20],
        ['West', 'priority', 'Widget', 7],
        ['East', 'priority', 'Gizmo', 5],
      ],
      rows: 3,
      cols: 2,
    })

    const exportedBytes = exportXlsx(exportedSnapshot)
    const exportedZip = unzipSync(exportedBytes)
    const cacheDefinitionXml = strFromU8(exportedZip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array())
    expect(cacheDefinitionXml).toContain('invalid="1"')
    expect(cacheDefinitionXml).toContain('refreshOnLoad="1"')
    expect(cacheDefinitionXml).toContain('recordCount="0"')
    expect(exportedZip['xl/pivotCache/pivotCacheRecords1.xml']).toBeUndefined()

    const roundTrip = importXlsx(exportedBytes, 'headless-pivot-refresh-roundtrip.xlsx')
    expect(roundTrip.snapshot.workbook.metadata?.pivots?.[0]).toMatchObject({
      cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
    })
    expect(roundTrip.snapshot.workbook.metadata?.pivots?.[0]?.cachedRecords).toBeUndefined()
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel GETPIVOTDATA after headless source edits and XLSX export',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const engine = await buildEditedPivotEngine()
      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-oracle-')
      try {
        const workbookPath = join(tempDir, 'headless-pivot-refresh-oracle.xlsx')
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excel = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Pivot',
          formulaCells: [],
          inspectCells: ['E2'],
          refreshWorkbook: true,
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excel.cells.map((cell) => cell.value)).toEqual([expectedExcelEastSales])

        const saved = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'excel-saved-pivot-refresh-oracle.xlsx')
        expect(saved.snapshot.workbook.metadata?.pivots?.[0]?.cachedRecords).toContainEqual(['East', 'priority', 'Widget', 20])
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel pivot package cleanup after deleting a pivot sheet',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-pivot-delete-oracle-')
      try {
        const sourceBytes = exportXlsx(importedSourceBackedPivotSnapshot())
        const importedSource = importXlsx(sourceBytes, 'pivot-delete-source.xlsx').snapshot
        expect(pivotPackageSummary(importedSource)).toEqual({
          pivotParts: ['xl/pivotCache/pivotCacheDefinition1.xml', 'xl/pivotTables/pivotTable1.xml'],
          workbookCacheIds: ['1'],
          workbookRelationshipTargets: ['pivotCache/pivotCacheDefinition1.xml'],
          pivotSheets: ['Pivot'],
          semanticPivotSheets: ['Pivot'],
        })

        const excelWorkbookPath = join(tempDir, 'excel-pivot-delete-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'deleteSheet', name: 'Pivot' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-pivot-delete-truth.xlsx')
        expect(excelTruth.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
        const excelSummary = pivotPackageSummary(excelTruth.snapshot)
        expect(excelSummary).toEqual({
          pivotParts: [],
          workbookCacheIds: [],
          workbookRelationshipTargets: [],
          pivotSheets: [],
          semanticPivotSheets: [],
        })

        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const pivotSheet = workpaper.getSheetId('Pivot')
          if (pivotSheet === undefined) {
            throw new Error('Expected Pivot sheet')
          }
          workpaper.removeSheet(pivotSheet)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
          expect(pivotPackageSummary(headlessSnapshot)).toEqual(excelSummary)

          const headlessWorkbookPath = join(tempDir, 'headless-pivot-delete.xlsx')
          writeFileSync(headlessWorkbookPath, exportXlsx(headlessSnapshot))
          runMacosExcelInspectionOracle({
            workbookPath: headlessWorkbookPath,
            worksheetName: 'Data',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 120_000,
          })
          const headlessExcelTruth = importXlsx(
            new Uint8Array(readFileSync(headlessWorkbookPath)),
            'headless-excel-pivot-delete-truth.xlsx',
          )
          expect(pivotPackageSummary(headlessExcelTruth.snapshot)).toEqual(excelSummary)
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

async function buildEditedPivotEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'headless-pivot-refresh-oracle' })
  await engine.ready()
  engine.importSnapshot(importedSourceBackedPivotSnapshot())
  engine.setCellValue('Data', 'D2', 20)
  engine.setCellFormula('Pivot', 'E2', 'GETPIVOTDATA("Sum of Sales",$B$2,"Region","East")')
  return engine
}

function importedSourceBackedPivotSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Headless pivot refresh oracle',
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
              ['East', 'priority', 'Widget', 1],
              ['West', 'priority', 'Widget', 2],
            ],
            rows: 1,
            cols: 1,
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
          { address: 'D2', value: 10 },
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

function pivotPackageSummary(snapshot: WorkbookSnapshot): {
  readonly pivotParts: readonly string[]
  readonly workbookCacheIds: readonly string[]
  readonly workbookRelationshipTargets: readonly string[]
  readonly pivotSheets: readonly string[]
  readonly semanticPivotSheets: readonly string[]
} {
  const metadata = snapshot.workbook.metadata
  const workbookPivotCachesXml = metadata?.pivotArtifacts?.workbookPivotCachesXml ?? ''
  return {
    pivotParts: metadata?.pivotArtifacts?.parts.map((part) => part.path).toSorted() ?? [],
    workbookCacheIds: [...workbookPivotCachesXml.matchAll(/<pivotCache\b[^>]*\bcacheId="([^"]+)"/gu)]
      .map((match) => match[1] ?? '')
      .filter(Boolean),
    workbookRelationshipTargets:
      metadata?.pivotArtifacts?.workbookRelationships?.map((relationship) => relationship.target).toSorted() ?? [],
    pivotSheets: snapshot.sheets.flatMap((sheet) => (sheet.metadata?.pivotArtifacts ? [sheet.name] : [])).toSorted(),
    semanticPivotSheets: metadata?.pivots?.map((pivot) => pivot.sheetName).toSorted() ?? [],
  }
}
