import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle } from '@bilig/excel-fixtures'
import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

describe('macOS Desktop Excel pivot oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'exports refreshable pivot tables whose GETPIVOTDATA values match Desktop Excel',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const engine = new SpreadsheetEngine({ workbookName: 'desktop-excel-pivot-oracle' })
      await engine.ready()
      engine.createSheet('Data')
      engine.createSheet('Pivot')
      engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C5' }, [
        ['Region', 'Quarter', 'Sales'],
        ['East', 'Q1', 10],
        ['East', 'Q2', 5],
        ['West', 'Q2', 7],
        ['East', 'Q3', 9],
      ])
      engine.setPivotTable('Pivot', 'B2', {
        name: 'SalesByRegionQuarter',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'C5' },
        groupBy: ['Region'],
        columnFields: ['Quarter'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
      })

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-pivot-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-pivot-oracle.xlsx')
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        const excelResult = runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Pivot',
          refreshAll: true,
          formulaCells: [
            {
              address: 'H1',
              formula: '=GETPIVOTDATA("Sales Total",$B$2,"Region","East","Quarter","Q1")',
            },
            {
              address: 'H2',
              formula: '=GETPIVOTDATA("Sales Total",$B$2,"Region","West","Quarter","Q2")',
            },
            {
              address: 'H3',
              formula: '=GETPIVOTDATA("Sales Total",$B$2,"Region","East","Quarter","Q3")',
            },
          ],
          inspectCells: ['H1', 'H2', 'H3'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'H1', value: { kind: 'number', value: 10 } },
          { address: 'H2', value: { kind: 'number', value: 7 } },
          { address: 'H3', value: { kind: 'number', value: 9 } },
        ])

        const roundTrip = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-pivot-oracle-saved.xlsx')
        expect(roundTrip.snapshot.workbook.metadata?.pivots?.[0]).toMatchObject({
          name: 'SalesByRegionQuarter',
          sheetName: 'Pivot',
          address: 'B2',
          groupBy: ['Region'],
          columnFields: ['Quarter'],
          values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
        })
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'refreshes source-backed imported pivot values after headless edits and exports Excel-refreshable caches',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const engine = new SpreadsheetEngine({ workbookName: 'desktop-excel-pivot-stale-cache-oracle' })
      await engine.ready()
      engine.createSheet('Data')
      engine.createSheet('Pivot')
      engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C5' }, [
        ['Region', 'Quarter', 'Sales'],
        ['East', 'Q1', 10],
        ['East', 'Q2', 5],
        ['West', 'Q2', 7],
        ['East', 'Q3', 9],
      ])
      engine.setPivotTable('Pivot', 'B2', {
        name: 'SalesByRegionQuarter',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'C5' },
        groupBy: ['Region'],
        columnFields: ['Quarter'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales Total' }],
      })

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-pivot-stale-cache-oracle-'))
      try {
        const excelAuthoredPath = join(tempDir, 'excel-authored-pivot-cache.xlsx')
        writeFileSync(excelAuthoredPath, exportXlsx(engine.exportSnapshot()))

        const excelAuthored = runMacosExcelInspectionOracle({
          workbookPath: excelAuthoredPath,
          worksheetName: 'Pivot',
          refreshAll: true,
          formulaCells: [
            {
              address: 'H1',
              formula: '=GETPIVOTDATA("Sales Total",$B$2,"Region","East","Quarter","Q1")',
            },
            {
              address: 'H2',
              formula: '=GETPIVOTDATA("Sales Total",$B$2,"Region","East","Quarter","Q3")',
            },
          ],
          inspectCells: ['H1', 'H2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelAuthored.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'H1', value: { kind: 'number', value: 10 } },
          { address: 'H2', value: { kind: 'number', value: 9 } },
        ])

        const imported = importXlsx(new Uint8Array(readFileSync(excelAuthoredPath)), 'excel-authored-pivot-cache.xlsx')
        expect(imported.snapshot.workbook.metadata?.pivots?.[0]?.cachedRecords).toEqual([
          ['East', 'Q1', 10],
          ['East', 'Q2', 5],
          ['West', 'Q2', 7],
          ['East', 'Q3', 9],
        ])

        const headless = new SpreadsheetEngine({ workbookName: 'desktop-excel-pivot-stale-cache-roundtrip' })
        await headless.ready()
        headless.importSnapshot(imported.snapshot)
        headless.setCellFormula('Pivot', 'H1', 'GETPIVOTDATA("Sales Total",B2,"Region","East","Quarter","Q1")')
        headless.setCellFormula('Pivot', 'H2', 'GETPIVOTDATA("Sales Total",B2,"Region","East","Quarter","Q3")')
        headless.setCellValue('Data', 'C2', 100)

        expect(headless.getCellValue('Pivot', 'H1')).toEqual({ tag: ValueTag.Number, value: 100 })
        expect(headless.getCellValue('Pivot', 'H2')).toEqual({ tag: ValueTag.Number, value: 9 })

        const headlessExportPath = join(tempDir, 'headless-pivot-cache-roundtrip.xlsx')
        writeFileSync(headlessExportPath, exportXlsx(headless.exportSnapshot()))
        const excelRoundTrip = runMacosExcelInspectionOracle({
          workbookPath: headlessExportPath,
          worksheetName: 'Pivot',
          refreshAll: true,
          formulaCells: [
            {
              address: 'H1',
              formula: '=GETPIVOTDATA("Sales Total",$B$2,"Region","East","Quarter","Q1")',
            },
            {
              address: 'H2',
              formula: '=GETPIVOTDATA("Sales Total",$B$2,"Region","East","Quarter","Q3")',
            },
          ],
          inspectCells: ['H1', 'H2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelRoundTrip.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'H1', value: { kind: 'number', value: 100 } },
          { address: 'H2', value: { kind: 'number', value: 9 } },
        ])

        const roundTrip = importXlsx(new Uint8Array(readFileSync(headlessExportPath)), 'headless-pivot-cache-roundtrip.xlsx')
        expect(roundTrip.snapshot.workbook.metadata?.pivots?.[0]?.cachedRecords).toEqual([
          ['East', 'Q1', 100],
          ['East', 'Q2', 5],
          ['West', 'Q2', 7],
          ['East', 'Q3', 9],
        ])
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )
})
