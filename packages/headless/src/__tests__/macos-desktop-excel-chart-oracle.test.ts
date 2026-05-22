import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { describe, expect, it } from 'vitest'

describe('macOS Desktop Excel chart oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips Bilig-authored chart series through Desktop Excel',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const engine = await buildChartWorkbook('desktop-excel-chart-oracle')
      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-chart-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-chart-oracle.xlsx')
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Dashboard',
          formulaCells: [],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const roundTrip = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-chart-oracle-saved.xlsx')
        expect(roundTrip.snapshot.workbook.metadata?.charts).toEqual([
          {
            id: 'RevenueTrend',
            sheetName: 'Dashboard',
            address: 'B2',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'C4' },
            chartType: 'line',
            rows: 12,
            cols: 8,
            title: 'Revenue trend',
            seriesOrientation: 'columns',
            firstRowAsHeaders: true,
            firstColumnAsLabels: true,
            legendPosition: 'bottom',
          },
        ])
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    120_000,
  )

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'imports chart source rewrites made by Desktop Excel structural edits',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const engine = await buildChartWorkbook('desktop-excel-chart-structural-oracle')
      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-chart-structure-oracle-'))
      try {
        const workbookPath = join(tempDir, 'headless-chart-structure-oracle.xlsx')
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        runMacosExcelStructuralOperationOracle({
          workbookPath,
          worksheetName: 'Data',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        const roundTrip = importXlsx(new Uint8Array(readFileSync(workbookPath)), 'headless-chart-structure-oracle-saved.xlsx')
        expect(roundTrip.snapshot.workbook.metadata?.charts).toEqual([
          {
            id: 'RevenueTrend',
            sheetName: 'Dashboard',
            address: 'B2',
            source: { sheetName: 'Data', startAddress: 'A2', endAddress: 'C5' },
            chartType: 'line',
            rows: 12,
            cols: 8,
            title: 'Revenue trend',
            seriesOrientation: 'columns',
            firstRowAsHeaders: true,
            firstColumnAsLabels: true,
            legendPosition: 'bottom',
          },
        ])
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    120_000,
  )
})

async function buildChartWorkbook(workbookName: string): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName })
  await engine.ready()
  engine.createSheet('Data')
  engine.createSheet('Dashboard')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C4' }, [
    ['Month', 'Revenue', 'Cost'],
    ['Jan', 120, 50],
    ['Feb', 135, 60],
    ['Mar', 150, 65],
  ])
  engine.setCellValue('Dashboard', 'A1', 'Chart host')
  engine.setChart({
    id: 'RevenueTrend',
    sheetName: 'Dashboard',
    address: 'B2',
    source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'C4' },
    chartType: 'line',
    rows: 12,
    cols: 8,
    title: 'Revenue trend',
    seriesOrientation: 'columns',
    firstRowAsHeaders: true,
    firstColumnAsLabels: true,
    legendPosition: 'bottom',
  })
  return engine
}
