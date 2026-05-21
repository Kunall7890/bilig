import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle } from '@bilig/excel-fixtures'
import { describe, expect, it } from 'vitest'

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
              formula: '=GETPIVOTDATA("Sales Total",$B$2)',
            },
          ],
          inspectCells: ['H1', 'H2', 'H3'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelResult.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'H1', value: { kind: 'number', value: 10 } },
          { address: 'H2', value: { kind: 'number', value: 7 } },
          { address: 'H3', value: { kind: 'number', value: 31 } },
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
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    120_000,
  )
})
