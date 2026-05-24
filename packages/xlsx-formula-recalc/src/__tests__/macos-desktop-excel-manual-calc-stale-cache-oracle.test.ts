import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { isMacosExcelInstalled, runMacosExcelInspectionOracle } from '@bilig/excel-fixtures'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper, exportXlsx, recalculateXlsx } from '../index.js'

describe('macOS Desktop Excel manual-calc stale-cache oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel full recalculation for stale manual-calc cached formulas',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-xlsx-recalc-stale-cache-oracle-')
      try {
        const sourceBytes = buildManualCalcStaleCacheXlsx()
        expect(readCachedFormulaValue(sourceBytes, 'xl/worksheets/sheet1.xml', 'B2')).toBe('999')

        const sourcePath = join(tempDir, 'manual-stale-cache-source.xlsx')
        writeFileSync(sourcePath, sourceBytes)

        const excelTruth = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: ['B2'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelTruth.excelVersion).toMatch(/^\d+\./u)
        expect(excelTruth.cells[0]).toMatchObject({
          address: 'B2',
          formula: '=A2*10',
          value: { kind: 'number', value: 20 },
        })
        expect(readCachedFormulaValue(new Uint8Array(readFileSync(sourcePath)), 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')

        const result = recalculateXlsx(sourceBytes, {
          fileName: 'manual-stale-cache-source.xlsx',
          reads: ['Model!B2'],
        })
        expect(numberCell(result.reads['Model!B2'])).toBe(20)
        expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')

        const headlessPath = join(tempDir, 'manual-stale-cache-headless.xlsx')
        writeFileSync(headlessPath, result.xlsx)
        const excelHeadless = runMacosExcelInspectionOracle({
          workbookPath: headlessPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: ['B2'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelHeadless.cells[0]).toMatchObject(excelTruth.cells[0])
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    180_000,
  )
})

function buildManualCalcStaleCacheXlsx(): Uint8Array {
  const sourceWorkbook = WorkPaper.buildFromSheets({
    Model: [
      ['Input', 'Output'],
      [2, '=A2*10'],
    ],
  })
  try {
    return setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcMode="manual" fullCalcOnLoad="0"/>',
    )
  } finally {
    sourceWorkbook.dispose()
  }
}

function numberCell(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number') {
    return value.value
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}

function createExcelAccessibleTempDir(prefix: string): string {
  return mkdtempSync(join(tmpdir(), prefix))
}

function replaceCellXml(bytes: Uint8Array, sheetPath: string, address: string, replacement: string): Uint8Array {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const pattern = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u')
  if (!pattern.test(sheetXml)) {
    throw new Error(`Missing cell XML for ${address}`)
  }
  zip[sheetPath] = strToU8(sheetXml.replace(pattern, replacement))
  return zipSync(zip)
}

function setWorkbookCalcPr(bytes: Uint8Array, calcPrXml: string): Uint8Array {
  const zip = unzipSync(bytes)
  const workbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  if (/<calcPr\b[\s\S]*?\/>/u.test(workbookXml)) {
    zip['xl/workbook.xml'] = strToU8(workbookXml.replace(/<calcPr\b[\s\S]*?\/>/u, calcPrXml))
    return zipSync(zip)
  }
  zip['xl/workbook.xml'] = strToU8(workbookXml.replace('</workbook>', `${calcPrXml}</workbook>`))
  return zipSync(zip)
}

function readCachedFormulaValue(bytes: Uint8Array, sheetPath: string, address: string): string | null {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const match = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)
  return match?.[1] ?? null
}
