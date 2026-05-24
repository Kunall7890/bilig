import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { runXlsxFormulaRecalcCli } from '../cli-api.js'
import { WorkPaper, exportXlsx, importXlsx, parseQualifiedA1, recalculateSheetjsWorkbook, recalculateXlsx } from '../index.js'

describe('xlsx-formula-recalc', () => {
  it('edits XLSX inputs, recalculates formulas, and exports a reimportable workbook', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const sourceBytes = exportXlsx(sourceWorkbook.exportSnapshot())
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'pricing.xlsx',
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })

    expect(readNumber(result.reads['Summary!B2'])).toBe(72_000)
    expect(result.warnings).toEqual([])
    expect(result.changes.length).toBeGreaterThan(0)

    const imported = importXlsx(result.xlsx, 'pricing.recalculated.xlsx')
    const restored = WorkPaper.buildFromSnapshot(imported.snapshot)
    const summary = restored.getSheetId('Summary')
    expect(summary).toBeTypeOf('number')
    expect(readNumber(restored.getCellValue({ sheet: summary!, row: 1, col: 1 }))).toBe(72_000)
    restored.dispose()
  })

  it('recalculates formula cells written without cached formula values', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const sourceBytes = replaceCellXml(
      exportXlsx(sourceWorkbook.exportSnapshot()),
      'xl/worksheets/sheet2.xml',
      'B2',
      '<c r="B2"><f>Inputs!B2*Inputs!B3</f></c>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'pricing-without-formula-cache.xlsx',
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })

    expect(readNumber(result.reads['Summary!B2'])).toBe(72_000)
  })

  it('refreshes stale manual-calc caches for supported formulas before reads and export', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcMode="manual" fullCalcOnLoad="0"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'manual-stale-cache.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcMode')).toBe('manual')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullCalcOnLoad')).toBeNull()

    const imported = importXlsx(result.xlsx, 'manual-stale-cache.recalculated.xlsx')
    const restored = WorkPaper.buildFromSnapshot(imported.snapshot)
    const model = restored.getSheetId('Model')
    expect(model).toBeTypeOf('number')
    expect(readNumber(restored.getCellValue({ sheet: model!, row: 1, col: 1 }))).toBe(20)
    restored.dispose()
  })

  it('clears stale calculation metadata after explicit formula recalculation', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcMode="manual" calcOnSave="1" calcCompleted="0" fullCalcOnLoad="0" forceFullCalc="1"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'manual-stale-cache-flags.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcMode')).toBe('manual')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcOnSave')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcCompleted')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullCalcOnLoad')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'forceFullCalc')).toBeNull()
  })

  it('drops incomplete automatic calc metadata after explicit formula recalculation', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcCompleted="0"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'automatic-incomplete-cache.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    expect(readWorkbookXml(result.xlsx)).not.toContain('<calcPr')
  })

  it('carries import warnings when unsupported cached formulas recalculate to Excel errors', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [['AAPL', 0]],
    })
    const sourceBytes = replaceCellXml(
      exportXlsx(sourceWorkbook.exportSnapshot()),
      'xl/worksheets/sheet1.xml',
      'B1',
      '<c r="B1"><f>_xldudf_WISEPRICE(A1,&quot;Shares Outstanding&quot;)</f><v>14935800000</v></c>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'unsupported-cache-warning.xlsx',
      reads: ['Model!B1'],
    })

    expect(result.warnings).toContain(
      'Unsupported formulas were preserved from cached XLSX values; recalculation may return Excel error values.',
    )
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B1')).toBe('#NAME?')
  })

  it('preserves non-stale calculation preferences after explicit formula recalculation', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcMode="manual" fullPrecision="0" iterate="1" iterateCount="32" iterateDelta="0.001" concurrentCalc="0" calcCompleted="0" fullCalcOnLoad="0"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'manual-stale-cache-with-preferences.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcMode')).toBe('manual')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullPrecision')).toBe('0')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'iterate')).toBe('1')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'iterateCount')).toBe('32')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'iterateDelta')).toBe('0.001')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'concurrentCalc')).toBe('0')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcCompleted')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullCalcOnLoad')).toBeNull()
  })

  it('parses quoted sheet names and absolute A1 addresses', () => {
    expect(parseQualifiedA1("'Pricing Model'!$AB$12")).toEqual({
      sheetName: 'Pricing Model',
      row: 11,
      col: 27,
    })
  })

  it('exposes a SheetJS-named API and CLI alias from the live xlsx package', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const sourceBytes = exportXlsx(sourceWorkbook.exportSnapshot())
    sourceWorkbook.dispose()

    const result = recalculateSheetjsWorkbook(sourceBytes, {
      edits: [{ target: 'Inputs!B2', value: 48 }],
      reads: ['Summary!B2'],
    })
    expect(readNumber(result.reads['Summary!B2'])).toBe(57_600)

    let help = ''
    const exitCode = runXlsxFormulaRecalcCli(['--help'], {
      commandName: 'sheetjs-recalc',
      stdout: (text) => {
        help += text
      },
    })
    expect(exitCode).toBe(0)
    expect(help).toContain('Usage: sheetjs-recalc')
  })
})

function readNumber(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number') {
    return value.value
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
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

function readWorkbookXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array())
}

function readWorkbookCalcPrAttribute(bytes: Uint8Array, name: string): string | null {
  const calcPr = /<calcPr\b([^>]*)\/?>/u.exec(readWorkbookXml(bytes))?.[1] ?? ''
  const match = new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(calcPr)
  return match?.[1] ?? null
}

function readCachedFormulaValue(bytes: Uint8Array, sheetPath: string, address: string): string | null {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const match = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)
  return match?.[1] ?? null
}
