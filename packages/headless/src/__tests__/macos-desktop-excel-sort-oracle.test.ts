import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const inspectedCells = [
  'A2',
  'B2',
  'C2',
  'D2',
  'A3',
  'B3',
  'C3',
  'D3',
  'A4',
  'B4',
  'C4',
  'D4',
  'A5',
  'B5',
  'C5',
  'D5',
  'A6',
  'B6',
  'C6',
  'D6',
  'E1',
] as const

const expectedSortedRows = [
  ['East', 50, 'invoice-005', 100],
  ['West', 40, 'invoice-002', 80],
  ['East', 30, 'invoice-003', 60],
  ['West', 20, 'invoice-004', 40],
  ['East', 10, 'invoice-001', 20],
] as const

const expectedSortedFormulas = [
  { address: 'D2', formula: 'B2*2' },
  { address: 'D3', formula: 'B3*2' },
  { address: 'D4', formula: 'B4*2' },
  { address: 'D5', formula: 'B5*2' },
  { address: 'D6', formula: 'B6*2' },
] as const

describe('macOS Desktop Excel sort oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel structural range sort semantics before preserving sort metadata',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-sort-oracle-'))
      try {
        const excelWorkbookPath = join(tempDir, 'excel-sort-oracle.xlsx')
        writeFileSync(excelWorkbookPath, exportXlsx(ledgerSnapshot()))

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Ledger',
          operations: [
            {
              kind: 'applySort',
              range: 'A1:D6',
              keys: [{ key: 'B1', order: 'descending' }],
              header: 'yes',
              orientation: 'rows',
            },
          ],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(excelRows(excelResult.cells)).toEqual(expectedSortedRows)
        expect(oracleFormulas(excelResult.cells)).toEqual(expectedSortedFormulas)
        expect(excelResult.cells.at(-1)?.value).toEqual({ kind: 'number', value: 50 })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-sort-oracle.xlsx')
        expect(importedRows(excelTruth.snapshot)).toEqual(expectedSortedRows)
        expect(snapshotFormulas(excelTruth.snapshot)).toEqual(expectedSortedFormulas)
        expect(styleIdAt(excelTruth.snapshot, 'Ledger', 'C2')).toBeDefined()
        expect(styleIdAt(excelTruth.snapshot, 'Ledger', 'C6')).toBeUndefined()
        expect(numberFormatAt(excelTruth.snapshot, 'Ledger', 'B2')).toBe('0.00')
        expect(numberFormatAt(excelTruth.snapshot, 'Ledger', 'B6')).toBeUndefined()
        expect(excelTruth.snapshot.sheets[0]?.metadata?.commentThreads).toEqual([
          {
            threadId: 'xlsx-comment:Ledger:C2',
            sheetName: 'Ledger',
            address: 'C2',
            comments: [{ id: 'xlsx-comment:Ledger:C2:1', body: 'Approved largest invoice', authorDisplayName: 'Audit' }],
          },
        ])
        expect(excelTruth.snapshot.sheets[0]?.metadata?.sorts?.[0]?.keys.length).toBeGreaterThan(0)

        const headless = new SpreadsheetEngine({ workbookName: 'headless-sort-oracle' })
        await headless.ready()
        headless.importSnapshot(ledgerSnapshot())
        expect(
          headless.sortRange(
            'Ledger',
            { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'D6' },
            [{ keyAddress: 'B1', direction: 'desc' }],
            { header: true },
          ),
        ).toBe(true)

        expect(engineRows(headless)).toEqual(expectedSortedRows)
        expect(engineFormulas(headless)).toEqual(expectedSortedFormulas)
        expect(headless.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 50 })
        expect(headless.getCellStyle(headless.getCell('Ledger', 'C2').styleId)?.fill?.backgroundColor).toBe('#fef3c7')
        expect(headless.getCell('Ledger', 'C6').styleId).toBeUndefined()
        expect(headless.getCell('Ledger', 'B2').format).toBe('0.00')
        expect(headless.getCell('Ledger', 'B6').format).toBeUndefined()
        expect(headless.getCommentThreads('Ledger')).toEqual([
          {
            threadId: 'thread-largest-invoice',
            sheetName: 'Ledger',
            address: 'C2',
            comments: [{ id: 'comment-largest-invoice', body: 'Approved largest invoice', authorDisplayName: 'Audit' }],
          },
        ])

        const headlessWorkbookPath = join(tempDir, 'headless-sort-oracle.xlsx')
        writeFileSync(headlessWorkbookPath, exportXlsx(headless.exportSnapshot()))
        const headlessExcel = runMacosExcelInspectionOracle({
          workbookPath: headlessWorkbookPath,
          worksheetName: 'Ledger',
          formulaCells: [],
          inspectCells: inspectedCells,
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(headlessExcel.cells.map(({ address, value }) => ({ address, value }))).toEqual(
          excelResult.cells.map(({ address, value }) => ({ address, value })),
        )
        expect(oracleFormulas(headlessExcel.cells)).toEqual(expectedSortedFormulas)
        const headlessExcelTruth = importXlsx(new Uint8Array(readFileSync(headlessWorkbookPath)), 'headless-sort-oracle.xlsx')
        expect(snapshotFormulas(headlessExcelTruth.snapshot)).toEqual(expectedSortedFormulas)
        expect(styleIdAt(headlessExcelTruth.snapshot, 'Ledger', 'C2')).toBeDefined()
        expect(styleIdAt(headlessExcelTruth.snapshot, 'Ledger', 'C6')).toBeUndefined()
        expect(numberFormatAt(headlessExcelTruth.snapshot, 'Ledger', 'B2')).toBe('0.00')
        expect(numberFormatAt(headlessExcelTruth.snapshot, 'Ledger', 'B6')).toBeUndefined()
        expect(headlessExcelTruth.snapshot.sheets[0]?.metadata?.commentThreads).toEqual([
          {
            threadId: 'xlsx-comment:Ledger:C2',
            sheetName: 'Ledger',
            address: 'C2',
            comments: [{ id: 'xlsx-comment:Ledger:C2:1', body: 'Approved largest invoice', authorDisplayName: 'Audit' }],
          },
        ])
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function ledgerSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Sort oracle',
      metadata: {
        styles: [{ id: 'style-largest-invoice', fill: { backgroundColor: '#fef3c7' } }],
        formats: [{ id: 'format-amount-decimal', code: '0.00', kind: 'number' }],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Ledger',
        order: 0,
        metadata: {
          styleRanges: [{ range: { sheetName: 'Ledger', startAddress: 'C6', endAddress: 'C6' }, styleId: 'style-largest-invoice' }],
          formatRanges: [{ range: { sheetName: 'Ledger', startAddress: 'B6', endAddress: 'B6' }, formatId: 'format-amount-decimal' }],
          commentThreads: [
            {
              threadId: 'thread-largest-invoice',
              sheetName: 'Ledger',
              address: 'C6',
              comments: [{ id: 'comment-largest-invoice', body: 'Approved largest invoice', authorDisplayName: 'Audit' }],
            },
          ],
        },
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'C1', value: 'Invoice' },
          { address: 'D1', value: 'Double' },
          { address: 'A2', value: 'East' },
          { address: 'B2', value: 10 },
          { address: 'C2', value: 'invoice-001' },
          { address: 'D2', formula: 'B2*2', value: 20 },
          { address: 'A3', value: 'West' },
          { address: 'B3', value: 40 },
          { address: 'C3', value: 'invoice-002' },
          { address: 'D3', formula: 'B3*2', value: 80 },
          { address: 'A4', value: 'East' },
          { address: 'B4', value: 30 },
          { address: 'C4', value: 'invoice-003' },
          { address: 'D4', formula: 'B4*2', value: 60 },
          { address: 'A5', value: 'West' },
          { address: 'B5', value: 20 },
          { address: 'C5', value: 'invoice-004' },
          { address: 'D5', formula: 'B5*2', value: 40 },
          { address: 'A6', value: 'East' },
          { address: 'B6', value: 50 },
          { address: 'C6', value: 'invoice-005' },
          { address: 'D6', formula: 'B6*2', value: 100 },
          { address: 'E1', formula: 'B2', value: 10 },
        ],
      },
    ],
  }
}

function styleIdAt(snapshot: WorkbookSnapshot, sheetName: string, address: string): string | undefined {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  const parsed = parseA1(address)
  return sheet?.metadata?.styleRanges?.find((record) =>
    rangeContains(record.range.startAddress, record.range.endAddress, parsed.row, parsed.col),
  )?.styleId
}

function formatIdAt(snapshot: WorkbookSnapshot, sheetName: string, address: string): string | undefined {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  const parsed = parseA1(address)
  return sheet?.metadata?.formatRanges?.find((record) =>
    rangeContains(record.range.startAddress, record.range.endAddress, parsed.row, parsed.col),
  )?.formatId
}

function numberFormatAt(snapshot: WorkbookSnapshot, sheetName: string, address: string): string | undefined {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  const cellFormat = sheet?.cells.find((cell) => cell.address === address)?.format
  if (cellFormat !== undefined) {
    return cellFormat
  }
  const formatId = formatIdAt(snapshot, sheetName, address)
  return snapshot.workbook.metadata?.formats?.find((format) => format.id === formatId)?.code
}

function rangeContains(startAddress: string, endAddress: string, row: number, col: number): boolean {
  const start = parseA1(startAddress)
  const end = parseA1(endAddress)
  return (
    row >= Math.min(start.row, end.row) &&
    row <= Math.max(start.row, end.row) &&
    col >= Math.min(start.col, end.col) &&
    col <= Math.max(start.col, end.col)
  )
}

function parseA1(address: string): { readonly row: number; readonly col: number } {
  const match = /^([A-Z]+)([1-9]\d*)$/u.exec(address)
  if (!match) {
    throw new Error(`Invalid test address: ${address}`)
  }
  const colLabel = match[1]
  let col = 0
  for (let index = 0; index < colLabel.length; index += 1) {
    col = col * 26 + colLabel.charCodeAt(index) - 64
  }
  return {
    row: Number.parseInt(match[2], 10),
    col,
  }
}

function excelRows(cells: readonly { readonly value: { readonly value?: unknown } }[]): unknown[][] {
  return [0, 4, 8, 12, 16].map((start) => [
    cells[start]?.value.value,
    cells[start + 1]?.value.value,
    cells[start + 2]?.value.value,
    cells[start + 3]?.value.value,
  ])
}

function importedRows(snapshot: WorkbookSnapshot): unknown[][] {
  const sheet = snapshot.sheets[0]
  const values = new Map(sheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])
  return [
    [values.get('A2'), values.get('B2'), values.get('C2'), values.get('D2')],
    [values.get('A3'), values.get('B3'), values.get('C3'), values.get('D3')],
    [values.get('A4'), values.get('B4'), values.get('C4'), values.get('D4')],
    [values.get('A5'), values.get('B5'), values.get('C5'), values.get('D5')],
    [values.get('A6'), values.get('B6'), values.get('C6'), values.get('D6')],
  ]
}

function engineRows(engine: SpreadsheetEngine): unknown[][] {
  return engine.getRangeValues({ sheetName: 'Ledger', startAddress: 'A2', endAddress: 'D6' }).map((row) => row.map(cellValue))
}

function oracleFormulas(
  cells: readonly { readonly address: string; readonly formula?: string }[],
): readonly { readonly address: string; readonly formula: string }[] {
  return expectedSortedFormulas.map(({ address }) => ({
    address,
    formula: normalizeFormula(cells.find((cell) => cell.address === address)?.formula),
  }))
}

function snapshotFormulas(snapshot: WorkbookSnapshot): readonly { readonly address: string; readonly formula: string }[] {
  const formulas = new Map(snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell.formula]) ?? [])
  return expectedSortedFormulas.map(({ address }) => ({
    address,
    formula: normalizeFormula(formulas.get(address)),
  }))
}

function engineFormulas(engine: SpreadsheetEngine): readonly { readonly address: string; readonly formula: string }[] {
  return expectedSortedFormulas.map(({ address }) => ({
    address,
    formula: normalizeFormula(engine.getCell('Ledger', address).formula),
  }))
}

function normalizeFormula(formula: string | undefined): string {
  if (formula === undefined) {
    throw new Error('Missing expected formula')
  }
  return formula.startsWith('=') ? formula.slice(1) : formula
}

function cellValue(value: CellValue): unknown {
  switch (value.tag) {
    case ValueTag.Number:
    case ValueTag.String:
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Empty:
    case ValueTag.Error:
      return null
    default:
      return null
  }
}
