import { describe, expect, it } from 'vitest'
import { strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { importXlsx } from '../index.js'

describe('XLSX sparse ranges', () => {
  it('imports actual cells without scanning every coordinate in a broad sparse ref', () => {
    const imported = importXlsx(buildBroadSparseWorkbookBytes(), 'broad-sparse.xlsx')
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells).toEqual([{ address: 'XFD512', formula: '40+2' }])
    expect(imported.preview.sheets[0]).toMatchObject({
      rowCount: 512,
      columnCount: 16_384,
      nonEmptyCellCount: 1,
    })
  }, 15_000)

  it('skips styled blank XML cells when importing a broad styled template range', () => {
    const sparseBytes = buildStyledBlankWorkbookBytes({ includeBlankCells: false })
    const denseBytes = buildStyledBlankWorkbookBytes({ includeBlankCells: true })
    const sparseMs = measureImport(sparseBytes, 'styled-sparse-control.xlsx').durationMs
    const denseMeasurements = [
      measureImport(denseBytes, 'styled-blank-template.xlsx'),
      measureImport(denseBytes, 'styled-blank-template.xlsx'),
    ]
    const { imported, durationMs: denseMs } = denseMeasurements.reduce((best, current) =>
      current.durationMs < best.durationMs ? current : best,
    )
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells).toEqual([{ address: 'A1', value: 123 }])
    expect(imported.preview.sheets[0]).toMatchObject({
      rowCount: styledBlankRowCount,
      columnCount: styledBlankColumnCount,
      nonEmptyCellCount: 1,
    })
    expect(sheet?.metadata?.styleRanges).toHaveLength(1)
    expect(sheet?.metadata?.styleRanges?.[0]?.range).toEqual({
      sheetName: 'StyledBlanks',
      startAddress: 'A1',
      endAddress: 'A1',
    })
    expect(imported.snapshot.workbook.metadata?.styles?.[0]).toMatchObject({
      fill: { backgroundColor: '#ffcc00' },
    })
    const tolerance = readBenchmarkTolerance()
    expect(denseMs).toBeLessThan(Math.max(1_000 * tolerance, sparseMs * 12 * tolerance))
  }, 15_000)
})

function readBenchmarkTolerance(): number {
  const raw = process.env.BILIG_BENCH_TOLERANCE
  if (!raw) {
    return 1
  }
  const tolerance = Number(raw)
  return Number.isFinite(tolerance) && tolerance > 0 ? tolerance : 1
}

function measureImport(bytes: Uint8Array, fileName: string): { imported: ReturnType<typeof importXlsx>; durationMs: number } {
  const start = performance.now()
  const imported = importXlsx(bytes, fileName)
  return {
    imported,
    durationMs: performance.now() - start,
  }
}

function buildBroadSparseWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet: XLSX.WorkSheet = {
    XFD512: { t: 'n', f: '40+2', v: 42 },
    '!ref': 'A1:XFD512',
  }
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sparse')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

const styledBlankRowCount = 20_000
const styledBlankColumnCount = 26

function buildStyledBlankWorkbookBytes(options: { includeBlankCells: boolean }): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[123]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'StyledBlanks')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(buildStyledBlankWorksheetXml(options.includeBlankCells))
  zip['xl/styles.xml'] = strToU8(styledBlankWorkbookStylesXml)
  return zipSync(zip)
}

function buildStyledBlankWorksheetXml(includeBlankCells: boolean): string {
  const rows: string[] = []
  for (let row = 1; row <= styledBlankRowCount; row += 1) {
    const cells: string[] = []
    const columnCount = includeBlankCells ? styledBlankColumnCount : row === 1 ? 1 : 0
    for (let column = 0; column < columnCount; column += 1) {
      const address = `${encodeColumnName(column)}${String(row)}`
      cells.push(address === 'A1' ? '<c r="A1" s="1"><v>123</v></c>' : `<c r="${address}" s="1"/>`)
    }
    if (cells.length > 0) {
      rows.push(`<row r="${String(row)}">${cells.join('')}</row>`)
    }
  }
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<dimension ref="A1:${encodeColumnName(styledBlankColumnCount - 1)}${String(styledBlankRowCount)}"/>`,
    `<sheetData>${rows.join('')}</sheetData>`,
    '</worksheet>',
  ].join('')
}

function encodeColumnName(index: number): string {
  let value = index + 1
  let output = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    value = Math.floor((value - 1) / 26)
  }
  return output
}

const styledBlankWorkbookStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
  '<fills count="3">',
  '<fill><patternFill patternType="none"/></fill>',
  '<fill><patternFill patternType="gray125"/></fill>',
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/><bgColor indexed="64"/></patternFill></fill>',
  '</fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2">',
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>',
  '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>',
  '</cellXfs>',
  '</styleSheet>',
].join('')
