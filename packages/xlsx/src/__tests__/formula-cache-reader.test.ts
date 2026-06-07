import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { defaultXlsxFormulaCacheInspectionLimit, readXlsxFormulaCacheCellsFromFile, zipSourcePreservingEntries } from '../index.js'

const textEncoder = new TextEncoder()

describe('@bilig/xlsx formula cache reader', () => {
  it('streams formula cache cells from a file-backed XLSX source', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-formula-cache-'))
    try {
      const inputPath = join(tempDir, 'shared-formulas.xlsx')
      writeFileSync(inputPath, formulaCacheWorkbookBytes())

      const scan = readXlsxFormulaCacheCellsFromFile(inputPath, { inspectLimit: 2 })

      expect(scan.sheetNames).toEqual(['Revenue & Ops'])
      expect(scan.formulaCellCount).toBe(3)
      expect(scan.cells).toEqual([
        {
          target: "'Revenue & Ops'!B1",
          formula: '=A1&" units"',
          cachedValue: 'old units',
        },
        {
          target: "'Revenue & Ops'!C1",
          formula: '=A1*10',
          cachedValue: 20,
        },
      ])
      expect(scan.inputBytes).toBeGreaterThan(0)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('translates shared formulas while counting cells beyond the inspect limit', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-shared-formula-cache-'))
    try {
      const inputPath = join(tempDir, 'shared-formulas.xlsx')
      writeFileSync(inputPath, formulaCacheWorkbookBytes())

      const scan = readXlsxFormulaCacheCellsFromFile(inputPath)

      expect(scan.formulaCellCount).toBe(3)
      expect(scan.cells.map((cell) => [cell.target, cell.formula, cell.cachedValue])).toEqual([
        ["'Revenue & Ops'!B1", '=A1&" units"', 'old units'],
        ["'Revenue & Ops'!C1", '=A1*10', 20],
        ["'Revenue & Ops'!C2", '=A2*10', 30],
      ])
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('bounds default file-backed formula cache collection', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-default-formula-cache-limit-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, manyFormulaCacheWorkbookBytes(defaultXlsxFormulaCacheInspectionLimit + 1))

      const defaultScan = readXlsxFormulaCacheCellsFromFile(inputPath)
      const fullScan = readXlsxFormulaCacheCellsFromFile(inputPath, { inspectLimit: 'all' })

      expect(defaultScan.formulaCellCount).toBe(defaultXlsxFormulaCacheInspectionLimit + 1)
      expect(defaultScan.cells).toHaveLength(defaultXlsxFormulaCacheInspectionLimit)
      expect(fullScan.cells).toHaveLength(defaultXlsxFormulaCacheInspectionLimit + 1)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

function formulaCacheWorkbookBytes(): Uint8Array {
  return zipSourcePreservingEntries({
    '[Content_Types].xml': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`),
    '_rels/.rels': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/workbook.xml': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Revenue &amp; Ops" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>old units</t></si>
</sst>`),
    'xl/worksheets/sheet1.xml': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C2"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>2</v></c>
      <c r="B1" t="s"><f>A1&amp;" units"</f><v>0</v></c>
      <c r="C1"><f t="shared" si="1">A1*10</f><v>20</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="C2"><f t="shared" si="1"/><v>30</v></c>
    </row>
  </sheetData>
</worksheet>`),
  })
}

function bytes(value: string): Uint8Array {
  return textEncoder.encode(value)
}

function manyFormulaCacheWorkbookBytes(formulaCount: number): Uint8Array {
  const rows = Array.from({ length: formulaCount }, (_value, index) => {
    const row = index + 1
    return `<row r="${row.toString()}"><c r="A${row.toString()}"><v>${row.toString()}</v></c><c r="B${row.toString()}"><f>A${row.toString()}*2</f><v>${(row * 2).toString()}</v></c></row>`
  }).join('')
  return zipSourcePreservingEntries({
    '[Content_Types].xml': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/workbook.xml': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': bytes(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B${formulaCount.toString()}"/>
  <sheetData>${rows}</sheetData>
</worksheet>`),
  })
}
