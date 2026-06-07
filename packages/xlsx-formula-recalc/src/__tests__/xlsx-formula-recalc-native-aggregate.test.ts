import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { recalculateXlsxFileToFile } from '../index.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('xlsx-formula-recalc native aggregates', () => {
  it('loads public-corpus style dependency rows and recalculates SUM ranges through the native kernel', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-sum-range-'))
    try {
      const sourcePath = join(tempDir, 'public-sum-range.xlsx')
      const outputPath = join(tempDir, 'public-sum-range.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicSumRangeWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Table 4!A1'],
      })

      expect(readNumber(result.reads['Table 4!A1'])).toBe(425)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A1"><f>SUM(B5:B8)</f><v>425</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('loads same-sheet scalar dependency rows for public-corpus formulas', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-scalar-dependency-'))
    try {
      const sourcePath = join(tempDir, 'public-scalar-dependency.xlsx')
      const outputPath = join(tempDir, 'public-scalar-dependency.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicScalarDependencyWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!C7'],
      })

      expect(readNumber(result.reads['Budget!C7'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(3)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="C7"><f>B6+B8*2</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('loads cross-sheet scalar dependency rows for public-corpus formulas', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-cross-sheet-scalar-'))
    try {
      const sourcePath = join(tempDir, 'public-cross-sheet-scalar.xlsx')
      const outputPath = join(tempDir, 'public-cross-sheet-scalar.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicCrossSheetScalarDependencyWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!C7'],
      })

      expect(readNumber(result.reads['Budget!C7'])).toBe(160)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(3)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain("<c r=\"C7\"><f>'Data Sheet'!B2+'Data Sheet'!C7*2</f><v>160</v></c>")
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

function readNumber(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number') {
    return value.value
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}

function buildPublicSumRangeWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Table 4" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B8"/>
  <sheetData>
    <row r="1"><c r="A1"><f>SUM(B5:B8)</f><v>0</v></c></row>
    <row r="5"><c r="B5"><v>100</v></c></row>
    <row r="6"><c r="B6"><v>125</v></c></row>
    <row r="7"><c r="B7"><v>75</v></c></row>
    <row r="8"><c r="B8"><v>125</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicScalarDependencyWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B6:C8"/>
  <sheetData>
    <row r="6"><c r="B6"><v>50</v></c></row>
    <row r="7"><c r="C7"><f>B6+B8*2</f><v>0</v></c></row>
    <row r="8"><c r="B8"><v>100</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicCrossSheetScalarDependencyWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Data Sheet" sheetId="1" r:id="rId1"/>
    <sheet name="Budget" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:C7"/>
  <sheetData>
    <row r="2"><c r="B2"><v>40</v></c></row>
    <row r="7"><c r="C7"><v>60</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="C7"/>
  <sheetData>
    <row r="7"><c r="C7"><f>'Data Sheet'!B2+'Data Sheet'!C7*2</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}
