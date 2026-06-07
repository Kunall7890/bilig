import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { recalculateXlsxFileToFile } from '../index.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('xlsx-formula-recalc native aggregates', () => {
  it('hydrates high-index shared string targets without replaying the shared-string buffer', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-shared-string-buffer-'))
    try {
      const sourcePath = join(tempDir, 'shared-string-buffer.xlsx')
      const outputPath = join(tempDir, 'shared-string-buffer.recalculated.xlsx')
      writeFileSync(sourcePath, buildHighIndexSharedStringWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A1'],
      })

      expect(result.reads['Data!A1']).toMatchObject({ value: 'target-shared-string' })
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(0)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

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

  it('evaluates SUM and COUNTA formulas over scanned public-corpus ranges', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-sum-counta-range-'))
    try {
      const sourcePath = join(tempDir, 'public-sum-counta-range.xlsx')
      const outputPath = join(tempDir, 'public-sum-counta-range.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicSumCountaRangeWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2', 'Data!A6'],
      })

      expect(readNumber(result.reads['Data!A2'])).toBe(12)
      expect(readNumber(result.reads['Data!A6'])).toBe(2)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(2)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(2)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>SUM(B2:B5)</f><v>12</v></c>')
      expect(sheetXml).toContain('<c r="A6"><f>IF(COUNTA(B6:D6)=0,"",COUNTA(C6:E6))</f><v>2</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates ROUND and AVERAGE formulas from public financial forecasts', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-round-average-'))
    try {
      const sourcePath = join(tempDir, 'public-round-average.xlsx')
      const outputPath = join(tempDir, 'public-round-average.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicRoundAverageWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2', 'Data!A3', 'Data!A4', 'Data!A5', 'Data!A6'],
      })

      expect(readNumber(result.reads['Data!A2'])).toBe(3.1)
      expect(readNumber(result.reads['Data!A3'])).toBe(-2)
      expect(readNumber(result.reads['Data!A4'])).toBe(1200)
      expect(readBoolean(result.reads['Data!A5'])).toBe(true)
      expect(readNumber(result.reads['Data!A6'])).toBe(0)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(9)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(5)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>ROUND(AVERAGE(B2:B6),1)</f><v>3.1</v></c>')
      expect(sheetXml).toContain('<c r="A3"><f>ROUND(B7,0)</f><v>-2</v></c>')
      expect(sheetXml).toContain('<c r="A4"><f>ROUND(B8,-2)</f><v>1200</v></c>')
      expect(sheetXml).toContain('<c r="A5" t="b"><f>ISERROR(B9/B10)</f><v>1</v></c>')
      expect(sheetXml).toContain('<c r="A6"><f>IF(ISERROR(B9/B10),0,B9/B10)</f><v>0</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates exact INDEX MATCH header lookups from public forecasts', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-index-match-'))
    try {
      const sourcePath = join(tempDir, 'public-index-match.xlsx')
      const outputPath = join(tempDir, 'public-index-match.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicIndexMatchWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2'],
      })

      expect(readNumber(result.reads['Data!A2'])).toBe(30)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>INDEX(B5:B8,MATCH(B6,B5:B8,0)+1,1)</f><v>30</v></c>')
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

  it('evaluates exact VLOOKUP formulas over scanned public-corpus table rows', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-vlookup-'))
    try {
      const sourcePath = join(tempDir, 'public-vlookup.xlsx')
      const outputPath = join(tempDir, 'public-vlookup.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicVlookupWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!D3'],
      })

      expect(readNumber(result.reads['Budget!D3'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(4)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="D3"><f>VLOOKUP(B2,\'Lookup Table\'!$B$9:$D$10,2,FALSE)</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates omitted approximate VLOOKUP formulas over scanned table rows', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-approximate-vlookup-'))
    try {
      const sourcePath = join(tempDir, 'public-approximate-vlookup.xlsx')
      const outputPath = join(tempDir, 'public-approximate-vlookup.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicApproximateVlookupWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!D3'],
      })

      expect(readNumber(result.reads['Budget!D3'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="D3"><f>VLOOKUP(B2,\'Lookup Table\'!$B$9:$D$11,2)</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates VLOOKUP formulas with INDIRECT scalar cell-range table rows', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-indirect-vlookup-'))
    try {
      const sourcePath = join(tempDir, 'public-indirect-vlookup.xlsx')
      const outputPath = join(tempDir, 'public-indirect-vlookup.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicIndirectVlookupWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!D3'],
      })

      expect(readNumber(result.reads['Budget!D3'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(6)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="D3"><f>VLOOKUP(B2,INDIRECT("\'"&amp;C$4&amp;"\'!"&amp;"$B$9:$D$11"),2)</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('fails closed with diagnostics for external workbook references', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-external-reference-'))
    try {
      const sourcePath = join(tempDir, 'external-reference.xlsx')
      const outputPath = join(tempDir, 'external-reference.recalculated.xlsx')
      writeFileSync(sourcePath, buildExternalReferenceWorkbook())

      await expect(
        recalculateXlsxFileToFile(sourcePath, {
          outputPath,
          engine: 'streaming-native',
          reads: ['Budget!B2'],
        }),
      ).rejects.toMatchObject({
        diagnostics: expect.objectContaining({
          engineMode: 'streaming-native',
          unsupportedReason: 'external workbook references are not supported',
        }),
      })
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

function readBoolean(value: unknown): boolean {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'boolean') {
    return value.value
  }
  throw new Error(`Expected boolean cell value, received ${JSON.stringify(value)}`)
}

function buildHighIndexSharedStringWorkbook(): Uint8Array {
  const sharedStrings = Array.from({ length: 74 }, (_value, index) => `<si><t>unused-${String(index)}</t></si>`)
  sharedStrings.push('<si><t>target-shared-string</t></si>')
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
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
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="75" uniqueCount="75">${sharedStrings.join('')}</sst>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>74</v></c></row>
  </sheetData>
</worksheet>`),
  })
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

function buildPublicSumCountaRangeWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
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
  <dimension ref="A2:E6"/>
  <sheetData>
    <row r="2"><c r="A2"><f>SUM(B2:B5)</f><v>0</v></c><c r="B2"><v>5</v></c></row>
    <row r="3"><c r="B3" t="inlineStr"><is><t>ignored</t></is></c></row>
    <row r="4"/>
    <row r="5"><c r="B5"><v>7</v></c></row>
    <row r="6"><c r="A6"><f>IF(COUNTA(B6:D6)=0,"",COUNTA(C6:E6))</f><v>0</v></c><c r="C6"><v>1</v></c><c r="D6" t="inlineStr"><is><t>x</t></is></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicRoundAverageWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
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
  <dimension ref="A2:B10"/>
  <sheetData>
    <row r="2"><c r="A2"><f>ROUND(AVERAGE(B2:B6),1)</f><v>0</v></c><c r="B2"><v>2</v></c></row>
    <row r="3"><c r="A3"><f>ROUND(B7,0)</f><v>0</v></c><c r="B3"><v>3</v></c></row>
    <row r="4"><c r="A4"><f>ROUND(B8,-2)</f><v>0</v></c><c r="B4" t="inlineStr"><is><t>ignored</t></is></c></row>
    <row r="5"><c r="A5"><f>ISERROR(B9/B10)</f><v>0</v></c></row>
    <row r="6"><c r="A6"><f>IF(ISERROR(B9/B10),0,B9/B10)</f><v>99</v></c><c r="B6"><v>4.4</v></c></row>
    <row r="7"><c r="B7"><v>-1.5</v></c></row>
    <row r="8"><c r="B8"><v>1234.56</v></c></row>
    <row r="9"><c r="B9"><v>10</v></c></row>
    <row r="10"><c r="B10"><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicIndexMatchWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
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
  <dimension ref="A2:B8"/>
  <sheetData>
    <row r="2"><c r="A2"><f>INDEX(B5:B8,MATCH(B6,B5:B8,0)+1,1)</f><v>0</v></c></row>
    <row r="5"><c r="B5"><v>10</v></c></row>
    <row r="6"><c r="B6"><v>20</v></c></row>
    <row r="7"><c r="B7"><v>30</v></c></row>
    <row r="8"><c r="B8"><v>40</v></c></row>
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

function buildPublicVlookupWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Lookup Table" sheetId="1" r:id="rId1"/>
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
  <dimension ref="B9:D10"/>
  <sheetData>
    <row r="9"><c r="B9"><v>100</v></c><c r="C9"><v>150</v></c><c r="D9"><v>175</v></c></row>
    <row r="10"><c r="B10"><v>200</v></c><c r="C10"><v>250</v></c><c r="D10"><v>275</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:D3"/>
  <sheetData>
    <row r="2"><c r="B2"><v>200</v></c></row>
    <row r="3"><c r="D3"><f>VLOOKUP(B2,'Lookup Table'!$B$9:$D$10,2,FALSE)</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicApproximateVlookupWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Lookup Table" sheetId="1" r:id="rId1"/>
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
  <dimension ref="B9:D11"/>
  <sheetData>
    <row r="9"><c r="B9"><v>100</v></c><c r="C9"><v>150</v></c><c r="D9"><v>175</v></c></row>
    <row r="10"><c r="B10"><v>200</v></c><c r="C10"><v>250</v></c><c r="D10"><v>275</v></c></row>
    <row r="11"><c r="B11"><v>300</v></c><c r="C11"><v>350</v></c><c r="D11"><v>375</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:D3"/>
  <sheetData>
    <row r="2"><c r="B2"><v>250</v></c></row>
    <row r="3"><c r="D3"><f>VLOOKUP(B2,'Lookup Table'!$B$9:$D$11,2)</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicIndirectVlookupWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Lookup Table" sheetId="1" r:id="rId1"/>
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
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Lookup Table</t></si>
</sst>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B9:D11"/>
  <sheetData>
    <row r="9"><c r="B9"><v>100</v></c><c r="C9"><v>150</v></c><c r="D9"><v>175</v></c></row>
    <row r="10"><c r="B10"><v>200</v></c><c r="C10"><v>250</v></c><c r="D10"><v>275</v></c></row>
    <row r="11"><c r="B11"><v>300</v></c><c r="C11"><v>350</v></c><c r="D11"><v>375</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:D4"/>
  <sheetData>
    <row r="2"><c r="B2"><v>250</v></c></row>
    <row r="3"><c r="D3"><f>VLOOKUP(B2,INDIRECT("'"&amp;C$4&amp;"'!"&amp;"$B$9:$D$11"),2)</f><v>0</v></c></row>
    <row r="4"><c r="C4" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildExternalReferenceWorkbook(): Uint8Array {
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
  <dimension ref="B2"/>
  <sheetData>
    <row r="2"><c r="B2"><f>[1]QTR_detailed!$B130/10^6</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}
