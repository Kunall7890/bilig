import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import {
  decodeCellAddress,
  decodeCellRange,
  decodeColumnAddress,
  encodeCellAddress,
  encodeCellRange,
  normalizeCellAddress,
  readXlsxZipEntries,
  readXmlAttribute,
  worksheetCellElementPattern,
  writeSimpleXlsxWorkbook,
} from '../index.js'

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const textDecoder = new TextDecoder()

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readPackageJson(): Record<string, unknown> {
  const packageJson: unknown = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'))
  if (!isObjectRecord(packageJson)) {
    throw new TypeError('Expected package.json to parse to an object')
  }
  return packageJson
}

function sourceFiles(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name)
    if (entry.isDirectory()) {
      return entry.name === '__tests__' ? [] : sourceFiles(path)
    }
    return entry.isFile() && entry.name.endsWith('.ts') ? [path] : []
  })
}

describe('@bilig/xlsx package boundary', () => {
  it('does not depend on SheetJS or the xlsx CDN tarball', () => {
    const packageJson = readPackageJson()
    expect(packageJson.name).toBe('@bilig/xlsx')
    expect(isObjectRecord(packageJson.dependencies)).toBe(true)
    if (isObjectRecord(packageJson.dependencies)) {
      expect(packageJson.dependencies).not.toHaveProperty('xlsx')
      expect(packageJson.dependencies).not.toHaveProperty('xlsx-js-style')
    }

    const manifestSource = readFileSync(join(packageDir, 'package.json'), 'utf8')
    expect(manifestSource).not.toContain('cdn.sheetjs.com')
    expect(manifestSource).not.toMatch(/"xlsx"\s*:/u)
  })

  it('keeps @bilig/xlsx source free of SheetJS imports', () => {
    for (const sourceFile of sourceFiles(join(packageDir, 'src'))) {
      const source = readFileSync(sourceFile, 'utf8')
      expect(source, sourceFile).not.toMatch(/from\s+["']xlsx["']/u)
      expect(source, sourceFile).not.toMatch(/require\(["']xlsx["']\)/u)
      expect(source, sourceFile).not.toContain('cdn.sheetjs.com')
    }
  })

  it('normalizes XLSX addresses and ranges without SheetJS', () => {
    expect(decodeCellAddress('$AA$42')).toEqual({ r: 41, c: 26 })
    expect(decodeColumnAddress('$XFD')).toBe(16383)
    expect(encodeCellAddress({ r: 0, c: 701 })).toBe('ZZ1')
    expect(normalizeCellAddress('$b$7')).toBe('B7')
    expect(decodeCellRange('C3:A1')).toEqual({ s: { r: 0, c: 0 }, e: { r: 2, c: 2 } })
    expect(encodeCellRange({ s: { r: 0, c: 0 }, e: { r: 2, c: 2 } })).toBe('A1:C3')
  })

  it('exports XML helpers for @bilig/xlsx package readers without SheetJS', () => {
    expect(readXmlAttribute('<sheet name="A &amp; B" r:id="rId1"/>', 'name')).toBe('A & B')
    const cells = [
      ...'<sheetData><row><c r="A1"><v>1</v></c><c r="B2" t="str"><v>x</v></c></row></sheetData>'.matchAll(worksheetCellElementPattern),
    ].map((match) => match[0])

    expect(cells).toHaveLength(2)
  })

  it('writes deterministic simple workbook bytes for generated fixtures', () => {
    const workbook = {
      sheets: [
        {
          name: 'Sheet 1',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'segment' },
            { address: 'B1', row: 0, col: 1, formula: 'SUM(1,2)', value: 3 },
          ],
        },
      ],
    }

    expect(writeSimpleXlsxWorkbook(workbook)).toEqual(writeSimpleXlsxWorkbook(workbook))
  })

  it('writes compatibility theme and escaped formula XML in simple workbooks', () => {
    const workbook = {
      sheets: [
        {
          name: 'Sheet 1',
          cells: [{ address: 'A1', row: 0, col: 0, formula: 'TEXTJOIN("-",TRUE,B1:B2)' }],
        },
      ],
    }
    const zip = readXlsxZipEntries(writeSimpleXlsxWorkbook(workbook))

    expect(zip['xl/theme/theme1.xml']).toBeDefined()
    expect(textDecoder.decode(zip['xl/_rels/workbook.xml.rels'])).toContain('relationships/theme')
    expect(textDecoder.decode(zip['xl/worksheets/sheet1.xml'])).toContain('<f>TEXTJOIN(&quot;-&quot;,TRUE,B1:B2)</f>')
  })

  it('writes border styles with @bilig/xlsx simple workbooks', () => {
    const workbook = {
      styles: [
        {
          id: 'bordered-total',
          borders: {
            top: { style: 'double' as const, weight: 'medium' as const, color: '#AA0000' },
            bottom: { style: 'solid' as const, weight: 'thin' as const, color: '#000000' },
          },
        },
      ],
      sheets: [
        {
          name: 'Report',
          cells: [{ address: 'B7', row: 6, col: 1, value: 42, styleId: 'bordered-total' }],
        },
      ],
    }
    const zip = readXlsxZipEntries(writeSimpleXlsxWorkbook(workbook))
    const stylesXml = textDecoder.decode(zip['xl/styles.xml'])
    const sheetXml = textDecoder.decode(zip['xl/worksheets/sheet1.xml'])

    expect(stylesXml).toContain('<borders count="2">')
    expect(stylesXml).toContain('<top style="double"><color rgb="FFAA0000"/></top>')
    expect(stylesXml).toContain('<bottom style="thin"><color rgb="FF000000"/></bottom>')
    expect(stylesXml).toContain('borderId="1"')
    expect(stylesXml).toContain('applyBorder="1"')
    expect(sheetXml).toContain('<c r="B7" s="1"><v>42</v></c>')
  })

  it('preserves raw style XML and direct style indexes in simple workbooks', () => {
    const workbook = {
      stylesXml: minimalRawStylesXml,
      sheets: [
        {
          name: 'Sparse',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Header' },
            { address: 'CF65000', row: 64_999, col: 83, styleIndex: 1 },
          ],
        },
      ],
    }
    const zip = readXlsxZipEntries(writeSimpleXlsxWorkbook(workbook))
    const stylesXml = textDecoder.decode(zip['xl/styles.xml'])
    const sheetXml = textDecoder.decode(zip['xl/worksheets/sheet1.xml'])

    expect(stylesXml).toBe(minimalRawStylesXml)
    expect(sheetXml).toContain('<dimension ref="A1:CF65000"/>')
    expect(sheetXml).toContain('<c r="CF65000" s="1"/>')
  })

  it('writes rich shared strings and inline strings with @bilig/xlsx simple workbooks', () => {
    const sharedStringsXml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">',
      '<si><r><rPr><b/></rPr><t>Important</t></r></si>',
      '</sst>',
    ].join('')
    const inlineStringXml = '<is><r><rPr><i/></rPr><t>Inline rich</t></r></is>'
    const workbook = {
      sharedStringsXml,
      sheets: [
        {
          name: 'Labels',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Important', sharedStringIndex: 0 },
            { address: 'B1', row: 0, col: 1, value: 'Inline rich', inlineStringXml },
          ],
        },
      ],
    }
    const zip = readXlsxZipEntries(writeSimpleXlsxWorkbook(workbook))
    const contentTypesXml = textDecoder.decode(zip['[Content_Types].xml'])
    const relsXml = textDecoder.decode(zip['xl/_rels/workbook.xml.rels'])
    const sheetXml = textDecoder.decode(zip['xl/worksheets/sheet1.xml'])

    expect(textDecoder.decode(zip['xl/sharedStrings.xml'])).toBe(sharedStringsXml)
    expect(contentTypesXml).toContain('/xl/sharedStrings.xml')
    expect(relsXml).toContain('relationships/sharedStrings')
    expect(sheetXml).toContain('<c r="A1" t="s"><v>0</v></c>')
    expect(sheetXml).toContain(`<c r="B1" t="inlineStr">${inlineStringXml}</c>`)
  })

  it('only applies direct style indexes when raw styles XML is supplied', () => {
    const workbook = {
      sheets: [
        {
          name: 'Sheet 1',
          cells: [{ address: 'B2', row: 1, col: 1, value: 7, styleIndex: 42 }],
        },
      ],
    }
    const zip = readXlsxZipEntries(writeSimpleXlsxWorkbook(workbook))
    const sheetXml = textDecoder.decode(zip['xl/worksheets/sheet1.xml'])

    expect(sheetXml).toContain('<c r="B2"><v>7</v></c>')
    expect(sheetXml).not.toContain('s="42"')
  })
})

const minimalRawStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')
