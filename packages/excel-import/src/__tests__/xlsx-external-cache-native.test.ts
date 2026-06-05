import { strToU8, zipSync } from 'fflate'
import { describe, expect, it, vi } from 'vitest'

import type { WorkbookExternalWorkbookReferenceSnapshot } from '@bilig/protocol'
import type { ImportedExternalLinkCacheUsage } from '../xlsx-external-cache.js'

vi.mock('xlsx', () => ({
  read() {
    throw new Error('SheetJS read should not be used for external companion cache hydration')
  },
}))

describe('@bilig/xlsx external companion cache hydration', () => {
  it('reads referenced companion workbook cells without SheetJS', async () => {
    const { readExternalWorkbookCacheFromInput } = await import('../xlsx-external-cache.js')
    const reference: WorkbookExternalWorkbookReferenceSnapshot = {
      bookIndex: 1,
      packagePath: 'xl/externalLinks/externalLink1.xml',
      target: 'file:///tmp/rates.xlsx',
      targetMode: 'External',
      workbookName: 'rates.xlsx',
      sheetNames: ['Rates'],
    }
    const usage: ImportedExternalLinkCacheUsage = new Map([[1, new Map([['rates', new Set(['A1', 'B2', 'C3', 'D4', 'E5', 'F6', 'G7'])]])]])

    const sheets = readExternalWorkbookCacheFromInput({ fileName: 'rates.xlsx', bytes: buildExternalCompanionWorkbook() }, reference, usage)
    const cells = sheets.get('rates')?.cells

    expect(cells?.get('A1')).toEqual({ kind: 'string', value: 'SKU & Name' })
    expect(cells?.get('B2')).toEqual({ kind: 'number', value: 42.5 })
    expect(cells?.get('C3')).toEqual({ kind: 'boolean', value: true })
    expect(cells?.get('D4')).toEqual({ kind: 'string', value: 'Formula cache text' })
    expect(cells?.get('E5')).toEqual({ kind: 'error', value: '#N/A' })
    expect(cells?.get('F6')).toEqual({ kind: 'blank' })
    expect(cells?.get('G7')).toEqual({ kind: 'string', value: 'Inline Value' })
    expect(cells?.has('H8')).toBe(false)
  })
})

function buildExternalCompanionWorkbook(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
        '<Override PartName="/xl/worksheets/custom-rates.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '</Types>',
      ].join(''),
    ),
    '_rels/.rels': strToU8(
      [
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
        '</Relationships>',
      ].join(''),
    ),
    'xl/workbook.xml': strToU8(
      [
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Rates" sheetId="1" r:id="rIdRates"/></sheets>',
        '</workbook>',
      ].join(''),
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      [
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdRates" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/custom-rates.xml"/>',
        '<Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
        '</Relationships>',
      ].join(''),
    ),
    'xl/sharedStrings.xml': strToU8(
      [
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">',
        '<si><t>SKU &amp; Name</t></si>',
        '</sst>',
      ].join(''),
    ),
    'xl/worksheets/custom-rates.xml': strToU8(
      [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData>',
        '<row r="1"><c r="A1" t="s"><v>0</v></c></row>',
        '<row r="2"><c r="B2"><v>42.5</v></c></row>',
        '<row r="3"><c r="C3" t="b"><v>1</v></c></row>',
        '<row r="4"><c r="D4" t="str"><v>Formula cache text</v></c></row>',
        '<row r="5"><c r="E5" t="e"><v>#N/A</v></c></row>',
        '<row r="7"><c r="G7" t="inlineStr"><is><t>Inline Value</t></is></c><c r="H8"><v>999</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    ),
  })
}
