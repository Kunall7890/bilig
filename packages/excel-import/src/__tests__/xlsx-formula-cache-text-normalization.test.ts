import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '@bilig/core'

import { importXlsx } from '../index.js'
import { readWorksheetFormulaCells } from '../xlsx-formulas.js'

describe('xlsx formula cache text normalization', () => {
  it('normalizes cached string formula line endings while reading worksheet formulas', () => {
    const cells = readWorksheetFormulaCells(
      [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1">',
        '<c r="A1" t="str"><f>[1]External!A1</f><v>Line 1\r\nLine 2\rLine 3</v></c>',
        '</row></sheetData>',
        '</worksheet>',
      ].join(''),
    )

    expect(cells[0]?.cachedValue).toBe('Line 1\nLine 2\nLine 3')
  })

  it('keeps imported unresolved external formula caches in oracle-normalized text form after recalculation', async () => {
    const imported = importXlsx(buildExternalFormulaCacheWorkbook(), 'external-string-cache.xlsx')
    const formulaCell = imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')

    expect(formulaCell).toMatchObject({
      formula: '[1]External!A1',
      value: 'Line 1\nLine 2\nLine 3',
    })

    const engine = new SpreadsheetEngine({ workbookName: 'external-string-cache' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Report', 'A1')).toEqual({
      tag: ValueTag.String,
      value: 'Line 1\nLine 2\nLine 3',
      stringId: expect.any(Number),
    })
  })
})

function buildExternalFormulaCacheWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Report',
          cells: [{ address: 'A1', row: 0, col: 0, formula: '[1]External!A1', value: 'Line 1\r\nLine 2\rLine 3' }],
        },
      ],
    }),
  )
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1"><sheetNames><sheetName val="External"/></sheetNames></externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink1.xml.rels'] = strToU8(
    [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/external.xlsx" TargetMode="External"/>',
      '</Relationships>',
    ].join(''),
  )
  return zipSync(zip)
}
