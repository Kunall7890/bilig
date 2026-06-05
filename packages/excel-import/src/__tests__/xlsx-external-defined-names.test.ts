import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'

import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '@bilig/core'
import { externalWorkbookReferencesWarning, importXlsx } from '../index.js'

describe('xlsx external defined names', () => {
  it('resolves external workbook defined names from saved XLSX external-link caches', async () => {
    const imported = importXlsx(buildExternalDefinedNameCacheWorkbook(), 'external-defined-name-cache.xlsx')
    const definedNames = imported.snapshot.workbook.metadata?.definedNames ?? []

    expect(definedNames).toContainEqual({
      name: 'ReportDate',
      value: { kind: 'scalar', value: 'January 30, 2024' },
    })
    expect(imported.warnings).toEqual([externalWorkbookReferencesWarning])

    const engine = new SpreadsheetEngine({ workbookName: 'external-defined-name-cache-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Report', 'A1')).toMatchObject({
      tag: ValueTag.String,
      value: 'Includes forecasts finalized on or before January 30, 2024.',
    })
  })
})

function buildExternalDefinedNameCacheWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Report',
          cells: [
            {
              address: 'A1',
              row: 0,
              col: 0,
              formula: '="Includes forecasts finalized on or before "&ReportDate&"."',
              value: 'Includes forecasts finalized on or before January 30, 2024.',
            },
          ],
        },
      ],
      definedNames: [{ name: 'ReportDate', formula: "'[1]Contacts, Cutoffs, and Data'!$B$1" }],
    }),
  )
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'])
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="Contacts, Cutoffs, and Data"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="1"><cell r="B1" t="str"><v>January 30, 2024</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/budget-source.xlsx" TargetMode="External"/>' +
      '</Relationships>',
  )
  return zipSync(zip)
}
