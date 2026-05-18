import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { importXlsx } from '../index.js'

describe('MS-OI29500 external data provenance import', () => {
  it('imports connections, external workbook links, DDE, and OLE provenance without refresh execution', () => {
    const imported = importXlsx(buildExternalDataWorkbookBytes(), 'ms-oi29500-external-data.xlsx')
    const externalConnections = imported.snapshot.workbook.metadata?.externalConnections

    expect(externalConnections).toMatchObject({
      refreshExecution: 'disabled',
      connections: [
        expect.objectContaining({
          id: 1,
          name: 'Sales Query',
          sourceKind: 'database',
          command: 'SELECT * FROM Sales',
          refreshOnLoad: false,
          clause: '18.13',
        }),
      ],
      externalLinks: expect.arrayContaining([
        expect.objectContaining({
          kind: 'external-workbook',
          bookIndex: 1,
          target: 'file:///tmp/source.xlsx',
          sheetNames: ['Source'],
          clause: '18.14',
        }),
        expect.objectContaining({
          kind: 'dde',
          service: 'cmd',
          topic: 'topic',
          itemNames: ['A1'],
          refreshExecution: 'disabled',
          clause: '18.14',
        }),
        expect.objectContaining({
          kind: 'ole',
          progId: 'Word.Document',
          relationshipId: 'rId2',
          refreshExecution: 'disabled',
          clause: '18.14',
        }),
      ]),
    })
  })
})

function buildExternalDataWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([['Local'], [1]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  zip['xl/workbook.xml'] = strToU8(
    sourceWorkbookXml.replace(
      '</workbook>',
      '<externalReferences><externalReference r:id="rIdExternal1"/></externalReferences></workbook>',
    ),
  )
  const sourceRelsXml = strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array())
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    sourceRelsXml.replace(
      '</Relationships>',
      '<Relationship Id="rIdExternal1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>',
    ),
  )
  zip['xl/connections.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<connection id="1" name="Sales Query" type="5" refreshedVersion="8" refreshOnLoad="0">',
      '<dbPr connection="Provider=SQLOLEDB;Data Source=example" command="SELECT * FROM Sales" commandType="2"/>',
      '</connection>',
      '</connections>',
    ].join(''),
  )
  zip['xl/externalLinks/externalLink1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<externalBook r:id="rId1"><sheetNames><sheetName val="Source"/></sheetNames></externalBook>',
      '<ddeLink ddeService="cmd" ddeTopic="topic"><ddeItems count="1"><ddeItem name="A1"/></ddeItems></ddeLink>',
      '<oleLink progId="Word.Document" r:id="rId2"><oleItems count="1"><oleItem name="Document"/></oleItems></oleLink>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink1.xml.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/source.xlsx" TargetMode="External"/>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///tmp/document.docx" TargetMode="External"/>',
      '</Relationships>',
    ].join(''),
  )

  return zipSync(zip)
}
