import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle } from '@bilig/excel-fixtures'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const connectionsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const externalLinkPathRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'
const connectionsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'
const externalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'

describe('macOS Desktop Excel external data provenance oracle', () => {
  it('preserves imported external data provenance through headless import/export', async () => {
    const source = buildExternalDataWorkbookBytes()
    const imported = importXlsx(source, 'external-data-provenance.xlsx')
    const engine = new SpreadsheetEngine({ workbookName: 'external-data-provenance' })
    await engine.ready()

    engine.importSnapshot(imported.snapshot)
    const snapshot = engine.exportSnapshot()
    const exported = exportXlsx(snapshot)

    expect(snapshot.workbook.metadata?.externalConnections).toMatchObject({
      refreshExecution: 'disabled',
      connections: [expect.objectContaining({ name: 'Sales Query', sourceKind: 'database' })],
      externalLinks: [expect.objectContaining({ kind: 'external-workbook', target: 'file:///tmp/source.xlsx' })],
    })
    expect(externalDataMetrics(exported)).toEqual(externalDataMetrics(source))
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'emits external data package parts accepted by Desktop Excel open/save',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const source = buildExternalDataWorkbookBytes()
      const imported = importXlsx(source, 'external-data-provenance.xlsx')
      const engine = new SpreadsheetEngine({ workbookName: 'external-data-provenance-excel-oracle' })
      await engine.ready()
      engine.importSnapshot(imported.snapshot)

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-headless-excel-external-data-provenance-oracle-'))
      try {
        const workbookPath = join(tempDir, 'external-data-provenance.xlsx')
        writeFileSync(workbookPath, exportXlsx(engine.exportSnapshot()))

        runMacosExcelInspectionOracle({
          workbookPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })

        expect(externalDataMetrics(new Uint8Array(readFileSync(workbookPath)))).toEqual(externalDataMetrics(source))
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    120_000,
  )
})

function buildExternalDataWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([['Local'], [1]])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/workbook.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/workbook.xml')).replace(
      '</sheets>',
      '</sheets><externalReferences><externalReference r:id="rIdExternal1"/></externalReferences>',
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      [
        `<Relationship Id="rIdConnections" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
        `<Relationship Id="rIdExternal1" Type="${externalLinkRelationshipType}" Target="externalLinks/externalLink1.xml"/>`,
        '</Relationships>',
      ].join(''),
    ),
  )
  zip['xl/connections.xml'] = strToU8(connectionsXml)
  zip['xl/externalLinks/externalLink1.xml'] = strToU8(externalLinkXml)
  zip['xl/externalLinks/_rels/externalLink1.xml.rels'] = strToU8(externalLinkRelationshipsXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/externalLinks/externalLink1.xml', contentType: externalLinkContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function externalDataMetrics(bytes: Uint8Array): {
  readonly connectionParts: readonly string[]
  readonly externalLinkParts: readonly string[]
  readonly connectionRelationships: number
  readonly externalLinkRelationships: number
  readonly externalLinkPathRelationships: number
  readonly externalReferences: number
} {
  const zip = unzipSync(bytes)
  const workbookXml = readZipTextFromZip(zip, 'xl/workbook.xml')
  const workbookRelationshipsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  const externalLinkRelationshipsXml = readZipTextFromZip(zip, 'xl/externalLinks/_rels/externalLink1.xml.rels')
  return {
    connectionParts: Object.keys(zip).filter((path) => path === 'xl/connections.xml'),
    externalLinkParts: Object.keys(zip)
      .filter((path) => path.startsWith('xl/externalLinks/'))
      .toSorted(),
    connectionRelationships: relationshipsWithType(workbookRelationshipsXml, connectionsRelationshipType),
    externalLinkRelationships: relationshipsWithType(workbookRelationshipsXml, externalLinkRelationshipType),
    externalLinkPathRelationships: relationshipsWithType(externalLinkRelationshipsXml, externalLinkPathRelationshipType),
    externalReferences: [...workbookXml.matchAll(/<externalReference\b/gu)].length,
  }
}

function relationshipsWithType(relationshipsXml: string, relationshipType: string): number {
  return [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].filter((match) => {
    const attributes = match[1] ?? ''
    return readXmlAttribute(attributes, 'Type') === relationshipType
  }).length
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

function upsertContentTypeOverride(xml: string, entry: { readonly partName: string; readonly contentType: string }): string {
  if (xml.includes(`PartName="${entry.partName}"`)) {
    return xml
  }
  return xml.replace('</Types>', `<Override PartName="${entry.partName}" ContentType="${entry.contentType}"/></Types>`)
}

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<workbook\b([^>]*)>/u, `<workbook$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function readXmlAttribute(attributes: string, name: string): string | undefined {
  const match = new RegExp(`\\s${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)
  return match?.[2]
}

const connectionsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<connection id="1" name="Sales Query" type="5" refreshedVersion="8" refreshOnLoad="0">',
  '<dbPr connection="Provider=SQLOLEDB;Data Source=example" command="SELECT * FROM Sales" commandType="2"/>',
  '</connection>',
  '</connections>',
].join('')

const externalLinkXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
  '<externalBook r:id="rId1"><sheetNames><sheetName val="Source"/></sheetNames></externalBook>',
  '</externalLink>',
].join('')

const externalLinkRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rId1" Type="${externalLinkPathRelationshipType}" Target="file:///tmp/source.xlsx" TargetMode="External"/>`,
  '</Relationships>',
].join('')
