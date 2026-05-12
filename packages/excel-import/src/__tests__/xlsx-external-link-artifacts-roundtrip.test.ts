import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const externalLinkPathRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'
const externalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'

describe('xlsx external link artifacts roundtrip', () => {
  it('preserves external workbook link package artifacts', () => {
    const source = buildWorkbookWithExternalLinkArtifacts()

    const imported = importXlsx(source, 'external-links.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.externalLinkArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/externalLinks/_rels/externalLink1.xml.rels',
      'xl/externalLinks/externalLink1.xml',
    ])
    expect(imported.snapshot.workbook.metadata?.externalLinkArtifacts?.workbookExternalReferencesXml).toBe(
      '<externalReferences><externalReference r:id="rId99"/></externalReferences>',
    )
    expect(imported.snapshot.workbook.metadata?.externalLinkArtifacts?.workbookRelationships).toEqual([
      { id: 'rId99', type: externalLinkRelationshipType, target: 'externalLinks/externalLink1.xml' },
    ])
    expect(externalLinkMetrics(exported)).toEqual(externalLinkMetrics(source))
    expect(readZipText(exported, 'xl/externalLinks/externalLink1.xml')).toBe(externalLinkXml)
    expect(readZipText(exported, 'xl/externalLinks/_rels/externalLink1.xml.rels')).toBe(externalLinkRelationshipsXml)
    expect(readWorkbookRelationship(exported, externalLinkRelationshipType)).toMatchObject({
      target: 'externalLinks/externalLink1.xml',
      type: externalLinkRelationshipType,
    })
    expect(
      readRelationshipTargetMode(readZipText(exported, 'xl/externalLinks/_rels/externalLink1.xml.rels'), externalLinkPathRelationshipType),
    ).toBe('External')
    expect(readContentTypeOverride(exported, '/xl/externalLinks/externalLink1.xml')).toBe(externalLinkContentType)
  })
})

function buildWorkbookWithExternalLinkArtifacts(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  zip['xl/workbook.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/workbook.xml')).replace(
      '</sheets>',
      '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      `<Relationship Id="rId99" Type="${externalLinkRelationshipType}" Target="externalLinks/externalLink1.xml"/></Relationships>`,
    ),
  )
  zip['xl/externalLinks/externalLink1.xml'] = strToU8(externalLinkXml)
  zip['xl/externalLinks/_rels/externalLink1.xml.rels'] = strToU8(externalLinkRelationshipsXml)
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(readZipTextFromZip(zip, '[Content_Types].xml'), {
      partName: '/xl/externalLinks/externalLink1.xml',
      contentType: externalLinkContentType,
    }),
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'External links' },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        cells: [{ address: 'A1', value: 'External link fixture' }],
      },
    ],
  }
}

function externalLinkMetrics(bytes: Uint8Array): {
  externalLinkPackageParts: string[]
  externalLinkPathRelationships: number
  workbookExternalLinkRelationships: number
  workbookExternalReferences: number
} {
  const zip = unzipSync(bytes)
  const workbookXml = readZipTextFromZip(zip, 'xl/workbook.xml')
  const workbookRelationshipsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  const externalLinkRelationshipsXml = readZipTextFromZip(zip, 'xl/externalLinks/_rels/externalLink1.xml.rels')
  return {
    externalLinkPackageParts: Object.keys(zip)
      .filter((path) => path.startsWith('xl/externalLinks/'))
      .toSorted(),
    externalLinkPathRelationships: relationshipsWithType(externalLinkRelationshipsXml, externalLinkPathRelationshipType).length,
    workbookExternalLinkRelationships: relationshipsWithType(workbookRelationshipsXml, externalLinkRelationshipType).length,
    workbookExternalReferences: [...workbookXml.matchAll(/<externalReference\b/gu)].length,
  }
}

function relationshipsWithType(relationshipsXml: string, relationshipType: string): string[] {
  return [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    return readXmlAttribute(attributes, 'Type') === relationshipType ? [match[0]] : []
  })
}

function readZipText(bytes: Uint8Array, path: string): string {
  return readZipTextFromZip(unzipSync(bytes), path)
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}

function readWorkbookRelationship(bytes: Uint8Array, relationshipType: string): { target: string; type: string } | undefined {
  const relationshipsXml = readZipText(bytes, 'xl/_rels/workbook.xml.rels')
  for (const match of relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'Type') === relationshipType) {
      return {
        target: readXmlAttribute(attributes, 'Target') ?? '',
        type: readXmlAttribute(attributes, 'Type') ?? '',
      }
    }
  }
  return undefined
}

function readRelationshipTargetMode(relationshipsXml: string, relationshipType: string): string | undefined {
  for (const match of relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'Type') === relationshipType) {
      return readXmlAttribute(attributes, 'TargetMode') ?? undefined
    }
  }
  return undefined
}

function readContentTypeOverride(bytes: Uint8Array, partName: string): string | undefined {
  const contentTypesXml = readZipText(bytes, '[Content_Types].xml')
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'PartName') === partName) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
}

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<workbook\b([^>]*)>/u, `<workbook$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function upsertContentTypeOverride(
  contentTypesXml: string,
  input: {
    readonly partName: string
    readonly contentType: string
  },
): string {
  if (contentTypesXml.includes(`PartName="${input.partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${input.partName}" ContentType="${input.contentType}"/></Types>`)
}

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

const externalLinkXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
  '<externalBook r:id="rId1">',
  '<sheetNames><sheetName val="Source"/></sheetNames>',
  '<sheetDataSet><sheetData sheetId="0">',
  '<row r="1"><cell r="A1"><v>42</v></cell></row>',
  '</sheetData></sheetDataSet>',
  '</externalBook>',
  '</externalLink>',
].join('')

const externalLinkRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rId1" Type="${externalLinkPathRelationshipType}" Target="file:///tmp/source.xlsx" TargetMode="External"/>`,
  '</Relationships>',
].join('')
