import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const connectionsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
const slicerCacheRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
const slicerRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
const tableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const queryTableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable'
const connectionsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'
const slicerCacheContentType = 'application/vnd.ms-excel.slicerCache+xml'
const slicerContentType = 'application/vnd.ms-excel.slicer+xml'
const queryTableContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml'

describe('xlsx slicer and connection artifacts roundtrip', () => {
  it('preserves slicer caches, slicers, and workbook connections as package artifacts', () => {
    const source = buildWorkbookWithSlicerAndConnectionArtifacts()

    const imported = importXlsx(source, 'slicer-connections.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/connections.xml',
      'xl/slicerCaches/_rels/slicerCache1.xml.rels',
      'xl/slicerCaches/slicerCache1.xml',
      'xl/slicers/slicer1.xml',
    ])
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.workbookSlicerCachesExtXml).toBe(workbookSlicerCachesExtXml)
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.workbookRelationships).toEqual([
      { id: 'rId80', type: slicerCacheRelationshipType, target: 'slicerCaches/slicerCache1.xml' },
      { id: 'rId81', type: connectionsRelationshipType, target: 'connections.xml' },
    ])
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts).toEqual([
      {
        sheetName: 'Revenue',
        sheetSlicerListExtXml,
        relationships: [{ id: 'rId20', type: slicerRelationshipType, target: '../slicers/slicer1.xml' }],
      },
    ])
    expect(slicerConnectionMetrics(exported)).toEqual(slicerConnectionMetrics(source))
    expect(readZipText(exported, 'xl/connections.xml')).toBe(connectionsXml)
    expect(readZipText(exported, 'xl/slicerCaches/slicerCache1.xml')).toBe(slicerCacheXml)
    expect(readZipText(exported, 'xl/slicers/slicer1.xml')).toBe(slicerXml)
    expect(readContentTypeOverride(exported, '/xl/connections.xml')).toBe(connectionsContentType)
    expect(readContentTypeOverride(exported, '/xl/slicerCaches/slicerCache1.xml')).toBe(slicerCacheContentType)
    expect(readContentTypeOverride(exported, '/xl/slicers/slicer1.xml')).toBe(slicerContentType)
  })

  it('preserves query-table package topology for external data tables', () => {
    const source = buildWorkbookWithQueryTableArtifacts()

    const imported = importXlsx(source, 'query-table-connections.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/connections.xml',
      'xl/queryTables/queryTable1.xml',
      'xl/tables/_rels/table1.xml.rels',
    ])
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.workbookRelationships).toEqual([
      { id: 'rIdQueryConnections', type: connectionsRelationshipType, target: 'connections.xml' },
    ])
    expect(queryTableMetrics(exported)).toEqual(queryTableMetrics(source))
    expect(readZipText(exported, 'xl/queryTables/queryTable1.xml')).toBe(queryTableXml)
    expect(readZipText(exported, 'xl/tables/_rels/table1.xml.rels')).toBe(queryTableRelationshipsXml)
    expect(readContentTypeOverride(exported, '/xl/connections.xml')).toBe(connectionsContentType)
    expect(readContentTypeOverride(exported, '/xl/queryTables/queryTable1.xml')).toBe(queryTableContentType)
  })

  it('reattaches table-owned query-table relationships when table paths are regenerated', () => {
    const source = buildWorkbookWithTwoTableQueryTableArtifacts()

    expect(tablePathForDisplayName(source, 'ARevenueQuery')).toBe('xl/tables/table2.xml')

    const imported = importXlsx(source, 'two-table-query-table-connections.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.tables?.map((table) => table.name)).toEqual(['ARevenueQuery', 'ZPlain'])
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.tableArtifacts).toEqual([
      {
        tableName: 'ARevenueQuery',
        sheetName: 'Revenue',
        relationshipPartPath: 'xl/tables/_rels/table2.xml.rels',
        relationships: [{ id: 'rIdQueryTable1', type: queryTableRelationshipType, target: '../queryTables/queryTable1.xml' }],
      },
    ])
    expect(tablePathForDisplayName(exported, 'ARevenueQuery')).toBe('xl/tables/table1.xml')
    expect(tablePathForDisplayName(exported, 'ZPlain')).toBe('xl/tables/table2.xml')
    expect(queryTableRelationshipCountForTable(exported, 'ARevenueQuery')).toBe(1)
    expect(queryTableRelationshipCountForTable(exported, 'ZPlain')).toBe(0)
    expect(readZipText(exported, 'xl/queryTables/queryTable1.xml')).toBe(twoTableQueryTableXml)
    expect(readContentTypeOverride(exported, '/xl/queryTables/queryTable1.xml')).toBe(queryTableContentType)
  })

  it('preserves Desktop Excel worksheet-level query-table relationships', () => {
    const source = buildWorkbookWithWorksheetQueryTableArtifacts()

    const imported = importXlsx(source, 'desktop-excel-query-table.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/connections.xml',
      'xl/queryTables/queryTable1.xml',
    ])
    expect(imported.snapshot.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts).toEqual([
      {
        sheetName: 'Revenue',
        relationships: [{ id: 'rIdQueryTable1', type: queryTableRelationshipType, target: '../queryTables/queryTable1.xml' }],
      },
    ])
    expect(queryTableMetrics(exported)).toEqual(queryTableMetrics(source))
    expect(readZipText(exported, 'xl/queryTables/queryTable1.xml')).toBe(desktopExcelQueryTableXml)
    expect(readContentTypeOverride(exported, '/xl/queryTables/queryTable1.xml')).toBe(queryTableContentType)
  })
})

function buildWorkbookWithSlicerAndConnectionArtifacts(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table[1-9][0-9]*\.xml$/u.test(path))
  if (!tablePath) {
    throw new Error('Expected exported workbook to include a table part')
  }
  const tableTarget = `../tables/${tablePath.slice(tablePath.lastIndexOf('/') + 1)}`
  zip['xl/workbook.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/workbook.xml')).replace(
      '</workbook>',
      `<extLst>${workbookSlicerCachesExtXml}</extLst></workbook>`,
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      [
        `<Relationship Id="rId80" Type="${slicerCacheRelationshipType}" Target="slicerCaches/slicerCache1.xml"/>`,
        `<Relationship Id="rId81" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
        '</Relationships>',
      ].join(''),
    ),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    ensureRelationshipNamespace(readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')).replace(
      '</worksheet>',
      `<extLst>${sheetSlicerListExtXml}</extLst></worksheet>`,
    ),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/worksheets/_rels/sheet1.xml.rels'),
      `<Relationship Id="rId20" Type="${slicerRelationshipType}" Target="../slicers/slicer1.xml"/>`,
    ),
  )
  zip['xl/connections.xml'] = strToU8(connectionsXml)
  zip['xl/slicerCaches/slicerCache1.xml'] = strToU8(slicerCacheXml)
  zip['xl/slicerCaches/_rels/slicerCache1.xml.rels'] = strToU8(slicerCacheRelationshipsXml(tableTarget))
  zip['xl/slicers/slicer1.xml'] = strToU8(slicerXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/slicerCaches/slicerCache1.xml', contentType: slicerCacheContentType },
      { partName: '/xl/slicers/slicer1.xml', contentType: slicerContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function buildWorkbookWithQueryTableArtifacts(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  const tablePath = Object.keys(zip).find((path) => /^xl\/tables\/table[1-9][0-9]*\.xml$/u.test(path))
  if (tablePath !== 'xl/tables/table1.xml') {
    throw new Error(`Expected exported workbook table path to be xl/tables/table1.xml, got ${tablePath ?? 'none'}`)
  }
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      `<Relationship Id="rIdQueryConnections" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
    ),
  )
  zip['xl/connections.xml'] = strToU8(connectionsXml)
  zip['xl/tables/_rels/table1.xml.rels'] = strToU8(queryTableRelationshipsXml)
  zip['xl/queryTables/queryTable1.xml'] = strToU8(queryTableXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/queryTables/queryTable1.xml', contentType: queryTableContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function buildWorkbookWithTwoTableQueryTableArtifacts(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildTwoTableWorkbook()))
  if (tablePathForDisplayNameFromZip(zip, 'ZPlain') !== 'xl/tables/table1.xml') {
    throw new Error('Expected ZPlain to start as xl/tables/table1.xml')
  }
  if (tablePathForDisplayNameFromZip(zip, 'ARevenueQuery') !== 'xl/tables/table2.xml') {
    throw new Error('Expected ARevenueQuery to start as xl/tables/table2.xml')
  }
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      `<Relationship Id="rIdQueryConnections" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
    ),
  )
  zip['xl/connections.xml'] = strToU8(connectionsXml)
  zip['xl/tables/_rels/table2.xml.rels'] = strToU8(queryTableRelationshipsXml)
  zip['xl/queryTables/queryTable1.xml'] = strToU8(twoTableQueryTableXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/queryTables/queryTable1.xml', contentType: queryTableContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function buildWorkbookWithWorksheetQueryTableArtifacts(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildPlainWorkbook()))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    appendRelationship(
      readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels'),
      `<Relationship Id="rIdQueryConnections" Type="${connectionsRelationshipType}" Target="connections.xml"/>`,
    ),
  )
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(desktopExcelWorksheetQueryTableRelationshipsXml)
  zip['xl/connections.xml'] = strToU8(desktopExcelTextConnectionXml)
  zip['xl/queryTables/queryTable1.xml'] = strToU8(desktopExcelQueryTableXml)
  zip['[Content_Types].xml'] = strToU8(
    [
      { partName: '/xl/connections.xml', contentType: connectionsContentType },
      { partName: '/xl/queryTables/queryTable1.xml', contentType: queryTableContentType },
    ].reduce((xml, entry) => upsertContentTypeOverride(xml, entry), readZipTextFromZip(zip, '[Content_Types].xml')),
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Slicer connection artifacts',
      metadata: {
        tables: [
          {
            name: 'RevenueTable',
            sheetName: 'Revenue',
            startAddress: 'A1',
            endAddress: 'B4',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'North' },
          { address: 'B2', value: 1200 },
          { address: 'A3', value: 'South' },
          { address: 'B3', value: 900 },
          { address: 'A4', value: 'North' },
          { address: 'B4', value: 300 },
        ],
      },
    ],
  }
}

function buildPlainWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Desktop Excel query table',
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'North' },
          { address: 'B2', value: 1200 },
          { address: 'A3', value: 'South' },
          { address: 'B3', value: 900 },
        ],
      },
    ],
  }
}

function buildTwoTableWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Two table query artifacts',
      metadata: {
        tables: [
          {
            name: 'ZPlain',
            sheetName: 'Revenue',
            startAddress: 'A1',
            endAddress: 'B3',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
          {
            name: 'ARevenueQuery',
            sheetName: 'Revenue',
            startAddress: 'D1',
            endAddress: 'E3',
            columnNames: ['Region', 'Amount'],
            headerRow: true,
            totalsRow: false,
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [
          { address: 'A1', value: 'Region' },
          { address: 'B1', value: 'Amount' },
          { address: 'A2', value: 'West' },
          { address: 'B2', value: 400 },
          { address: 'D1', value: 'Region' },
          { address: 'E1', value: 'Amount' },
          { address: 'D2', value: 'North' },
          { address: 'E2', value: 1200 },
        ],
      },
    ],
  }
}

function slicerConnectionMetrics(bytes: Uint8Array): {
  packageParts: string[]
  sheetSlicerRelationships: number
  sheetSlicerRefs: number
  workbookConnectionsRelationships: number
  workbookSlicerCacheRelationships: number
  workbookSlicerCacheRefs: number
} {
  const zip = unzipSync(bytes)
  const workbookXml = readZipTextFromZip(zip, 'xl/workbook.xml')
  const workbookRelationshipsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  const sheetXml = readZipTextFromZip(zip, 'xl/worksheets/sheet1.xml')
  const sheetRelationshipsXml = readZipTextFromZip(zip, 'xl/worksheets/_rels/sheet1.xml.rels')
  return {
    packageParts: Object.keys(zip)
      .filter((path) => path === 'xl/connections.xml' || path.startsWith('xl/slicerCaches/') || path.startsWith('xl/slicers/'))
      .toSorted(),
    sheetSlicerRelationships: relationshipsWithType(sheetRelationshipsXml, slicerRelationshipType).length,
    sheetSlicerRefs: [...sheetXml.matchAll(/<x14:slicer\b/gu)].length,
    workbookConnectionsRelationships: relationshipsWithType(workbookRelationshipsXml, connectionsRelationshipType).length,
    workbookSlicerCacheRelationships: relationshipsWithType(workbookRelationshipsXml, slicerCacheRelationshipType).length,
    workbookSlicerCacheRefs: [...workbookXml.matchAll(/<x15:slicerCache\b/gu)].length,
  }
}

function queryTableMetrics(bytes: Uint8Array): {
  packageParts: string[]
  queryTableConnectionIds: string[]
  tableQueryTableRelationships: number
  worksheetQueryTableRelationships: number
  workbookConnectionsRelationships: number
} {
  const zip = unzipSync(bytes)
  const workbookRelationshipsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  const tableRelationshipsXml = readOptionalZipTextFromZip(zip, 'xl/tables/_rels/table1.xml.rels') ?? ''
  const worksheetRelationshipsXml = readOptionalZipTextFromZip(zip, 'xl/worksheets/_rels/sheet1.xml.rels') ?? ''
  const queryTableXml = readOptionalZipTextFromZip(zip, 'xl/queryTables/queryTable1.xml') ?? ''
  return {
    packageParts: Object.keys(zip)
      .filter((path) => path === 'xl/connections.xml' || path === 'xl/tables/_rels/table1.xml.rels' || path.startsWith('xl/queryTables/'))
      .toSorted(),
    queryTableConnectionIds: [...queryTableXml.matchAll(/<queryTable\b([^>]*)>/gu)].flatMap((match) => {
      const connectionId = readXmlAttribute(match[1] ?? '', 'connectionId')
      return connectionId ? [connectionId] : []
    }),
    tableQueryTableRelationships: relationshipsWithType(tableRelationshipsXml, queryTableRelationshipType).length,
    worksheetQueryTableRelationships: relationshipsWithType(worksheetRelationshipsXml, queryTableRelationshipType).length,
    workbookConnectionsRelationships: relationshipsWithType(workbookRelationshipsXml, connectionsRelationshipType).length,
  }
}

function queryTableRelationshipCountForTable(bytes: Uint8Array, tableName: string): number {
  const zip = unzipSync(bytes)
  const tablePath = tablePathForDisplayNameFromZip(zip, tableName)
  if (!tablePath) {
    throw new Error(`Missing table path for ${tableName}`)
  }
  const relationshipsXml = readOptionalZipTextFromZip(zip, relationshipPartPath(tablePath)) ?? ''
  return relationshipsWithType(relationshipsXml, queryTableRelationshipType).length
}

function tablePathForDisplayName(bytes: Uint8Array, tableName: string): string | undefined {
  return tablePathForDisplayNameFromZip(unzipSync(bytes), tableName)
}

function tablePathForDisplayNameFromZip(zip: Record<string, Uint8Array>, tableName: string): string | undefined {
  return Object.entries(zip)
    .filter(([path]) => /^xl\/tables\/table[1-9][0-9]*\.xml$/u.test(path))
    .find(([, data]) => readXmlAttribute(strFromU8(data), 'displayName') === tableName)?.[0]
}

function relationshipPartPath(partPath: string): string {
  const slashIndex = partPath.lastIndexOf('/')
  const directory = slashIndex >= 0 ? partPath.slice(0, slashIndex) : ''
  const fileName = slashIndex >= 0 ? partPath.slice(slashIndex + 1) : partPath
  return directory.length > 0 ? `${directory}/_rels/${fileName}.rels` : `_rels/${fileName}.rels`
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

function readOptionalZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string | undefined {
  const bytes = zip[path]
  return bytes ? strFromU8(bytes) : undefined
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
  return xml.replace(/<([A-Za-z0-9:]+)\b([^>]*)>/u, `<$1$2 xmlns:r="${officeRelationshipNamespace}">`)
}

function appendRelationship(relationshipsXml: string, relationshipXml: string): string {
  return relationshipsXml.replace('</Relationships>', `${relationshipXml}</Relationships>`)
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

const workbookSlicerCachesExtXml = [
  '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" ',
  'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">',
  '<x15:slicerCaches><x15:slicerCache r:id="rId80"/></x15:slicerCaches>',
  '</ext>',
].join('')

const sheetSlicerListExtXml = [
  '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" ',
  'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">',
  '<x14:slicerList><x14:slicer r:id="rId20"/></x14:slicerList>',
  '</ext>',
].join('')

const connectionsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1">',
  '<connection id="1" name="Revenue connection" type="5" refreshedVersion="8" background="1">',
  '<dbPr connection="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=revenue.xlsx" command="SELECT * FROM Revenue" commandType="2"/>',
  '</connection>',
  '</connections>',
].join('')

const queryTableRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rIdQueryTable1" Type="${queryTableRelationshipType}" Target="../queryTables/queryTable1.xml"/>`,
  '</Relationships>',
].join('')

const queryTableXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
  'name="RevenueTable" headers="1" rowNumbers="0" disableRefresh="1" backgroundRefresh="0" connectionId="1">',
  '<queryTableRefresh nextId="3">',
  '<queryTableFields count="2">',
  '<queryTableField id="1" name="Region" tableColumnId="1"/>',
  '<queryTableField id="2" name="Amount" tableColumnId="2"/>',
  '</queryTableFields>',
  '</queryTableRefresh>',
  '</queryTable>',
].join('')

const twoTableQueryTableXml = queryTableXml.replace('name="RevenueTable"', 'name="ARevenueQuery"')

const desktopExcelWorksheetQueryTableRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rIdQueryTable1" Type="${queryTableRelationshipType}" Target="../queryTables/queryTable1.xml"/>`,
  '</Relationships>',
].join('')

const desktopExcelTextConnectionXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr16" ',
  'xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16">',
  '<connection id="1" xr16:uid="{881E9239-E69B-3140-9DE4-947575726693}" ',
  'name="bilig-query-table-revenue" type="6" refreshedVersion="8" background="1" saveData="1">',
  '<textPr codePage="10000" sourceFile="/Users/gregkonush/Downloads/bilig-query-table-revenue.csv" comma="1">',
  '<textFields count="2"><textField/><textField/></textFields>',
  '</textPr>',
  '</connection>',
  '</connections>',
].join('')

const desktopExcelQueryTableXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr16" ',
  'xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16" ',
  'name="bilig-query-table-revenue" connectionId="1" xr16:uid="{29AC6D69-DFA8-664F-AA09-24F79E4E8EE2}" ',
  'autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="1" ',
  'applyPatternFormats="1" applyAlignmentFormats="0" applyWidthHeightFormats="0"/>',
].join('')

const slicerCacheXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ',
  `xmlns:r="${officeRelationshipNamespace}" name="Slicer_Region" sourceName="Region" cache="Slicer_Region" r:id="rId1">`,
  '<tableSlicerCache tableId="1" column="1">',
  '<items count="2"><i x="0" s="1"/><i x="1"/></items>',
  '</tableSlicerCache>',
  '</slicerCacheDefinition>',
].join('')

function slicerCacheRelationshipsXml(tableTarget: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    `<Relationship Id="rId1" Type="${tableRelationshipType}" Target="${tableTarget}"/>`,
    '</Relationships>',
  ].join('')
}

const slicerXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ',
  'name="Slicer_Region" cache="Slicer_Region" caption="Region" startItem="0" columnCount="1"/>',
].join('')
