import { describe, expect, it } from 'vitest'

import type { WorkbookMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const officeRelationshipTypePrefix = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('engine preserved pivot package deletion', () => {
  it('prunes deleted sheet pivot package parts and orphan pivot caches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'preserved-pivot-package-delete' })
    await engine.ready()

    engine.importSnapshot(preservedPivotPackageDeletionSnapshot())
    engine.deleteSheet('Data')

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => sheet.name)).toEqual(['Keep'])
    expect(pivotPartPaths(exported.workbook.metadata)).toEqual([
      'xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels',
      'xl/pivotCache/pivotCacheDefinition2.xml',
      'xl/pivotCache/pivotCacheRecords2.xml',
      'xl/pivotTables/pivotTable2.xml',
    ])
    expect(pivotCacheIds(exported.workbook.metadata)).toEqual(['2'])
    expect(exported.workbook.metadata?.pivotArtifacts?.workbookRelationships).toEqual([
      {
        id: 'rIdPivotCache2',
        type: `${officeRelationshipTypePrefix}/pivotCacheDefinition`,
        target: 'pivotCache/pivotCacheDefinition2.xml',
      },
    ])
    expect(exported.workbook.metadata?.unsupportedPivots?.map((pivot) => pivot.sheetName)).toEqual(['Keep'])
    expect(exported.workbook.metadata?.unsupportedFormulaDependencies?.map((entry) => entry.sheetName)).toEqual(['Keep'])
    expect(exported.sheets[0]?.metadata?.pivotArtifacts?.relationships.map((relationship) => relationship.target)).toEqual([
      '../pivotTables/pivotTable2.xml',
    ])
  })

  it('prunes pivot-owned external model connection artifacts when deleting the only owning pivot sheet', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'preserved-external-model-pivot-delete' })
    await engine.ready()

    engine.importSnapshot(externalModelPivotConnectionDeletionSnapshot({ includeSurvivingPivot: false }))
    engine.deleteSheet('Pivot')

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => sheet.name)).toEqual(['Data'])
    expect(exported.workbook.metadata?.pivotArtifacts).toBeUndefined()
    expect(exported.workbook.metadata?.externalConnections).toBeUndefined()
    expect(exported.workbook.metadata?.slicerConnectionArtifacts).toBeUndefined()
    expect(exported.workbook.metadata?.dataModelArtifacts).toBeUndefined()
  })

  it('keeps shared external model connection artifacts when a surviving pivot cache still owns them', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'preserved-shared-external-model-pivot-delete' })
    await engine.ready()

    engine.importSnapshot(externalModelPivotConnectionDeletionSnapshot({ includeSurvivingPivot: true }))
    engine.deleteSheet('Pivot')

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Keep'])
    expect(pivotPartPaths(exported.workbook.metadata)).toEqual([
      'xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels',
      'xl/pivotCache/pivotCacheDefinition2.xml',
      'xl/pivotCache/pivotCacheRecords2.xml',
      'xl/pivotTables/pivotTable2.xml',
    ])
    expect(exported.workbook.metadata?.externalConnections?.connections?.map((connection) => connection.id)).toEqual([1])
    expect(slicerConnectionPartPaths(exported.workbook.metadata)).toEqual(['xl/connections.xml'])
    expect(dataModelPartPaths(exported.workbook.metadata)).toEqual(['xl/model/item.data'])
    expect(connectionIdsInPreservedConnectionsXml(exported.workbook.metadata)).toEqual(['1'])
  })
})

function preservedPivotPackageDeletionSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Preserved pivot package deletion',
      metadata: {
        unsupportedFormulaDependencies: [
          {
            context: 'worksheet-cell',
            clause: '18.3.1.40',
            formula: 'A1+1',
            importedFormula: 'A1+1',
            sheetName: 'Data',
            address: 'A2',
            cacheStatus: 'trustedCached',
          },
          {
            context: 'worksheet-cell',
            clause: '18.3.1.40',
            formula: 'A1+1',
            importedFormula: 'A1+1',
            sheetName: 'Keep',
            address: 'A2',
            cacheStatus: 'trustedCached',
          },
        ],
        unsupportedPivots: [
          {
            kind: 'raw-part',
            reason: 'preserved pivot on deleted sheet',
            packagePart: 'xl/pivotTables/pivotTable1.xml',
            sheetName: 'Data',
            address: 'B2',
            name: 'DeletedPivot',
          },
          {
            kind: 'raw-part',
            reason: 'preserved pivot on surviving sheet',
            packagePart: 'xl/pivotTables/pivotTable2.xml',
            sheetName: 'Keep',
            address: 'B2',
            name: 'KeepPivot',
          },
        ],
        pivotArtifacts: {
          parts: [
            {
              path: 'xl/pivotTables/pivotTable1.xml',
              xml: '<pivotTableDefinition name="DeletedPivot" cacheId="1"><location ref="B2:C4"/></pivotTableDefinition>',
            },
            {
              path: 'xl/pivotCache/pivotCacheDefinition1.xml',
              xml: '<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource ref="A1:D4" sheet="Data"/></cacheSource></pivotCacheDefinition>',
            },
            {
              path: 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
              xml: [
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                '<Relationship Id="rIdRecords1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>',
                '</Relationships>',
              ].join(''),
            },
            {
              path: 'xl/pivotCache/pivotCacheRecords1.xml',
              xml: '<pivotCacheRecords count="1"><r><s v="deleted"/></r></pivotCacheRecords>',
            },
            {
              path: 'xl/pivotTables/pivotTable2.xml',
              xml: '<pivotTableDefinition name="KeepPivot" cacheId="2"><location ref="B2:C4"/></pivotTableDefinition>',
            },
            {
              path: 'xl/pivotCache/pivotCacheDefinition2.xml',
              xml: '<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource ref="A1:D4" sheet="Keep"/></cacheSource></pivotCacheDefinition>',
            },
            {
              path: 'xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels',
              xml: [
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                '<Relationship Id="rIdRecords2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords2.xml"/>',
                '</Relationships>',
              ].join(''),
            },
            {
              path: 'xl/pivotCache/pivotCacheRecords2.xml',
              xml: '<pivotCacheRecords count="1"><r><s v="keep"/></r></pivotCacheRecords>',
            },
          ],
          workbookPivotCachesXml:
            '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/><pivotCache cacheId="2" r:id="rIdPivotCache2"/></pivotCaches>',
          workbookRelationships: [
            {
              id: 'rIdPivotCache1',
              type: `${officeRelationshipTypePrefix}/pivotCacheDefinition`,
              target: 'pivotCache/pivotCacheDefinition1.xml',
            },
            {
              id: 'rIdPivotCache2',
              type: `${officeRelationshipTypePrefix}/pivotCacheDefinition`,
              target: 'pivotCache/pivotCacheDefinition2.xml',
            },
          ],
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        metadata: {
          pivotArtifacts: {
            relationships: [
              {
                id: 'rIdPivotTable1',
                type: `${officeRelationshipTypePrefix}/pivotTable`,
                target: '../pivotTables/pivotTable1.xml',
              },
            ],
            pivotTableDefinitionsXml: '<pivotTableDefinition r:id="rIdPivotTable1"/>',
          },
        },
        cells: [{ address: 'A1', value: 'deleted' }],
      },
      {
        id: 2,
        name: 'Keep',
        order: 1,
        metadata: {
          pivotArtifacts: {
            relationships: [
              {
                id: 'rIdPivotTable2',
                type: `${officeRelationshipTypePrefix}/pivotTable`,
                target: '../pivotTables/pivotTable2.xml',
              },
            ],
            pivotTableDefinitionsXml: '<pivotTableDefinition r:id="rIdPivotTable2"/>',
          },
        },
        cells: [{ address: 'A1', value: 'keep' }],
      },
    ],
  }
}

function externalModelPivotConnectionDeletionSnapshot(options: { readonly includeSurvivingPivot: boolean }): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'External model pivot connection deletion',
      metadata: {
        externalConnections: {
          refreshExecution: 'disabled',
          connections: [
            {
              id: 1,
              name: 'ThisWorkbookDataModel',
              sourceKind: 'model',
              type: '5',
              connection: 'Provider=MSOLAP.8;Data Source=$Workbook$;Initial Catalog=Model',
              command: 'Model',
              commandType: '1',
              saveData: true,
              clause: '18.13',
            },
          ],
        },
        dataModelArtifacts: {
          parts: [encodedTextPart('xl/model/item.data', 'data-model-payload')],
          workbookRelationships: [
            {
              id: 'rIdModel',
              type: `${officeRelationshipTypePrefix}/powerPivotData`,
              target: 'model/item.data',
            },
          ],
          contentTypeDefaults: [{ extension: 'data', contentType: 'application/octet-stream' }],
        },
        slicerConnectionArtifacts: {
          parts: [encodedTextPart('xl/connections.xml', modelConnectionXml())],
          workbookRelationships: [
            {
              id: 'rIdConnections',
              type: `${officeRelationshipTypePrefix}/connections`,
              target: 'connections.xml',
            },
          ],
          contentTypeOverrides: [
            {
              partName: '/xl/connections.xml',
              contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml',
            },
          ],
        },
        pivotArtifacts: {
          parts: [
            ...externalModelPivotParts({ cacheId: '1', pivotTablePath: 'xl/pivotTables/pivotTable1.xml' }),
            ...(options.includeSurvivingPivot
              ? externalModelPivotParts({ cacheId: '2', pivotTablePath: 'xl/pivotTables/pivotTable2.xml' })
              : []),
          ],
          workbookPivotCachesXml: options.includeSurvivingPivot
            ? '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/><pivotCache cacheId="2" r:id="rIdPivotCache2"/></pivotCaches>'
            : '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/></pivotCaches>',
          workbookRelationships: [
            {
              id: 'rIdPivotCache1',
              type: `${officeRelationshipTypePrefix}/pivotCacheDefinition`,
              target: 'pivotCache/pivotCacheDefinition1.xml',
            },
            ...(options.includeSurvivingPivot
              ? [
                  {
                    id: 'rIdPivotCache2',
                    type: `${officeRelationshipTypePrefix}/pivotCacheDefinition`,
                    target: 'pivotCache/pivotCacheDefinition2.xml',
                  },
                ]
              : []),
          ],
        },
      },
    },
    sheets: [
      { id: 1, name: 'Data', order: 0, cells: [{ address: 'A1', value: 'data' }] },
      {
        id: 2,
        name: 'Pivot',
        order: 1,
        metadata: {
          pivotArtifacts: {
            relationships: [
              {
                id: 'rIdPivotTable1',
                type: `${officeRelationshipTypePrefix}/pivotTable`,
                target: '../pivotTables/pivotTable1.xml',
              },
            ],
            pivotTableDefinitionsXml: '<pivotTableDefinition r:id="rIdPivotTable1"/>',
          },
        },
        cells: [{ address: 'A1', value: 'delete me' }],
      },
      ...(options.includeSurvivingPivot
        ? [
            {
              id: 3,
              name: 'Keep',
              order: 2,
              metadata: {
                pivotArtifacts: {
                  relationships: [
                    {
                      id: 'rIdPivotTable2',
                      type: `${officeRelationshipTypePrefix}/pivotTable`,
                      target: '../pivotTables/pivotTable2.xml',
                    },
                  ],
                  pivotTableDefinitionsXml: '<pivotTableDefinition r:id="rIdPivotTable2"/>',
                },
              },
              cells: [{ address: 'A1', value: 'keep' }],
            },
          ]
        : []),
    ],
  }
}

function externalModelPivotParts(input: { readonly cacheId: '1' | '2'; readonly pivotTablePath: string }) {
  return [
    {
      path: input.pivotTablePath,
      xml: `<pivotTableDefinition name="ModelPivot${input.cacheId}" cacheId="${input.cacheId}"><location ref="A1:B3"/></pivotTableDefinition>`,
    },
    {
      path: `xl/pivotCache/pivotCacheDefinition${input.cacheId}.xml`,
      xml: [
        '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="1">',
        '<cacheSource type="external" connectionId="1"/>',
        '<cacheFields count="1"><cacheField name="Region"><sharedItems count="1"><s v="East"/></sharedItems></cacheField></cacheFields>',
        '</pivotCacheDefinition>',
      ].join(''),
    },
    {
      path: `xl/pivotCache/_rels/pivotCacheDefinition${input.cacheId}.xml.rels`,
      xml: [
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        `<Relationship Id="rIdRecords${input.cacheId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords${input.cacheId}.xml"/>`,
        '</Relationships>',
      ].join(''),
    },
    {
      path: `xl/pivotCache/pivotCacheRecords${input.cacheId}.xml`,
      xml: '<pivotCacheRecords count="1"><r><s v="East"/></r></pivotCacheRecords>',
    },
  ]
}

function modelConnectionXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1">',
    '<connection id="1" name="ThisWorkbookDataModel" description="Embedded Data Model" type="5" model="1" saveData="1">',
    '<dbPr connection="Provider=MSOLAP.8;Data Source=$Workbook$;Initial Catalog=Model" command="Model" commandType="1"/>',
    '<olapPr sendLocale="1" rowDrillCount="1000"/>',
    '</connection>',
    '</connections>',
  ].join('')
}

function pivotPartPaths(metadata: WorkbookMetadataSnapshot | undefined): string[] {
  return metadata?.pivotArtifacts?.parts.map((part) => part.path).toSorted() ?? []
}

function pivotCacheIds(metadata: WorkbookMetadataSnapshot | undefined): string[] {
  const xml = metadata?.pivotArtifacts?.workbookPivotCachesXml ?? ''
  return [...xml.matchAll(/<pivotCache\b[^>]*\bcacheId="([^"]+)"/gu)].map((match) => match[1] ?? '').filter(Boolean)
}

function slicerConnectionPartPaths(metadata: WorkbookMetadataSnapshot | undefined): string[] {
  return metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path).toSorted() ?? []
}

function dataModelPartPaths(metadata: WorkbookMetadataSnapshot | undefined): string[] {
  return metadata?.dataModelArtifacts?.parts.map((part) => part.path).toSorted() ?? []
}

function connectionIdsInPreservedConnectionsXml(metadata: WorkbookMetadataSnapshot | undefined): string[] {
  const part = metadata?.slicerConnectionArtifacts?.parts.find((entry) => entry.path === 'xl/connections.xml')
  const xml = part ? decodedTextPart(part.dataBase64) : ''
  return [...xml.matchAll(/<connection\b[^>]*\bid=(["'])([^"']+)\1/gu)].map((match) => match[2] ?? '').filter(Boolean)
}

function encodedTextPart(path: string, text: string) {
  const bytes = Buffer.from(text, 'utf8')
  return {
    path,
    storage: 'base64' as const,
    dataBase64: bytes.toString('base64'),
    byteLength: bytes.byteLength,
  }
}

function decodedTextPart(dataBase64: string): string {
  return Buffer.from(dataBase64, 'base64').toString('utf8')
}
