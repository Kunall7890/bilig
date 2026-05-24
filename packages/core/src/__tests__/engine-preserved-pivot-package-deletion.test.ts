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
      'xl/pivotCache/pivotCacheDefinition2.xml',
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
              path: 'xl/pivotTables/pivotTable2.xml',
              xml: '<pivotTableDefinition name="KeepPivot" cacheId="2"><location ref="B2:C4"/></pivotTableDefinition>',
            },
            {
              path: 'xl/pivotCache/pivotCacheDefinition2.xml',
              xml: '<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource ref="A1:D4" sheet="Keep"/></cacheSource></pivotCacheDefinition>',
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

function pivotPartPaths(metadata: WorkbookMetadataSnapshot | undefined): string[] {
  return metadata?.pivotArtifacts?.parts.map((part) => part.path).toSorted() ?? []
}

function pivotCacheIds(metadata: WorkbookMetadataSnapshot | undefined): string[] {
  const xml = metadata?.pivotArtifacts?.workbookPivotCachesXml ?? ''
  return [...xml.matchAll(/<pivotCache\b[^>]*\bcacheId="([^"]+)"/gu)].map((match) => match[1] ?? '').filter(Boolean)
}
