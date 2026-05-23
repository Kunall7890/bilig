import { Buffer } from 'node:buffer'

import { describe, expect, it } from 'vitest'

import type { WorkbookMetadataSnapshot, WorkbookPreservedPackagePartSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const officeRelationshipTypePrefix = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('engine imported package metadata preservation', () => {
  it('keeps import/export-only workbook and worksheet metadata after engine edits and restore', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-preservation' })
    await engine.ready()

    engine.importSnapshot(packageMetadataSnapshot())
    engine.setCellValue('Data', 'B1', 'headless edit')

    const exported = engine.exportSnapshot()
    expect(exported.workbook.metadata).toMatchObject(preservedWorkbookMetadata)
    expect(exported.sheets[0]?.metadata).toMatchObject(preservedSheetMetadata)

    const restored = new SpreadsheetEngine({ workbookName: 'package-metadata-preservation-restored' })
    await restored.ready()
    restored.importSnapshot(exported)

    expect(restored.exportSnapshot().workbook.metadata).toMatchObject(preservedWorkbookMetadata)
    expect(restored.exportSnapshot().sheets[0]?.metadata).toMatchObject(preservedSheetMetadata)
  })

  it('structurally rewrites preserved worksheet style artifacts while keeping sheet view refs stable', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-structural-rewrite' })
    await engine.ready()

    engine.importSnapshot(packageMetadataSnapshot())
    engine.insertRows('Data', 0, 1)
    engine.insertColumns('Data', 0, 1)

    const metadata = engine.exportSnapshot().sheets[0]?.metadata
    expect(metadata?.styleArtifacts).toEqual({
      cellStyleIndexes: [{ address: 'B2', styleIndex: 1 }],
      blankCellAddresses: ['D4'],
    })
    expect(metadata?.viewState).toEqual({
      sheetViewsXml:
        '<sheetViews><sheetView workbookViewId="0" topLeftCell="B2" tabSelected="1"><selection activeCell="C3" sqref="C3 D4:E5"/></sheetView></sheetViews>',
    })
  })

  it('structurally rewrites preserved pivot package output and source refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-pivot-artifact-structural-rewrite' })
    await engine.ready()

    engine.importSnapshot(packageMetadataSnapshot())
    engine.insertRows('Data', 0, 1)
    engine.insertColumns('Data', 0, 1)

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(pivotLocationRef(metadata, 'xl/pivotTables/pivotTable1.xml')).toBe('C3:D5')
    expect(pivotCacheSourceRef(metadata, 'xl/pivotCache/pivotCacheDefinition1.xml')).toBe('B2:E5')
  })

  it('drops deleted preserved worksheet style refs without shifting sheet view refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-deleted-structural-rewrite' })
    await engine.ready()

    engine.importSnapshot(packageMetadataSnapshot())
    engine.deleteRows('Data', 2, 1)

    const metadata = engine.exportSnapshot().sheets[0]?.metadata
    expect(metadata?.styleArtifacts).toEqual({
      cellStyleIndexes: [{ address: 'A1', styleIndex: 1 }],
    })
    expect(metadata?.viewState).toEqual({
      sheetViewsXml:
        '<sheetViews><sheetView workbookViewId="0" topLeftCell="B2" tabSelected="1"><selection activeCell="C3" sqref="C3 D4:E5"/></sheetView></sheetViews>',
    })
  })

  it('treats preserved pivot package refs as structural delete impact even without cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-pivot-artifact-delete-impact' })
    await engine.ready()

    engine.importSnapshot(rawPivotOnlySnapshot())
    engine.deleteRows('Data', 0, 1)

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(pivotLocationRef(metadata, 'xl/pivotTables/pivotTable1.xml')).toBe('B1:C3')
    expect(pivotCacheSourceRef(metadata, 'xl/pivotCache/pivotCacheDefinition1.xml')).toBe('A1:D3')
  })
})

const preservedWorkbookMetadata = {
  documentPropertyArtifacts: {
    core: {
      path: 'docProps/core.xml',
      xml: '<cp:coreProperties><dc:title>Preserved</dc:title></cp:coreProperties>',
      relationship: {
        id: 'rIdCore',
        type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
        target: 'docProps/core.xml',
      },
      contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
    },
  },
  externalWorkbookReferences: [
    {
      bookIndex: 1,
      packagePath: 'xl/externalLinks/externalLink1.xml',
      target: 'linked.xlsx',
      targetMode: 'External',
      workbookName: 'linked.xlsx',
      sheetNames: ['Rates'],
    },
  ],
  unsupportedFormulaDependencies: [
    {
      kind: 'external-workbook-reference',
      sheetName: 'Data',
      address: 'A1',
      formula: '[1]Rates!A1',
      importedFormula: '[1]Rates!A1',
      linkedWorkbooks: [
        {
          bookIndex: 1,
          target: 'linked.xlsx',
          targetMode: 'External',
          workbookName: 'linked.xlsx',
          sheetNames: ['Rates'],
        },
      ],
      cachedValuesUsed: true,
      cachedFormulaValuePreserved: true,
      cachedExternalReferenceValuesUsed: true,
      resolvedExternalReferenceCount: 1,
      unresolvedExternalReferenceCount: 0,
      reason: 'Preserved external workbook cache',
    },
  ],
  unsupportedPivots: [
    {
      kind: 'raw-part',
      reason: 'external cache preserved as package artifact',
      packagePart: 'xl/pivotTables/pivotTable1.xml',
      sheetName: 'Data',
      address: 'D4',
      name: 'PivotTable1',
    },
  ],
  formulaAudit: {
    formulas: [
      {
        context: 'worksheet-cell',
        clause: '18.3.1.40',
        formula: 'A1+1',
        sheetName: 'Data',
        address: 'A2',
        cacheStatus: 'trustedCached',
      },
    ],
    calcChain: {
      packagePath: 'xl/calcChain.xml',
      cells: [{ sheetIndex: 0, sheetName: 'Data', address: 'A2' }],
    },
  },
  externalConnections: {
    refreshExecution: 'disabled',
    connections: [{ id: 1, name: 'PowerQuery', sourceKind: 'model', refreshOnLoad: true, clause: '18.13' }],
  },
  pivotArtifacts: {
    parts: [
      {
        path: 'xl/pivotTables/pivotTable1.xml',
        xml: '<pivotTableDefinition name="PivotTable1"><location ref="B2:C4" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/></pivotTableDefinition>',
      },
      {
        path: 'xl/pivotCache/pivotCacheDefinition1.xml',
        xml: '<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource ref="A1:D4" sheet="Data"/></cacheSource></pivotCacheDefinition>',
      },
    ],
    workbookPivotCachesXml: '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/></pivotCaches>',
    workbookRelationships: [
      {
        id: 'rIdPivotCache1',
        type: `${officeRelationshipTypePrefix}/pivotCacheDefinition`,
        target: 'pivotCache/pivotCacheDefinition1.xml',
      },
    ],
  },
  chartArtifacts: {
    parts: [
      encodedPart('xl/charts/chart1.xml', '<c:chartSpace><c:chart/></c:chartSpace>'),
      encodedPart('xl/chartsheets/sheet2.xml', '<chartsheet><sheetViews/></chartsheet>'),
    ],
    contentTypeOverrides: [
      {
        partName: '/xl/chartsheets/sheet2.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
      },
    ],
  },
  chartSheetArtifacts: [{ name: 'Chart 1', relationshipTarget: 'chartsheets/sheet2.xml', sheetId: 2, state: 'hidden' }],
  dataModelArtifacts: {
    parts: [encodedPart('xl/model/item.data', 'model-bytes')],
    workbookRelationships: [{ id: 'rIdModel', type: `${officeRelationshipTypePrefix}/model`, target: 'model/item.data' }],
    contentTypeOverrides: [{ partName: '/xl/model/item.data', contentType: 'application/vnd.ms-excel.model' }],
  },
  slicerConnectionArtifacts: {
    parts: [encodedPart('xl/slicerCaches/slicerCache1.xml', '<slicerCache/>')],
    workbookSlicerCachesExtXml: '<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}"/></extLst>',
    workbookRelationships: [
      { id: 'rIdSlicer', type: `${officeRelationshipTypePrefix}/slicerCache`, target: 'slicerCaches/slicerCache1.xml' },
    ],
    sheetArtifacts: [{ sheetName: 'Data', sheetSlicerListExtXml: '<extLst><ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"/></extLst>' }],
  },
  viewState: {
    bookViewsXml: '<bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>',
  },
  styleArtifacts: {
    stylesXml: '<styleSheet><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>',
    theme: {
      path: 'xl/theme/theme1.xml',
      xml: '<a:theme name="Office"/>',
      relationship: { id: 'rIdTheme', type: `${officeRelationshipTypePrefix}/theme`, target: 'theme/theme1.xml' },
      contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
    },
  },
} satisfies WorkbookMetadataSnapshot

const preservedSheetMetadata = {
  styleArtifacts: {
    cellStyleIndexes: [{ address: 'A1', styleIndex: 1 }],
    blankCellAddresses: ['C3'],
  },
  pivotArtifacts: {
    relationships: [{ id: 'rIdPivotTable1', type: `${officeRelationshipTypePrefix}/pivotTable`, target: '../pivotTables/pivotTable1.xml' }],
    pivotTableDefinitionsXml: '<pivotTableDefinition name="PivotTable1"/>',
  },
  viewState: {
    sheetViewsXml:
      '<sheetViews><sheetView workbookViewId="0" topLeftCell="B2" tabSelected="1"><selection activeCell="C3" sqref="C3 D4:E5"/></sheetView></sheetViews>',
  },
} satisfies NonNullable<WorkbookSnapshot['sheets'][number]['metadata']>

function packageMetadataSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Package metadata preservation',
      metadata: preservedWorkbookMetadata,
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        metadata: preservedSheetMetadata,
        cells: [{ address: 'A1', value: 'source' }],
      },
    ],
  }
}

function rawPivotOnlySnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Raw pivot package only',
      metadata: {
        pivotArtifacts: preservedWorkbookMetadata.pivotArtifacts,
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        metadata: {
          pivotArtifacts: preservedSheetMetadata.pivotArtifacts,
        },
        cells: [],
      },
    ],
  }
}

function encodedPart(path: string, text: string): WorkbookPreservedPackagePartSnapshot {
  const bytes = Buffer.from(text, 'utf8')
  return {
    path,
    storage: 'base64',
    dataBase64: bytes.toString('base64'),
    byteLength: bytes.byteLength,
  }
}

function pivotLocationRef(metadata: WorkbookMetadataSnapshot | undefined, path: string): string | undefined {
  return readXmlAttribute(pivotPartXml(metadata, path), 'ref')
}

function pivotCacheSourceRef(metadata: WorkbookMetadataSnapshot | undefined, path: string): string | undefined {
  return readXmlAttribute(pivotPartXml(metadata, path), 'ref')
}

function pivotPartXml(metadata: WorkbookMetadataSnapshot | undefined, path: string): string {
  return metadata?.pivotArtifacts?.parts.find((part) => part.path === path)?.xml ?? ''
}

function readXmlAttribute(xml: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(xml)?.[1]
}
