import { Buffer } from 'node:buffer'

import { describe, expect, it } from 'vitest'

import type { WorkbookMetadataSnapshot, WorkbookPreservedPackagePartSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const officeRelationshipTypePrefix = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const microsoftOfficeRelationshipTypePrefix = 'http://schemas.microsoft.com/office/2007/relationships'

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

  it('structurally rewrites preserved chart package formula refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-chart-artifact-structural-rewrite' })
    await engine.ready()

    engine.importSnapshot(packageMetadataSnapshot())
    engine.insertRows('Data', 1, 1)

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(chartFormulaRefs(metadata, 'xl/charts/chart1.xml')).toEqual([
      'Data!$B$1',
      'Data!$A$3:$A$4',
      'Data!$B$3:$B$4',
      'SUM(Data!$B$3:$B$4,Other!$B$2:$B$3)',
    ])
  })

  it('structurally rewrites quoted sheet refs in preserved chart package formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-chart-artifact-quoted-structural-rewrite' })
    await engine.ready()

    engine.importSnapshot(rawChartOnlySnapshot("Owner's Data", quotedChartFormulaXml()))
    engine.insertRows("Owner's Data", 1, 1)

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(chartFormulaRefs(metadata, 'xl/charts/chart1.xml')).toEqual(["'Owner''s Data'!$B$3:$B$4"])
  })

  it('renames preserved chart package formula sheet refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-chart-artifact-sheet-rename' })
    await engine.ready()

    engine.importSnapshot(packageMetadataSnapshot())
    engine.renameSheet('Data', 'Revenue Data')

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(chartFormulaRefs(metadata, 'xl/charts/chart1.xml')).toEqual([
      "'Revenue Data'!$B$1",
      "'Revenue Data'!$A$2:$A$3",
      "'Revenue Data'!$B$2:$B$3",
      "SUM('Revenue Data'!$B$2:$B$3,Other!$B$2:$B$3)",
    ])
  })

  it('renames preserved pivot package and workbook metadata sheet refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-pivot-artifact-sheet-rename' })
    await engine.ready()

    engine.importSnapshot(packageMetadataSnapshot())
    engine.renameSheet('Data', 'Revenue Data')

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(pivotCacheSourceSheet(metadata, 'xl/pivotCache/pivotCacheDefinition1.xml')).toBe('Revenue Data')
    expect(metadata?.unsupportedFormulaDependencies?.map((entry) => entry.sheetName)).toEqual(['Revenue Data'])
    expect(metadata?.unsupportedPivots?.map((entry) => entry.sheetName)).toEqual(['Revenue Data'])
    expect(metadata?.formulaAudit?.formulas.map((entry) => entry.sheetName)).toEqual(['Revenue Data'])
    expect(metadata?.formulaAudit?.calcChain?.cells.map((entry) => entry.sheetName)).toEqual(['Revenue Data'])
    expect(metadata?.slicerConnectionArtifacts?.sheetArtifacts?.map((entry) => entry.sheetName)).toEqual(['Revenue Data'])
  })

  it('prunes preserved slicer sheet topology when deleting its owning sheet', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-slicer-artifact-sheet-delete' })
    await engine.ready()

    engine.importSnapshot(slicerConnectionSheetDeletionSnapshot())
    engine.deleteSheet('Revenue')

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => sheet.name)).toEqual(['Keep'])
    expect(exported.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts).toBeUndefined()
    expect(exported.workbook.metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/slicerCaches/_rels/slicerCache1.xml.rels',
      'xl/slicerCaches/slicerCache1.xml',
    ])
    expect(exported.workbook.metadata?.slicerConnectionArtifacts?.workbookRelationships).toEqual([
      {
        id: 'rIdSlicerCache',
        type: `${microsoftOfficeRelationshipTypePrefix}/slicerCache`,
        target: 'slicerCaches/slicerCache1.xml',
      },
    ])
    expect(exported.workbook.metadata?.slicerConnectionArtifacts?.contentTypeOverrides).toEqual([
      {
        partName: '/xl/slicerCaches/slicerCache1.xml',
        contentType: 'application/vnd.ms-excel.slicerCache+xml',
      },
    ])
  })

  it('keeps shared preserved slicer parts when deleting one referencing sheet', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-shared-slicer-artifact-sheet-delete' })
    await engine.ready()

    engine.importSnapshot(sharedSlicerConnectionSheetDeletionSnapshot())
    engine.deleteSheet('Revenue')

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => sheet.name)).toEqual(['Keep'])
    expect(exported.workbook.metadata?.slicerConnectionArtifacts?.sheetArtifacts).toEqual([
      {
        sheetName: 'Keep',
        sheetSlicerListExtXml:
          '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"><x14:slicerList><x14:slicer r:id="rIdKeepSlicer"/></x14:slicerList></ext>',
        relationships: [
          {
            id: 'rIdKeepSlicer',
            type: `${microsoftOfficeRelationshipTypePrefix}/slicer`,
            target: '../slicers/slicer1.xml',
          },
        ],
      },
    ])
    expect(exported.workbook.metadata?.slicerConnectionArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/slicerCaches/_rels/slicerCache1.xml.rels',
      'xl/slicerCaches/slicerCache1.xml',
      'xl/slicers/_rels/slicer1.xml.rels',
      'xl/slicers/slicer1.xml',
    ])
    expect(exported.workbook.metadata?.slicerConnectionArtifacts?.contentTypeOverrides).toEqual([
      {
        partName: '/xl/slicerCaches/slicerCache1.xml',
        contentType: 'application/vnd.ms-excel.slicerCache+xml',
      },
      {
        partName: '/xl/slicers/slicer1.xml',
        contentType: 'application/vnd.ms-excel.slicer+xml',
      },
    ])
  })

  it('rewrites preserved workbook view tab indexes after sheet deletion', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-workbook-view-delete-sheet' })
    await engine.ready()

    engine.importSnapshot(workbookViewStateSheetDeletionSnapshot())
    engine.deleteSheet('Data')

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Report'])
    expect(exported.workbook.metadata?.viewState).toEqual({
      bookViewsXml: '<bookViews><workbookView activeTab="1"/></bookViews>',
    })
  })

  it('rewrites preserved workbook view tab indexes after sheet reorder', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-workbook-view-reorder-sheet' })
    await engine.ready()

    engine.importSnapshot(workbookViewStateSheetReorderSnapshot())
    engine.moveSheet('Report', 1)

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => `${sheet.order}:${sheet.name}`)).toEqual(['0:Data', '1:Report', '2:Inputs'])
    expect(exported.workbook.metadata?.viewState).toEqual({
      bookViewsXml: '<bookViews><workbookView activeTab="1" firstSheet="1"/></bookViews>',
    })
  })

  it('drops deleted-sheet preserved calc-chain cells without reindexing sheet ids', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-calc-chain-delete-sheet' })
    await engine.ready()

    engine.importSnapshot(calcChainSheetDeletionSnapshot())
    engine.deleteSheet('Data')

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(metadata?.formulaAudit?.formulas.map((entry) => `${entry.sheetName}:${entry.address}`)).toEqual(['Inputs:A1', 'Report:A1'])
    expect(metadata?.formulaAudit?.calcChain?.cells).toEqual([
      { sheetIndex: 2, sheetName: 'Inputs', address: 'A1' },
      { sheetIndex: 3, sheetName: 'Report', address: 'A1' },
    ])
  })

  it('preserves calc-chain sheet ids after sheet reorder', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-calc-chain-reorder-sheet' })
    await engine.ready()

    engine.importSnapshot(calcChainSheetDeletionSnapshot())
    engine.moveSheet('Report', 1)

    const exported = engine.exportSnapshot()
    expect(exported.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report', 'Inputs'])
    expect(exported.workbook.metadata?.formulaAudit?.calcChain?.cells).toEqual([
      { sheetIndex: 1, sheetName: 'Data', address: 'A1' },
      { sheetIndex: 2, sheetName: 'Inputs', address: 'A1' },
      { sheetIndex: 3, sheetName: 'Report', address: 'A1' },
    ])
  })

  it('renames quoted sheet refs in preserved chart package formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-chart-artifact-quoted-sheet-rename' })
    await engine.ready()

    engine.importSnapshot(rawChartOnlySnapshot("Owner's Data", quotedChartFormulaXml()))
    engine.renameSheet("Owner's Data", 'Renamed Data')

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(chartFormulaRefs(metadata, 'xl/charts/chart1.xml')).toEqual(["'Renamed Data'!$B$2:$B$3"])
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

  it('treats preserved chart package formulas as structural delete impact even without cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'package-metadata-chart-artifact-delete-impact' })
    await engine.ready()

    engine.importSnapshot(rawChartOnlySnapshot('Data', chartFormulaXml()))
    engine.deleteRows('Data', 0, 1)

    const metadata = engine.exportSnapshot().workbook.metadata
    expect(chartFormulaRefs(metadata, 'xl/charts/chart1.xml')).toEqual([
      '#REF!',
      'Data!$A$1:$A$2',
      'Data!$B$1:$B$2',
      'SUM(Data!$B$1:$B$2,Other!$B$2:$B$3)',
    ])
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
      encodedPart('xl/charts/chart1.xml', chartFormulaXml()),
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

function workbookViewStateSheetDeletionSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Workbook view sheet deletion',
      metadata: {
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>',
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [{ address: 'A1', value: 'data' }],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [{ address: 'A1', value: 'inputs' }],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        cells: [{ address: 'A1', value: 'report' }],
      },
    ],
  }
}

function slicerConnectionSheetDeletionSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Slicer connection sheet deletion',
      metadata: {
        slicerConnectionArtifacts: {
          parts: [
            encodedPart('xl/slicerCaches/slicerCache1.xml', '<slicerCacheDefinition/>'),
            encodedPart('xl/slicerCaches/_rels/slicerCache1.xml.rels', '<Relationships/>'),
            encodedPart('xl/slicers/slicer1.xml', '<slicer/>'),
            encodedPart('xl/slicers/_rels/slicer1.xml.rels', '<Relationships/>'),
          ],
          workbookSlicerCachesExtXml:
            '<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}"><x15:slicerCaches><x15:slicerCache r:id="rIdSlicerCache"/></x15:slicerCaches></ext>',
          workbookRelationships: [
            {
              id: 'rIdSlicerCache',
              type: `${microsoftOfficeRelationshipTypePrefix}/slicerCache`,
              target: 'slicerCaches/slicerCache1.xml',
            },
          ],
          sheetArtifacts: [
            {
              sheetName: 'Revenue',
              sheetSlicerListExtXml:
                '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"><x14:slicerList><x14:slicer r:id="rIdSlicer"/></x14:slicerList></ext>',
              relationships: [
                {
                  id: 'rIdSlicer',
                  type: `${microsoftOfficeRelationshipTypePrefix}/slicer`,
                  target: '../slicers/slicer1.xml',
                },
              ],
            },
          ],
          contentTypeOverrides: [
            {
              partName: '/xl/slicerCaches/slicerCache1.xml',
              contentType: 'application/vnd.ms-excel.slicerCache+xml',
            },
            {
              partName: '/xl/slicers/slicer1.xml',
              contentType: 'application/vnd.ms-excel.slicer+xml',
            },
          ],
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Revenue',
        order: 0,
        cells: [{ address: 'A1', value: 'revenue' }],
      },
      {
        id: 2,
        name: 'Keep',
        order: 1,
        cells: [{ address: 'A1', value: 'keep' }],
      },
    ],
  }
}

function sharedSlicerConnectionSheetDeletionSnapshot(): WorkbookSnapshot {
  const snapshot = slicerConnectionSheetDeletionSnapshot()
  const artifacts = snapshot.workbook.metadata?.slicerConnectionArtifacts
  if (!artifacts) {
    throw new Error('Expected slicer connection artifacts')
  }
  artifacts.sheetArtifacts = [
    ...(artifacts.sheetArtifacts ?? []),
    {
      sheetName: 'Keep',
      sheetSlicerListExtXml:
        '<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"><x14:slicerList><x14:slicer r:id="rIdKeepSlicer"/></x14:slicerList></ext>',
      relationships: [
        {
          id: 'rIdKeepSlicer',
          type: `${microsoftOfficeRelationshipTypePrefix}/slicer`,
          target: '../slicers/slicer1.xml',
        },
      ],
    },
  ]
  return snapshot
}

function workbookViewStateSheetReorderSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Workbook view sheet reorder',
      metadata: {
        viewState: {
          bookViewsXml: '<bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>',
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [{ address: 'A1', value: 'data' }],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [{ address: 'A1', value: 'inputs' }],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        cells: [{ address: 'A1', value: 'report' }],
      },
    ],
  }
}

function calcChainSheetDeletionSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Calc chain sheet deletion',
      metadata: {
        formulaAudit: {
          formulas: [
            {
              context: 'worksheet-cell',
              clause: '18.3.1.40',
              formula: '1+1',
              sheetName: 'Data',
              address: 'A1',
              cacheStatus: 'trustedCached',
            },
            {
              context: 'worksheet-cell',
              clause: '18.3.1.40',
              formula: '10+1',
              sheetName: 'Inputs',
              address: 'A1',
              cacheStatus: 'trustedCached',
            },
            {
              context: 'worksheet-cell',
              clause: '18.3.1.40',
              formula: 'Inputs!A1+1',
              sheetName: 'Report',
              address: 'A1',
              cacheStatus: 'trustedCached',
            },
          ],
          calcChain: {
            packagePath: 'xl/calcChain.xml',
            cells: [
              { sheetIndex: 1, sheetName: 'Data', address: 'A1' },
              { sheetIndex: 2, sheetName: 'Inputs', address: 'A1' },
              { sheetIndex: 3, sheetName: 'Report', address: 'A1' },
            ],
          },
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [{ address: 'A1', formula: '1+1', value: 2 }],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [{ address: 'A1', formula: '10+1', value: 11 }],
      },
      {
        id: 3,
        name: 'Report',
        order: 2,
        cells: [{ address: 'A1', formula: 'Inputs!A1+1', value: 12 }],
      },
    ],
  }
}

function rawChartOnlySnapshot(sheetName: string, chartXml: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Raw chart package only',
      metadata: {
        chartArtifacts: {
          parts: [encodedPart('xl/charts/chart1.xml', chartXml)],
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: sheetName,
        order: 0,
        cells: [],
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

function pivotCacheSourceSheet(metadata: WorkbookMetadataSnapshot | undefined, path: string): string | undefined {
  return readXmlAttribute(pivotPartXml(metadata, path), 'sheet')
}

function pivotPartXml(metadata: WorkbookMetadataSnapshot | undefined, path: string): string {
  return metadata?.pivotArtifacts?.parts.find((part) => part.path === path)?.xml ?? ''
}

function chartFormulaRefs(metadata: WorkbookMetadataSnapshot | undefined, path: string): string[] {
  const part = metadata?.chartArtifacts?.parts.find((candidate) => candidate.path === path)
  if (!part) {
    return []
  }
  const xml = Buffer.from(part.dataBase64, 'base64').toString('utf8')
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/gu)].map((match) =>
    decodeXmlText(match[1] ?? ''),
  )
}

function readXmlAttribute(xml: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(xml)?.[1]
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function chartFormulaXml(): string {
  return [
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">',
    '<c:chart><c:plotArea><c:lineChart><c:ser>',
    '<c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>',
    '<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f></c:strRef></c:cat>',
    '<c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f></c:numRef></c:val>',
    '<c:extLst><c:ext><c:f>SUM(Data!$B$2:$B$3,Other!$B$2:$B$3)</c:f></c:ext></c:extLst>',
    '</c:ser></c:lineChart></c:plotArea></c:chart>',
    '</c:chartSpace>',
  ].join('')
}

function quotedChartFormulaXml(): string {
  return [
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">',
    '<c:chart><c:plotArea><c:lineChart><c:ser>',
    '<c:val><c:numRef><c:f>&apos;Owner&apos;&apos;s Data&apos;!$B$2:$B$3</c:f></c:numRef></c:val>',
    '</c:ser></c:lineChart></c:plotArea></c:chart>',
    '</c:chartSpace>',
  ].join('')
}
