import { describe, expect, it } from 'vitest'

import { internLargeSimpleSheetMetadataInput, internLargeSimpleWorksheetMetadata } from '../xlsx-large-simple-metadata-interning.js'
import type { LargeSimpleSheetMetadataInput } from '../xlsx-large-simple-import-types.js'
import { ImportedWorkbookStringPool } from '../xlsx-large-simple-string-pool.js'
import type { LargeSimpleWorksheetScannedMetadata } from '../xlsx-large-simple-worksheet-metadata.js'

describe('large simple XLSX metadata interning', () => {
  it('pools streamed metadata strings without dropping fidelity-only fields', () => {
    const pool = new ImportedWorkbookStringPool()
    const metadata: LargeSimpleWorksheetScannedMetadata = {
      cellMetadataRefs: [
        { address: 'A1', cm: '1', vm: '1' },
        { address: 'A1', cm: '1' },
      ],
      columns: {
        entries: [{ id: 'col:0', index: 0, size: 64 }],
        metadata: [{ start: 0, count: 1, size: 64 }],
      },
      conditionalFormats: [
        {
          id: 'xlsx-cf:Data:A1:B2:1',
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' },
          rule: { kind: 'formula', formula: '=A1>0' },
          style: {},
          priority: 1,
        },
      ],
      conditionalFormattingXml: ['<conditionalFormatting sqref="A1:B2"/>'],
      controlArtifacts: {
        controlsXml: '<oleObjects><oleObject r:id="rIdControl"/></oleObjects>',
        worksheetRootOpenTag: '<worksheet xmlns:r="relationships">',
        legacyDrawingRelationshipId: 'rIdLegacy',
      },
      dataValidations: [
        {
          range: { sheetName: 'Data', startAddress: 'C1', endAddress: 'C4' },
          rule: { kind: 'list', source: { kind: 'range-ref', sheetName: 'Model', startAddress: 'A1', endAddress: 'A4' } },
          promptTitle: 'Status',
          promptMessage: 'Pick a value',
        },
      ],
      drawingRelationshipId: 'rIdDrawing',
      filters: [
        {
          sheetName: 'Data',
          startAddress: 'A1',
          endAddress: 'B2',
          criteria: [{ colId: 0, filters: { values: ['Open'] }, customFilters: { filters: [{ value: 'Closed' }] } }],
        },
      ],
      hyperlinks: [
        {
          ref: 'A1',
          relationshipId: 'rIdHyperlink',
          location: '#Data!A1',
          tooltip: 'Open row',
          display: 'Open row',
        },
      ],
      legacyDrawingRelationshipId: 'rIdLegacy',
      merges: [{ startAddress: 'A1', endAddress: 'B2' }],
      printPageSetup: {
        printOptionsXml: '<printOptions horizontalCentered="1"/>',
        pageMarginsXml: '<pageMargins left="0.7" right="0.7"/>',
      },
      rows: {
        entries: [{ id: 'row:0', index: 0, size: 20 }],
        metadata: [{ start: 0, count: 1, size: 20 }],
      },
      sheetFormatPr: { defaultRowHeight: 15 },
      sheetSlicerListExtXml: '<ext><x14:slicerList/></ext>',
      tableRelationshipIds: ['rIdTable'],
    }
    const original = structuredClone(metadata)

    const interned = internLargeSimpleWorksheetMetadata(metadata, pool)

    expect(interned).toBe(metadata)
    expect(interned?.columns).toBe(metadata.columns)
    expect(interned?.columns?.entries).toBe(metadata.columns?.entries)
    expect(interned?.columns?.metadata).toBe(metadata.columns?.metadata)
    expect(interned?.conditionalFormats).toBe(metadata.conditionalFormats)
    expect(interned?.dataValidations).toBe(metadata.dataValidations)
    expect(interned?.filters).toBe(metadata.filters)
    expect(interned?.hyperlinks).toBe(metadata.hyperlinks)
    expect(interned?.merges).toBe(metadata.merges)
    expect(interned?.rows).toBe(metadata.rows)
    expect(interned?.rows?.entries).toBe(metadata.rows?.entries)
    expect(interned?.rows?.metadata).toBe(metadata.rows?.metadata)
    expect(interned).toEqual(original)
    expect(interned?.cellMetadataRefs).toEqual(original.cellMetadataRefs)
    expect(interned?.controlArtifacts).toEqual(original.controlArtifacts)
    expect(interned?.legacyDrawingRelationshipId).toBe('rIdLegacy')
    expect(interned?.sheetSlicerListExtXml).toBe('<ext><x14:slicerList/></ext>')
    expect(pool.count).toBeLessThan(30)
  })

  it('pools resolved metadata input strings assembled after worksheet XML parsing', () => {
    const pool = new ImportedWorkbookStringPool()
    const input: LargeSimpleSheetMetadataInput = {
      conditionalFormatArtifacts: { xml: '<conditionalFormatting sqref="A1"><cfRule type="dataBar"/></conditionalFormatting>' },
      conditionalFormats: [
        {
          id: 'xlsx-cf:Data:A1:A1:1',
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A1' },
          rule: { kind: 'formula', formula: '=Data!A1>0' },
          style: {},
          priority: 1,
        },
      ],
      controlArtifacts: {
        controlsXml: '<controls><control r:id="rIdShared"/></controls>',
        worksheetRootOpenTag: '<worksheet xmlns:r="relationships">',
        relationships: [{ id: 'rIdShared', type: 'control', target: '../ctrlProp/control1.xml' }],
      },
      validations: [
        {
          range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A1' },
          rule: { kind: 'list', source: { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'A1' } },
          promptTitle: 'Open row',
          promptMessage: 'Open row',
        },
      ],
      drawingArtifacts: {
        relationshipTarget: '../drawings/drawing1.xml',
        preservedChartRelationshipIds: ['rIdShared'],
      },
      filters: [
        {
          sheetName: 'Data',
          startAddress: 'A1',
          endAddress: 'A1',
          criteria: [{ colId: 0, filters: { values: ['Open row'] }, customFilters: { filters: [{ value: 'Open row' }] } }],
        },
      ],
      hyperlinks: [
        {
          sheetName: 'Data',
          address: 'A1',
          target: '#Data!A1',
          tooltip: 'Open row',
          display: 'Open row',
        },
      ],
      legacyCommentVml: {
        relationshipTarget: '../drawings/vmlDrawing1.vml',
        vmlXml: '<xml><v:shape id="comment"/></xml>',
        commentsRelationshipTarget: '../comments1.xml',
        commentsXml: '<comments/>',
        commentSignature: 'comment-signature',
      },
      pivotArtifacts: {
        relationships: [{ id: 'rIdShared', type: 'pivotTable', target: '../pivotTables/pivotTable1.xml' }],
        pivotTableDefinitionsXml: '<pivotTableDefinition name="Pivot1"/>',
      },
      printerSettings: [
        {
          relationshipTarget: '../printerSettings/printerSettings1.bin',
          storage: 'base64',
          dataBase64: 'AQIDBA==',
          byteLength: 4,
          pageSetupXml: '<pageSetup r:id="rIdShared"/>',
        },
      ],
      printPageSetup: {
        printOptionsXml: '<printOptions horizontalCentered="1"/>',
        pageMarginsXml: '<pageMargins left="0.7" right="0.7"/>',
        pageSetupXml: '<pageSetup r:id="rIdShared"/>',
      },
      sheetProtection: {
        sheetName: 'Data',
        xmlAttributes: [{ name: 'algorithmName', value: 'SHA-512' }],
      },
    }
    const original = structuredClone(input)

    const interned = internLargeSimpleSheetMetadataInput(input, pool)

    expect(interned).toBe(input)
    expect(interned).toEqual(original)
    expect(interned.controlArtifacts).toBe(input.controlArtifacts)
    expect(interned.validations).toBe(input.validations)
    expect(interned.filters).toBe(input.filters)
    expect(interned.hyperlinks).toBe(input.hyperlinks)
    expect(interned.printerSettings).toBe(input.printerSettings)
    expect(pool.count).toBeLessThan(35)
  })
})
