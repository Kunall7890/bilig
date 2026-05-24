import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { exportXlsx, externalPivotCachesWarning, unsupportedCellStylesWarning } from '../index.js'
import { tryInspectLargeSimpleXlsxHeadless, tryInspectLargeSimpleXlsxHeadlessAsync } from '../xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const controlRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/control'
const oleObjectRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'
const imageRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const printerSettingsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
const pivotTableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'

describe('large simple XLSX control artifacts', () => {
  it('streams OLE object metadata without inflating worksheet XML', () => {
    const bytes = buildWorkbookWithOleObject()
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'ole-object.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.stats.cellCount).toBe(3)
    expect(imported?.snapshot.workbook.metadata?.controlArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/drawings/vmlDrawing1.vml',
      'xl/embeddings/Microsoft_Word_97_-_2003_Document.doc',
      'xl/media/image1.emf',
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.controlArtifacts?.controlsXml).toContain('<oleObjects>')
    expect(
      imported?.snapshot.sheets[0]?.metadata?.controlArtifacts?.relationships.map((relationship) => relationship.type).toSorted(),
    ).toEqual([imageRelationshipType, oleObjectRelationshipType, vmlDrawingRelationshipType])

    const exported = exportXlsx(imported!.snapshot)
    expect(readZipText(exported, 'xl/worksheets/sheet1.xml')).toContain('<legacyDrawing r:id=')
    expect(readZipText(exported, 'xl/worksheets/sheet1.xml')).toContain('<oleObjects>')
    expect(readZipText(exported, 'xl/drawings/vmlDrawing1.vml')).toBe(vmlDrawingXml)
    expect(readZipText(exported, 'xl/embeddings/Microsoft_Word_97_-_2003_Document.doc')).toBe(embeddedDocumentText)
    expect(readZipText(exported, 'xl/media/image1.emf')).toBe(emfImageText)
  })

  it('keeps headless verifier eligibility for OLE metadata worksheets', () => {
    const bytes = buildWorkbookWithOleObject()
    const inspected = tryInspectLargeSimpleXlsxHeadless(bytes, 'headless-ole-object.xlsx', readXlsxZipEntriesLazy(bytes), {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(inspected?.stats.cellCount).toBe(3)
    expect(inspected?.stats.sheetCount).toBe(2)
    expect(inspected?.workbookMetadataKeys).toEqual(expect.arrayContaining(['controlArtifacts', 'drawingArtifacts']))
    expect(inspected?.sheetMetadataKeys).toContain('controlArtifacts')
  })

  it('preserves headless package metadata keys and external pivot warnings', () => {
    const bytes = buildHeadlessMetadataPackageWorkbook()
    const inspected = tryInspectLargeSimpleXlsxHeadless(bytes, 'headless-metadata-package.xlsx', readXlsxZipEntriesLazy(bytes), {
      allowUnsupportedWorksheetFeaturesForMetrics: true,
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(inspected?.warnings).toContain(externalPivotCachesWarning)
    expect(inspected?.stats.warningCount).toBe(1)
    expect(inspected?.workbookMetadataKeys).toEqual(
      expect.arrayContaining(['dataModelArtifacts', 'documentPropertyArtifacts', 'pivotArtifacts', 'properties', 'styles', 'tables']),
    )
  })

  it('preserves headless unsupported style warnings for styled cells', () => {
    const bytes = buildHeadlessUnsupportedStylesWorkbook()
    const imported = tryImportLargeSimpleXlsx(bytes, 'headless-unsupported-styles.xlsx', readXlsxZipEntriesLazy(bytes), {
      minByteLength: 0,
      releaseZipSource: true,
    })
    const inspected = tryInspectLargeSimpleXlsxHeadless(bytes, 'headless-unsupported-styles.xlsx', readXlsxZipEntriesLazy(bytes), {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.warnings).toContain(unsupportedCellStylesWarning)
    expect(inspected?.warnings).toEqual(imported?.warnings)
    expect(inspected?.stats.warningCount).toBe(1)
  })

  it('matches async headless unsupported style warnings for styled cells', async () => {
    const bytes = buildHeadlessUnsupportedStylesWorkbook()
    const inspected = await tryInspectLargeSimpleXlsxHeadlessAsync(
      bytes,
      'headless-unsupported-styles.xlsx',
      readXlsxZipEntriesLazy(bytes),
      {
        minByteLength: 0,
        releaseZipSource: true,
      },
    )

    expect(inspected?.warnings).toContain(unsupportedCellStylesWarning)
    expect(inspected?.stats.warningCount).toBe(1)
    expect(inspected?.stats.cellCount).toBe(1)
  })

  it('lets async headless callers release scanner scratch during large worksheet streaming', async () => {
    const bytes = buildLargeHeadlessWorksheetWorkbook(30_000)
    let releaseCount = 0
    const inspected = await tryInspectLargeSimpleXlsxHeadlessAsync(bytes, 'headless-large-release.xlsx', readXlsxZipEntriesLazy(bytes), {
      afterWorksheetChunkBatch: () => {
        releaseCount += 1
      },
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(inspected?.stats.cellCount).toBe(30_000)
    expect(inspected?.stats.sheetCount).toBe(1)
    expect(releaseCount).toBeGreaterThan(0)
  })

  it('visits headless worksheet relationship and cell metadata without public snapshot materialization', () => {
    const bytes = buildHeadlessWorksheetMetadataVisitorWorkbook()
    const inspected = tryInspectLargeSimpleXlsxHeadless(bytes, 'headless-metadata-visitor.xlsx', readXlsxZipEntriesLazy(bytes), {
      allowUnsupportedWorksheetFeaturesForMetrics: true,
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(inspected?.stats.cellCount).toBe(1)
    expect(inspected?.workbookMetadataKeys).toEqual(expect.arrayContaining(['pivotArtifacts']))
    expect(inspected?.sheetMetadataKeys).toEqual(
      expect.arrayContaining(['cellMetadataRefs', 'pivotArtifacts', 'printPageSetup', 'printerSettings', 'richTextArtifacts']),
    )
  })

  it('streams controls metadata without forcing SheetJS fallback', () => {
    const bytes = buildWorkbookWithControls()
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'controls.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.stats.cellCount).toBe(3)
    expect(imported?.snapshot.workbook.metadata?.controlArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'xl/activeX/activeX1.xml',
      'xl/drawings/vmlDrawing1.vml',
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.controlArtifacts?.controlsXml).toContain('<controls>')
    expect(
      imported?.snapshot.sheets[0]?.metadata?.controlArtifacts?.relationships.map((relationship) => relationship.type).toSorted(),
    ).toEqual([controlRelationshipType, vmlDrawingRelationshipType])

    const exported = exportXlsx(imported!.snapshot)
    expect(readZipText(exported, 'xl/worksheets/sheet1.xml')).toContain('<controls>')
    expect(readZipText(exported, 'xl/activeX/activeX1.xml')).toBe(activeXControlXml)
    expect(readZipText(exported, 'xl/drawings/vmlDrawing1.vml')).toBe(vmlDrawingXml)
  })
})

function buildWorkbookWithOleObject(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(contentTypesXml),
    'xl/workbook.xml': strToU8(workbookXml),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRelationshipsXml),
    'xl/worksheets/sheet1.xml': strToU8(readMeWorksheetXml),
    'xl/worksheets/sheet2.xml': strToU8(dataWorksheetXml),
    'xl/worksheets/_rels/sheet1.xml.rels': strToU8(readMeWorksheetRelationshipsXml),
    'xl/drawings/vmlDrawing1.vml': strToU8(vmlDrawingXml),
    'xl/embeddings/Microsoft_Word_97_-_2003_Document.doc': strToU8(embeddedDocumentText),
    'xl/media/image1.emf': strToU8(emfImageText),
  })
}

function buildWorkbookWithControls(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(contentTypesXml),
    'xl/workbook.xml': strToU8(workbookXml),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRelationshipsXml),
    'xl/worksheets/sheet1.xml': strToU8(controlsWorksheetXml),
    'xl/worksheets/sheet2.xml': strToU8(dataWorksheetXml),
    'xl/worksheets/_rels/sheet1.xml.rels': strToU8(controlsWorksheetRelationshipsXml),
    'xl/drawings/vmlDrawing1.vml': strToU8(vmlDrawingXml),
    'xl/activeX/activeX1.xml': strToU8(activeXControlXml),
  })
}

function buildHeadlessUnsupportedStylesWorkbook(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '</Types>',
      ].join(''),
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rIdSheet1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}"><Relationship Id="rIdSheet1" Type="${worksheetRelationshipType}" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>',
    ),
    'xl/styles.xml': strToU8(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf/><xf applyBorder="1" borderId="1"/></cellXfs></styleSheet>',
    ),
  })
}

function buildHeadlessMetadataPackageWorkbook(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(contentTypesXml),
    'docProps/custom.xml': strToU8('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>'),
    'xl/workbook.xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Data" sheetId="1" r:id="rIdSheet1"/></sheets>',
        '<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache1"/></pivotCaches>',
        '</workbook>',
      ].join(''),
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      [
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}">`,
        `<Relationship Id="rIdSheet1" Type="${worksheetRelationshipType}" Target="worksheets/sheet1.xml"/>`,
        '<Relationship Id="rIdPivotCache1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>',
        '<Relationship Id="rIdQuery1" Type="http://schemas.microsoft.com/office/2014/relationships/query" Target="queries/query1.xml"/>',
        '</Relationships>',
      ].join(''),
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts></worksheet>',
    ),
    'xl/pivotCache/pivotCacheDefinition1.xml': strToU8(
      '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="external"/></pivotCacheDefinition>',
    ),
    'xl/styles.xml': strToU8('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'),
    'xl/tables/table1.xml': strToU8(
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" ref="A1:A1"/>',
    ),
    'xl/queries/query1.xml': strToU8('<query xmlns="http://schemas.microsoft.com/office/2014/queries" name="RevenueQuery"/>'),
  })
}

function buildLargeHeadlessWorksheetWorkbook(rowCount: number): Uint8Array {
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '</Types>',
      ].join(''),
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rIdSheet1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}"><Relationship Id="rIdSheet1" Type="${worksheetRelationshipType}" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`,
    ),
  })
}

function buildHeadlessWorksheetMetadataVisitorWorkbook(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(contentTypesXml),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rIdSheet1"/></sheets></workbook>',
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}"><Relationship Id="rIdSheet1" Type="${worksheetRelationshipType}" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetData><row r="1"><c r="A1" cm="1" t="s"><v>0</v></c></row></sheetData>',
        '<pageSetup r:id="rIdPrinterSettings1"/>',
        '</worksheet>',
      ].join(''),
    ),
    'xl/sharedStrings.xml': strToU8(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>Rich</t></r></si></sst>',
    ),
    'xl/worksheets/_rels/sheet1.xml.rels': strToU8(
      [
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}">`,
        `<Relationship Id="rIdPrinterSettings1" Type="${printerSettingsRelationshipType}" Target="../printerSettings/printerSettings1.bin"/>`,
        `<Relationship Id="rIdPivotTable1" Type="${pivotTableRelationshipType}" Target="../pivotTables/pivotTable1.xml"/>`,
        '</Relationships>',
      ].join(''),
    ),
    'xl/printerSettings/printerSettings1.bin': Uint8Array.from([1, 2, 3, 4]),
    'xl/pivotTables/pivotTable1.xml': strToU8(
      '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1"/>',
    ),
  })
}

function readZipText(bytes: Uint8Array, path: string): string {
  const part = unzipSync(bytes)[path]
  if (!part) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(part)
}

const contentTypesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
  '<Default Extension="xml" ContentType="application/xml"/>',
  '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
  '<Default Extension="emf" ContentType="image/x-emf"/>',
  '<Default Extension="doc" ContentType="application/msword"/>',
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
  '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
  '</Types>',
].join('')

const workbookXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  '<sheets>',
  '<sheet name="Read Me" sheetId="1" r:id="rId1"/>',
  '<sheet name="Data" sheetId="2" r:id="rId2"/>',
  '</sheets></workbook>',
].join('')

const workbookRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rId1" Type="${worksheetRelationshipType}" Target="worksheets/sheet1.xml"/>`,
  `<Relationship Id="rId2" Type="${worksheetRelationshipType}" Target="worksheets/sheet2.xml"/>`,
  '</Relationships>',
].join('')

const readMeWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
  'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14">',
  '<dimension ref="A1:A1"/>',
  '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Read embedded instructions</t></is></c></row></sheetData>',
  '<legacyDrawing r:id="rId3"/>',
  '<oleObjects><mc:AlternateContent><mc:Choice Requires="x14">',
  '<oleObject progId="Document" shapeId="1025" r:id="rId4">',
  '<objectPr defaultSize="0" r:id="rId5">',
  '<anchor><xdr:from><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:row>6</xdr:row></xdr:to></anchor>',
  '</objectPr></oleObject>',
  '</mc:Choice></mc:AlternateContent></oleObjects>',
  '</worksheet>',
].join('')

const dataWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<dimension ref="A1:B1"/>',
  '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>',
  '</worksheet>',
].join('')

const readMeWorksheetRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rId3" Type="${vmlDrawingRelationshipType}" Target="../drawings/vmlDrawing1.vml"/>`,
  `<Relationship Id="rId4" Type="${oleObjectRelationshipType}" Target="../embeddings/Microsoft_Word_97_-_2003_Document.doc"/>`,
  `<Relationship Id="rId5" Type="${imageRelationshipType}" Target="../media/image1.emf"/>`,
  '</Relationships>',
].join('')

const vmlDrawingXml = '<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="_x0000_s1025"/></xml>'
const controlsWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  '<dimension ref="A1"/>',
  '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Run control</t></is></c></row></sheetData>',
  '<legacyDrawing r:id="rId3"/>',
  '<controls><control shapeId="1025" name="Button 1" r:id="rId4"/></controls>',
  '</worksheet>',
].join('')
const controlsWorksheetRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rId3" Type="${vmlDrawingRelationshipType}" Target="../drawings/vmlDrawing1.vml"/>`,
  `<Relationship Id="rId4" Type="${controlRelationshipType}" Target="../activeX/activeX1.xml"/>`,
  '</Relationships>',
].join('')
const activeXControlXml =
  '<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" classid="{00000000-0000-0000-0000-000000000000}"/>'
const embeddedDocumentText = 'embedded-doc-fixture'
const emfImageText = 'emf-image-fixture'
