import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const powerPivotDataRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData'
const customXmlRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
const customDataPropertiesRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/customDataProps'
const customDataRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/customData'
const customXmlPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
const dataModelContentType = 'application/vnd.openxmlformats-officedocument.model+data'
const customDataPropertiesContentType = 'application/vnd.ms-excel.customDataProperties+xml'
const customXmlPropertiesContentType = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'

describe('xlsx data model artifacts roundtrip', () => {
  it('preserves Power Pivot data model parts and supporting custom XML package parts', () => {
    const source = buildWorkbookWithDataModelArtifacts()

    const imported = importXlsx(source, 'power-pivot-model.xlsx')
    const exported = exportXlsx(imported.snapshot)

    expect(imported.snapshot.workbook.metadata?.dataModelArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
      'customXml/_rels/item1.xml.rels',
      'customXml/item1.xml',
      'customXml/itemProps1.xml',
      'xl/customData/_rels/itemProps1.xml.rels',
      'xl/customData/item1.data',
      'xl/customData/itemProps1.xml',
      'xl/model/item.data',
    ])
    expect(imported.snapshot.workbook.metadata?.dataModelArtifacts?.workbookRelationships).toEqual([
      { id: 'rId8', type: powerPivotDataRelationshipType, target: 'model/item.data' },
      { id: 'rId9', type: customXmlRelationshipType, target: '../customXml/item1.xml' },
      { id: 'rId10', type: customDataPropertiesRelationshipType, target: 'customData/itemProps1.xml' },
    ])
    expect(dataModelMetrics(exported)).toEqual(dataModelMetrics(source))
    expect(readZipBytes(exported, 'xl/model/item.data')).toEqual(readZipBytes(source, 'xl/model/item.data'))
    expect(readZipBytes(exported, 'xl/customData/item1.data')).toEqual(readZipBytes(source, 'xl/customData/item1.data'))
    expect(readZipText(exported, 'xl/customData/itemProps1.xml')).toBe(customDataItemPropertiesXml)
    expect(readZipText(exported, 'xl/customData/_rels/itemProps1.xml.rels')).toBe(customDataItemRelationshipsXml)
    expect(readZipText(exported, 'customXml/item1.xml')).toBe(customXmlItemXml)
    expect(readZipText(exported, 'customXml/itemProps1.xml')).toBe(customXmlItemPropertiesXml)
    expect(readZipText(exported, 'customXml/_rels/item1.xml.rels')).toBe(customXmlItemRelationshipsXml)
    expect(readWorkbookRelationship(exported, powerPivotDataRelationshipType)).toMatchObject({
      target: 'model/item.data',
      type: powerPivotDataRelationshipType,
    })
    expect(readWorkbookRelationship(exported, customXmlRelationshipType)).toMatchObject({
      target: '../customXml/item1.xml',
      type: customXmlRelationshipType,
    })
    expect(readWorkbookRelationship(exported, customDataPropertiesRelationshipType)).toMatchObject({
      target: 'customData/itemProps1.xml',
      type: customDataPropertiesRelationshipType,
    })
    expect(readContentTypeDefault(exported, 'data')).toBe(dataModelContentType)
    expect(readContentTypeOverride(exported, '/customXml/itemProps1.xml')).toBe(customXmlPropertiesContentType)
    expect(readContentTypeOverride(exported, '/xl/customData/itemProps1.xml')).toBe(customDataPropertiesContentType)
  })
})

function buildWorkbookWithDataModelArtifacts(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildWorkbook()))
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      [
        `<Relationship Id="rId8" Type="${powerPivotDataRelationshipType}" Target="model/item.data"/>`,
        `<Relationship Id="rId9" Type="${customXmlRelationshipType}" Target="../customXml/item1.xml"/>`,
        `<Relationship Id="rId10" Type="${customDataPropertiesRelationshipType}" Target="customData/itemProps1.xml"/>`,
        '</Relationships>',
      ].join(''),
    ),
  )
  zip['xl/model/item.data'] = new Uint8Array([80, 111, 119, 101, 114, 80, 105, 118, 111, 116, 0, 1, 2, 255])
  zip['xl/customData/item1.data'] = new Uint8Array([67, 117, 115, 116, 111, 109, 68, 97, 116, 97, 0, 1, 2, 255])
  zip['xl/customData/itemProps1.xml'] = strToU8(customDataItemPropertiesXml)
  zip['xl/customData/_rels/itemProps1.xml.rels'] = strToU8(customDataItemRelationshipsXml)
  zip['customXml/item1.xml'] = strToU8(customXmlItemXml)
  zip['customXml/itemProps1.xml'] = strToU8(customXmlItemPropertiesXml)
  zip['customXml/_rels/item1.xml.rels'] = strToU8(customXmlItemRelationshipsXml)
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(
      upsertContentTypeOverride(addContentTypeDefault(readZipTextFromZip(zip, '[Content_Types].xml'), 'data', dataModelContentType), {
        contentType: customXmlPropertiesContentType,
        partName: '/customXml/itemProps1.xml',
      }),
      {
        contentType: customDataPropertiesContentType,
        partName: '/xl/customData/itemProps1.xml',
      },
    ),
  )
  return zipSync(zip)
}

function buildWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Data model artifacts' },
    sheets: [
      {
        id: 1,
        name: 'Model',
        order: 0,
        cells: [{ address: 'A1', value: 'Power Pivot fixture' }],
      },
    ],
  }
}

function dataModelMetrics(bytes: Uint8Array): {
  customDataPackageParts: number
  customXmlPackageParts: number
  dataModelPackageParts: number
  workbookCustomDataRelationships: number
  workbookCustomXmlRelationships: number
  workbookDataModelRelationships: number
} {
  const zip = unzipSync(bytes)
  const workbookRelationshipsXml = readZipTextFromZip(zip, 'xl/_rels/workbook.xml.rels')
  return {
    customDataPackageParts: Object.keys(zip).filter((path) => path.startsWith('xl/customData/')).length,
    customXmlPackageParts: Object.keys(zip).filter((path) => path.startsWith('customXml/')).length,
    dataModelPackageParts: Object.keys(zip).filter((path) => path.startsWith('xl/model/')).length,
    workbookCustomDataRelationships: relationshipsWithType(workbookRelationshipsXml, customDataPropertiesRelationshipType).length,
    workbookCustomXmlRelationships: relationshipsWithType(workbookRelationshipsXml, customXmlRelationshipType).length,
    workbookDataModelRelationships: relationshipsWithType(workbookRelationshipsXml, powerPivotDataRelationshipType).length,
  }
}

function relationshipsWithType(relationshipsXml: string, relationshipType: string): string[] {
  return [...relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    return readXmlAttribute(attributes, 'Type') === relationshipType ? [match[0]] : []
  })
}

function readZipBytes(bytes: Uint8Array, path: string): Uint8Array {
  const part = unzipSync(bytes)[path]
  if (!part) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return part
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

function readContentTypeDefault(bytes: Uint8Array, extension: string): string | undefined {
  const contentTypesXml = readZipText(bytes, '[Content_Types].xml')
  for (const match of contentTypesXml.matchAll(/<Default\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'Extension') === extension) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
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

function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  if (new RegExp(`<Default\\b[^>]*\\bExtension=(["'])${extension}\\1`, 'u').test(contentTypesXml)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`)
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

const customXmlItemXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<DataModelingSandbox.SerializedSandboxErrorCache xmlns="http://schemas.datacontract.org/2004/07/Microsoft.AnalysisServices">',
  '<CustomContent>Data model state</CustomContent>',
  '</DataModelingSandbox.SerializedSandboxErrorCache>',
].join('')

const customXmlItemPropertiesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="no"?>',
  '<ds:datastoreItem ds:itemID="{83A22A58-0005-4759-8BCC-9F38ABACA316}" ',
  'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>',
].join('')

const customXmlItemRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rId1" Type="${customXmlPropertiesRelationshipType}" Target="itemProps1.xml"/>`,
  '</Relationships>',
].join('')

const customDataItemPropertiesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<customDataProperties xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">',
  '<customDataProperties>',
  '<customDataProperty name="ModelData" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>',
  '</customDataProperties>',
  '</customDataProperties>',
].join('')

const customDataItemRelationshipsXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<Relationships xmlns="${relationshipNamespace}">`,
  `<Relationship Id="rId1" Type="${customDataRelationshipType}" Target="item1.data"/>`,
  '</Relationships>',
].join('')
