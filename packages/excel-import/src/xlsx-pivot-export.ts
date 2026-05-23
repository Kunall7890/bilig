import { unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { CellRangeRef, LiteralInput, WorkbookPivotSnapshot, WorkbookPivotValueSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import {
  addContentTypeOverride,
  addExportPreservedPivotArtifactsToXlsxBytes,
  buildRelationshipsXml,
  ensureRelationshipNamespace,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  pivotCacheDefinitionContentType,
  pivotCacheDefinitionRelationshipType,
  pivotTableContentType,
  pivotTableRelationshipType,
  setZipText,
  spreadsheetNamespace,
} from './xlsx-pivot-artifacts.js'
import { defaultDataFieldVerb, subtotalValue } from './xlsx-pivot-aggregate-xml.js'
import { getZipText, type XlsxZipEntries } from './xlsx-zip.js'

type ZipEntries = XlsxZipEntries

interface PivotCacheField {
  readonly name: string
  readonly values: readonly LiteralInput[]
}

interface PivotCacheTable {
  readonly fields: readonly PivotCacheField[]
  readonly rows: readonly (readonly LiteralInput[])[]
}

function nextPartIndex(zip: ZipEntries, prefix: string, suffix: string): number {
  let next = 1
  for (const path of Object.keys(zip)) {
    if (!path.startsWith(prefix) || !path.endsWith(suffix)) {
      continue
    }
    const raw = path.slice(prefix.length, -suffix.length)
    const value = Number(raw)
    if (Number.isInteger(value) && value >= next) {
      next = value + 1
    }
  }
  return next
}

function absoluteAddress(address: string): string {
  const decoded = XLSX.utils.decode_cell(address)
  return `$${XLSX.utils.encode_col(decoded.c)}$${decoded.r + 1}`
}

function rangeRefA1(range: CellRangeRef): string {
  const start = absoluteAddress(range.startAddress).replaceAll('$', '')
  const end = absoluteAddress(range.endAddress).replaceAll('$', '')
  return start === end ? start : `${start}:${end}`
}

function pivotOutputRange(pivot: WorkbookPivotSnapshot): string {
  const start = XLSX.utils.decode_cell(pivot.address)
  const end = {
    r: start.r + Math.max(1, pivot.rows) - 1,
    c: start.c + Math.max(1, pivot.cols) - 1,
  }
  return XLSX.utils.encode_range({ s: start, e: end })
}

function expandWorksheetDimension(sheetXml: string, pivot: WorkbookPivotSnapshot): string {
  const match = /<dimension\b[^>]*\bref="([^"]+)"[^>]*\/>/u.exec(sheetXml)
  if (!match) {
    return sheetXml
  }
  try {
    const existing = XLSX.utils.decode_range(match[1] ?? 'A1')
    const next = XLSX.utils.decode_range(pivotOutputRange(pivot))
    const expanded = XLSX.utils.encode_range({
      s: { r: Math.min(existing.s.r, next.s.r), c: Math.min(existing.s.c, next.s.c) },
      e: { r: Math.max(existing.e.r, next.e.r), c: Math.max(existing.e.c, next.e.c) },
    })
    return sheetXml.replace(match[0], `<dimension ref="${expanded}"/>`)
  } catch {
    return sheetXml
  }
}

function addWorksheetPivotTableDefinition(sheetXml: string, relationshipId: string, pivot: WorkbookPivotSnapshot): string {
  const withNamespace = ensureRelationshipNamespace(expandWorksheetDimension(sheetXml, pivot))
  return withNamespace.replace('</worksheet>', `<pivotTableDefinition r:id="${escapeXml(relationshipId)}"/></worksheet>`)
}

function addWorkbookPivotCache(workbookXml: string, cacheId: number, relationshipId: string): string {
  const withNamespace = ensureRelationshipNamespace(workbookXml)
  const entry = `<pivotCache cacheId="${String(cacheId)}" r:id="${escapeXml(relationshipId)}"/>`
  if (/<pivotCaches\b/u.test(withNamespace)) {
    return withNamespace.replace('</pivotCaches>', `${entry}</pivotCaches>`)
  }
  const pivotCaches = `<pivotCaches>${entry}</pivotCaches>`
  if (withNamespace.includes('</definedNames>')) {
    return withNamespace.replace('</definedNames>', `</definedNames>${pivotCaches}`)
  }
  if (withNamespace.includes('</sheets>')) {
    return withNamespace.replace('</sheets>', `</sheets>${pivotCaches}`)
  }
  return withNamespace.replace('</workbook>', `${pivotCaches}</workbook>`)
}

function buildCellValueMap(sheet: WorkbookSnapshot['sheets'][number]): Map<string, LiteralInput> {
  const values = new Map<string, LiteralInput>()
  for (const cell of sheet.cells) {
    values.set(cell.address.toUpperCase(), cell.value ?? null)
  }
  return values
}

function readCellValue(cellsByAddress: ReadonlyMap<string, LiteralInput>, address: string): LiteralInput {
  return cellsByAddress.get(address.toUpperCase()) ?? null
}

function fallbackColumnName(index: number): string {
  return `Column ${String(index + 1)}`
}

function buildPivotCacheTable(snapshot: WorkbookSnapshot, pivot: WorkbookPivotSnapshot): PivotCacheTable | null {
  if (!pivot.source) {
    return null
  }
  const sourceRange = pivot.source
  const sourceSheet = snapshot.sheets.find((sheet) => sheet.name === sourceRange.sheetName)
  if (!sourceSheet) {
    return null
  }
  let source: XLSX.Range
  try {
    source = XLSX.utils.decode_range(`${sourceRange.startAddress}:${sourceRange.endAddress}`)
  } catch {
    return null
  }
  const cellsByAddress = buildCellValueMap(sourceSheet)
  const fields: PivotCacheField[] = []
  for (let column = source.s.c; column <= source.e.c; column += 1) {
    const headerAddress = XLSX.utils.encode_cell({ r: source.s.r, c: column })
    const rawHeader = readCellValue(cellsByAddress, headerAddress)
    const name = typeof rawHeader === 'string' && rawHeader.trim().length > 0 ? rawHeader.trim() : fallbackColumnName(column - source.s.c)
    const values: LiteralInput[] = []
    for (let row = source.s.r + 1; row <= source.e.r; row += 1) {
      values.push(readCellValue(cellsByAddress, XLSX.utils.encode_cell({ r: row, c: column })))
    }
    fields.push({ name, values })
  }
  const rows: LiteralInput[][] = []
  for (let row = source.s.r + 1; row <= source.e.r; row += 1) {
    const values: LiteralInput[] = []
    for (let column = source.s.c; column <= source.e.c; column += 1) {
      values.push(readCellValue(cellsByAddress, XLSX.utils.encode_cell({ r: row, c: column })))
    }
    rows.push(values)
  }
  return { fields, rows }
}

function buildCacheFieldXml(field: PivotCacheField): string {
  return [`<cacheField name="${escapeXml(field.name)}" numFmtId="0">`, '<sharedItems/>', '</cacheField>'].join('')
}

function buildPivotCacheDefinitionXml(
  pivot: WorkbookPivotSnapshot & { source: CellRangeRef },
  cacheTable: PivotCacheTable,
  exportSourceSheetName: string,
): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotCacheDefinition xmlns="${spreadsheetNamespace}" invalid="1" refreshOnLoad="1" refreshedVersion="8" createdVersion="8" minRefreshableVersion="3" recordCount="0">`,
    '<cacheSource type="worksheet">',
    `<worksheetSource ref="${escapeXml(rangeRefA1(pivot.source))}" sheet="${escapeXml(exportSourceSheetName)}"/>`,
    '</cacheSource>',
    `<cacheFields count="${String(cacheTable.fields.length)}">`,
    ...cacheTable.fields.map(buildCacheFieldXml),
    '</cacheFields>',
    '</pivotCacheDefinition>',
  ].join('')
}

function pivotFieldXml(index: number, fields: readonly string[], pivot: WorkbookPivotSnapshot): string {
  const fieldName = fields[index] ?? ''
  if (pivot.groupBy.includes(fieldName)) {
    return '<pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
  }
  if ((pivot.columnFields ?? []).includes(fieldName)) {
    return '<pivotField axis="axisCol" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
  }
  if ((pivot.pageFields ?? []).some((field) => field.sourceColumn === fieldName)) {
    return '<pivotField axis="axisPage" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
  }
  if (pivot.values.some((value) => value.sourceColumn === fieldName)) {
    return '<pivotField dataField="1" showAll="0"/>'
  }
  return '<pivotField showAll="0"/>'
}

function buildRowFieldsXml(fieldIndexes: readonly number[]): string {
  if (fieldIndexes.length === 0) {
    return ''
  }
  return [
    `<rowFields count="${String(fieldIndexes.length)}">`,
    ...fieldIndexes.map((index) => `<field x="${String(index)}"/>`),
    '</rowFields>',
  ].join('')
}

function buildColumnFieldsXml(fieldIndexes: readonly number[]): string {
  if (fieldIndexes.length === 0) {
    return ''
  }
  return [
    `<colFields count="${String(fieldIndexes.length)}">`,
    ...fieldIndexes.map((index) => `<field x="${String(index)}"/>`),
    '</colFields>',
  ].join('')
}

function buildPageFieldsXml(fieldIndexes: readonly number[]): string {
  if (fieldIndexes.length === 0) {
    return ''
  }
  return [
    `<pageFields count="${String(fieldIndexes.length)}">`,
    ...fieldIndexes.map((index) => `<pageField fld="${String(index)}"/>`),
    '</pageFields>',
  ].join('')
}

function defaultDataFieldName(value: WorkbookPivotValueSnapshot): string {
  if (value.outputLabel && value.outputLabel.trim().length > 0) {
    return value.outputLabel.trim()
  }
  return `${defaultDataFieldVerb(value.summarizeBy)} of ${value.sourceColumn}`
}

function buildDataFieldsXml(values: readonly WorkbookPivotValueSnapshot[], fieldIndexesByName: ReadonlyMap<string, number>): string {
  if (values.length === 0) {
    return ''
  }
  const fields = values.flatMap((value) => {
    const fieldIndex = fieldIndexesByName.get(value.sourceColumn)
    if (fieldIndex === undefined) {
      return []
    }
    return [
      `<dataField name="${escapeXml(defaultDataFieldName(value))}" fld="${String(fieldIndex)}" subtotal="${subtotalValue(
        value.summarizeBy,
      )}"/>`,
    ]
  })
  return fields.length > 0 ? [`<dataFields count="${String(fields.length)}">`, ...fields, '</dataFields>'].join('') : ''
}

function buildPivotTableDefinitionXml(pivot: WorkbookPivotSnapshot, cacheId: number, cacheTable: PivotCacheTable): string {
  const fieldNames = cacheTable.fields.map((field) => field.name)
  const fieldIndexesByName = new Map(fieldNames.map((name, index) => [name, index]))
  const rowFieldIndexes = pivot.groupBy.flatMap((fieldName) => {
    const index = fieldIndexesByName.get(fieldName)
    return index === undefined ? [] : [index]
  })
  const columnFieldIndexes = (pivot.columnFields ?? []).flatMap((fieldName) => {
    const index = fieldIndexesByName.get(fieldName)
    return index === undefined ? [] : [index]
  })
  const pageFieldIndexes = (pivot.pageFields ?? []).flatMap((field) => {
    const index = fieldIndexesByName.get(field.sourceColumn)
    return index === undefined ? [] : [index]
  })
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotTableDefinition xmlns="${spreadsheetNamespace}" name="${escapeXml(pivot.name)}" cacheId="${String(
      cacheId,
    )}" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">`,
    `<location ref="${escapeXml(pivotOutputRange(pivot))}" firstHeaderRow="1" firstDataRow="2" firstDataCol="${String(
      Math.max(1, rowFieldIndexes.length),
    )}"/>`,
    `<pivotFields count="${String(fieldNames.length)}">`,
    ...fieldNames.map((_, index) => pivotFieldXml(index, fieldNames, pivot)),
    '</pivotFields>',
    buildRowFieldsXml(rowFieldIndexes),
    buildColumnFieldsXml(columnFieldIndexes),
    buildPageFieldsXml(pageFieldIndexes),
    buildDataFieldsXml(pivot.values, fieldIndexesByName),
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>',
    '</pivotTableDefinition>',
  ].join('')
}

export function addExportPivotsToXlsxBytes(
  bytes: Uint8Array,
  snapshot: WorkbookSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): Uint8Array {
  if (snapshot.workbook.metadata?.pivotArtifacts) {
    return addExportPreservedPivotArtifactsToXlsxBytes(bytes, snapshot)
  }
  const pivots = snapshot.workbook.metadata?.pivots ?? []
  if (pivots.length === 0) {
    return bytes
  }
  const zip = unzipSync(bytes)
  let nextPivotTableIndex = nextPartIndex(zip, 'xl/pivotTables/pivotTable', '.xml')
  let nextPivotCacheIndex = nextPartIndex(zip, 'xl/pivotCache/pivotCacheDefinition', '.xml')
  let contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  let workbookXml = getZipText(zip, 'xl/workbook.xml') ?? ''
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPivots = pivots.filter((pivot) => pivot.sheetName === sheet.name)
      if (sheetPivots.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      let sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      let updatedSheetXml = sheetXml
      const sheetRelsPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const sheetRelationships = parseRelationships(getZipText(zip, sheetRelsPath))

      sheetPivots.forEach((pivot) => {
        if (!pivot.source) {
          return
        }
        const sourcefulPivot = { ...pivot, source: pivot.source }
        const cacheTable = buildPivotCacheTable(snapshot, pivot)
        const exportSourceSheetName = exportSheetNamesByOriginalName.get(pivot.source.sheetName) ?? pivot.source.sheetName
        if (!cacheTable || cacheTable.fields.length === 0 || pivot.values.length === 0 || !exportSourceSheetName) {
          return
        }
        const pivotTableIndex = nextPivotTableIndex
        nextPivotTableIndex += 1
        const cacheIndex = nextPivotCacheIndex
        nextPivotCacheIndex += 1
        const pivotTablePath = `xl/pivotTables/pivotTable${String(pivotTableIndex)}.xml`
        const cacheDefinitionPath = `xl/pivotCache/pivotCacheDefinition${String(cacheIndex)}.xml`

        const workbookRelationshipId = nextRelationshipId(workbookRelationships)
        workbookRelationships.push({
          id: workbookRelationshipId,
          type: pivotCacheDefinitionRelationshipType,
          target: `pivotCache/pivotCacheDefinition${String(cacheIndex)}.xml`,
        })
        workbookXml = addWorkbookPivotCache(workbookXml, cacheIndex, workbookRelationshipId)

        setZipText(zip, pivotTablePath, buildPivotTableDefinitionXml(pivot, cacheIndex, cacheTable))
        setZipText(zip, cacheDefinitionPath, buildPivotCacheDefinitionXml(sourcefulPivot, cacheTable, exportSourceSheetName))

        const sheetRelationshipId = nextRelationshipId(sheetRelationships)
        sheetRelationships.push({
          id: sheetRelationshipId,
          type: pivotTableRelationshipType,
          target: `../pivotTables/pivotTable${String(pivotTableIndex)}.xml`,
        })
        updatedSheetXml = addWorksheetPivotTableDefinition(updatedSheetXml, sheetRelationshipId, pivot)

        contentTypesXml = addContentTypeOverride(contentTypesXml, `/${pivotTablePath}`, pivotTableContentType)
        contentTypesXml = addContentTypeOverride(contentTypesXml, `/${cacheDefinitionPath}`, pivotCacheDefinitionContentType)
      })

      setZipText(zip, sheetRelsPath, buildRelationshipsXml(sheetRelationships))
      setZipText(zip, sheetPath, updatedSheetXml)
    })

  if (workbookXml.length > 0) {
    setZipText(zip, 'xl/workbook.xml', workbookXml)
  }
  setZipText(zip, 'xl/_rels/workbook.xml.rels', buildRelationshipsXml(workbookRelationships))
  if (contentTypesXml.length > 0) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
  }
  return zipSync(zip)
}
