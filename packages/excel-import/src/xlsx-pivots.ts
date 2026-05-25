import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type {
  CellRangeRef,
  LiteralInput,
  PivotAggregation,
  WorkbookDefinedNameSnapshot,
  WorkbookPivotArtifactsSnapshot,
  WorkbookPivotSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookSheetPivotArtifactsSnapshot,
  WorkbookTableSnapshot,
  WorkbookUnsupportedPivotSnapshot,
} from '@bilig/protocol'
import {
  parseRelationships,
  pivotCacheDefinitionRelationshipType,
  pivotTableRelationshipType,
  readImportedPivotArtifacts,
  resolveTargetPath,
} from './xlsx-pivot-artifacts.js'
import { readCacheFieldSharedItems } from './xlsx-pivot-cache-items.js'
import { readPivotCacheRecordsForDefinition } from './xlsx-pivot-cache-records.js'
import { workbookSheetPathEntriesFromSource } from './xlsx-workbook-sheet-paths.js'
import { getZipText, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = XlsxZipEntries

interface ParsedPivotCacheField {
  readonly name: string
  readonly sharedItems: readonly LiteralInput[]
}

type PivotSourceKind = NonNullable<WorkbookPivotSnapshot['sourceKind']>

interface ParsedPivotCache {
  readonly cacheId: number
  readonly source?: CellRangeRef
  readonly sourceKind: PivotSourceKind
  readonly fields: readonly ParsedPivotCacheField[]
  readonly cachedRecordCount?: number
  readonly cachedRecords?: readonly (readonly LiteralInput[])[]
}

export interface ImportedWorkbookPivots {
  readonly pivots: WorkbookPivotSnapshot[] | undefined
  readonly unsupportedPivots: WorkbookUnsupportedPivotSnapshot[] | undefined
  readonly hasExternalPivotCaches: boolean
  readonly artifacts: WorkbookPivotArtifactsSnapshot | undefined
  readonly sheetArtifactsByName: Map<string, WorkbookSheetPivotArtifactsSnapshot>
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

function relationshipPathForPart(path: string): string {
  const slashIndex = path.lastIndexOf('/')
  return slashIndex >= 0 ? `${path.slice(0, slashIndex)}/_rels/${path.slice(slashIndex + 1)}.rels` : `_rels/${path}.rels`
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function stringAttribute(value: unknown): string | null {
  return typeof value === 'string' && value.trim().length > 0 ? value : null
}

function parseRangeRef(sheetName: string, ref: string): CellRangeRef | null {
  try {
    const decoded = XLSX.utils.decode_range(ref.replaceAll('$', ''))
    return {
      sheetName,
      startAddress: XLSX.utils.encode_cell(decoded.s),
      endAddress: XLSX.utils.encode_cell(decoded.e),
    }
  } catch {
    return null
  }
}

function numberAttribute(value: unknown): number | null {
  const number = typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : Number.NaN
  return Number.isFinite(number) ? number : null
}

function readWorkbookPivotCaches(zip: ZipEntries): Map<number, string> {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const output = new Map<number, string>()
  if (!workbookXml) {
    return output
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const pivotCaches = asArray(recordChild(recordChild(parsed, 'workbook'), 'pivotCaches')?.['pivotCache'])
  pivotCaches.forEach((entry) => {
    if (!isRecord(entry)) {
      return
    }
    const cacheId = numberAttribute(entry['cacheId'])
    const relationshipId = typeof entry['id'] === 'string' ? entry['id'] : null
    const relationship = relationshipId
      ? workbookRelationships.find(
          (candidate) => candidate.id === relationshipId && candidate.type === pivotCacheDefinitionRelationshipType,
        )
      : undefined
    if (cacheId !== null && relationship) {
      output.set(cacheId, resolveTargetPath('xl/workbook.xml', relationship.target))
    }
  })
  return output
}

function parsePivotCacheDefinition(
  cacheId: number,
  xml: string,
  zip: ZipEntries,
  path: string,
  tablesByName: ReadonlyMap<string, WorkbookTableSnapshot>,
  definedNamesByName: ReadonlyMap<string, WorkbookDefinedNameSnapshot>,
): ParsedPivotCache | null {
  const parsed: unknown = xmlParser.parse(xml)
  const definition = recordChild(parsed, 'pivotCacheDefinition')
  const cacheSource = recordChild(definition, 'cacheSource')
  const fields = asArray(recordChild(definition, 'cacheFields')?.['cacheField']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string' || entry['name'].trim().length === 0) {
      return []
    }
    return [
      {
        name: entry['name'].trim(),
        sharedItems: readCacheFieldSharedItems(entry),
      },
    ]
  })
  const sourceRecord = recordChild(cacheSource, 'worksheetSource')
  const sheetName = typeof sourceRecord?.['sheet'] === 'string' ? sourceRecord['sheet'] : null
  const ref = typeof sourceRecord?.['ref'] === 'string' ? sourceRecord['ref'] : null
  const sourceName = typeof sourceRecord?.['name'] === 'string' ? sourceRecord['name'].trim() : ''
  const namedSource = sourceName.length > 0 ? sourceRangeForName(sourceName, sheetName, tablesByName, definedNamesByName) : null
  const worksheetSource = sheetName && ref ? parseRangeRef(sheetName, ref) : null
  const source = worksheetSource ?? namedSource?.source
  const sourceType = stringAttribute(cacheSource?.['type'])
  const cachedRecords = fields.length > 0 ? readPivotCacheRecordsForDefinition(zip, path, fields) : []
  if (!source && (sourceType !== 'external' || cachedRecords.length === 0)) {
    return null
  }
  const sourceKind =
    sourceType === 'external' ? 'external-cache-only' : worksheetSource ? 'worksheet' : (namedSource?.sourceKind ?? 'worksheet')
  const recordCount = numberAttribute(definition?.['recordCount'])
  return fields.length > 0
    ? {
        cacheId,
        ...(source ? { source } : {}),
        sourceKind,
        fields,
        ...(recordCount !== null ? { cachedRecordCount: recordCount } : {}),
        ...(cachedRecords.length > 0 ? { cachedRecords } : {}),
      }
    : null
}

function pivotCacheDefinitionSourceType(xml: string): string | null {
  const parsed: unknown = xmlParser.parse(xml)
  const cacheSource = recordChild(recordChild(parsed, 'pivotCacheDefinition'), 'cacheSource')
  return typeof cacheSource?.['type'] === 'string' ? cacheSource['type'] : null
}

function sourceRangeForName(
  name: string,
  sheetName: string | null,
  tablesByName: ReadonlyMap<string, WorkbookTableSnapshot>,
  definedNamesByName: ReadonlyMap<string, WorkbookDefinedNameSnapshot>,
): { readonly source: CellRangeRef; readonly sourceKind: PivotSourceKind } | null {
  const normalizedName = name.toLocaleLowerCase('en-US')
  const table = tablesByName.get(normalizedName)
  if (table) {
    return {
      source: {
        sheetName: table.sheetName,
        startAddress: table.startAddress,
        endAddress: table.endAddress,
      },
      sourceKind: 'table',
    }
  }
  const definedName = sheetName
    ? (definedNamesByName.get(definedNameKey(name, sheetName)) ?? definedNamesByName.get(normalizedName))
    : definedNamesByName.get(normalizedName)
  if (!definedName) {
    return null
  }
  const value = definedName.value
  if (!isRecord(value)) {
    return null
  }
  if (
    value['kind'] === 'range-ref' &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  ) {
    return {
      source: { sheetName: value['sheetName'], startAddress: value['startAddress'], endAddress: value['endAddress'] },
      sourceKind: 'named-range',
    }
  }
  if (value['kind'] === 'cell-ref' && typeof value['sheetName'] === 'string' && typeof value['address'] === 'string') {
    return {
      source: { sheetName: value['sheetName'], startAddress: value['address'], endAddress: value['address'] },
      sourceKind: 'named-range',
    }
  }
  return null
}

function definedNameKey(name: string, scopeSheetName: string | undefined): string {
  return scopeSheetName
    ? `${scopeSheetName.toLocaleLowerCase('en-US')}:${name.toLocaleLowerCase('en-US')}`
    : name.toLocaleLowerCase('en-US')
}

function parsePivotCaches(
  zip: ZipEntries,
  tables: readonly WorkbookTableSnapshot[],
  definedNames: readonly WorkbookDefinedNameSnapshot[],
): {
  readonly caches: Map<number, ParsedPivotCache>
  readonly hasExternalPivotCaches: boolean
  readonly unsupportedCaches: Map<number, WorkbookUnsupportedPivotSnapshot>
} {
  const cacheDefinitions = readWorkbookPivotCaches(zip)
  const tablesByName = new Map(tables.map((table) => [table.name.toLocaleLowerCase('en-US'), table]))
  const definedNamesByName = new Map(
    definedNames.map((definedName) => [definedNameKey(definedName.name, definedName.scopeSheetName), definedName]),
  )
  const output = new Map<number, ParsedPivotCache>()
  const unsupportedCaches = new Map<number, WorkbookUnsupportedPivotSnapshot>()
  let hasExternalPivotCaches = false
  for (const [cacheId, path] of cacheDefinitions.entries()) {
    const xml = getZipText(zip, path) ?? ''
    const sourceType = pivotCacheDefinitionSourceType(xml)
    const hasExternalSource = sourceType === 'external'
    hasExternalPivotCaches ||= hasExternalSource
    if (hasExternalSource) {
      const parsedExternalCache = parsePivotCacheDefinition(cacheId, xml, zip, path, tablesByName, definedNamesByName)
      unsupportedCaches.set(cacheId, {
        kind: 'external-cache',
        cacheId,
        sourceType,
        packagePart: path,
        ...(parsedExternalCache?.cachedRecordCount !== undefined ? { cachedRecordCount: parsedExternalCache.cachedRecordCount } : {}),
        ...(parsedExternalCache ? { cacheFieldNames: parsedExternalCache.fields.map((field) => field.name), cacheOnly: true } : {}),
        reason:
          'External pivot cache is preserved as raw XLSX package parts; cached records are provenance-only unless a semantic source can be resolved.',
      })
    }
    const parsed = parsePivotCacheDefinition(cacheId, xml, zip, path, tablesByName, definedNamesByName)
    if (parsed) {
      output.set(cacheId, parsed)
    }
  }
  return { caches: output, hasExternalPivotCaches, unsupportedCaches }
}

function aggregationFromSubtotal(value: unknown): PivotAggregation | null {
  switch (value) {
    case undefined:
    case null:
    case 'sum':
      return 'sum'
    case 'count':
      return 'count'
    case 'countNums':
      return 'countNums'
    case 'average':
      return 'average'
    case 'min':
      return 'min'
    case 'max':
      return 'max'
    case 'product':
      return 'product'
    default:
      return null
  }
}

function cacheFieldName(cache: ParsedPivotCache, index: number | null): string | undefined {
  return index === null ? undefined : cache.fields[index]?.name
}

function cacheSharedItem(cache: ParsedPivotCache, fieldIndex: number, itemIndex: number | null): LiteralInput | undefined {
  if (itemIndex === null) {
    return undefined
  }
  const sharedItem = cache.fields[fieldIndex]?.sharedItems[itemIndex]
  if (sharedItem !== undefined) {
    return sharedItem
  }
  const cachedRecords = cache.cachedRecords
  if (!cachedRecords) {
    return undefined
  }
  const distinctValues: LiteralInput[] = []
  for (const record of cachedRecords) {
    const value = record[fieldIndex]
    if (value === undefined) {
      continue
    }
    if (!distinctValues.some((candidate) => candidate === value)) {
      distinctValues.push(value)
    }
  }
  return distinctValues[itemIndex]
}

function parseFieldNames(definition: Record<string, unknown>, fieldGroupName: string, cache: ParsedPivotCache): string[] {
  return asArray(recordChild(definition, fieldGroupName)?.['field']).flatMap((entry) => {
    const index = isRecord(entry) ? numberAttribute(entry['x']) : null
    const field = cacheFieldName(cache, index)
    return field ? [field] : []
  })
}

function parsePageFields(definition: Record<string, unknown>, cache: ParsedPivotCache): NonNullable<WorkbookPivotSnapshot['pageFields']> {
  return asArray(recordChild(definition, 'pageFields')?.['pageField']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const fieldIndex = numberAttribute(entry['fld'])
    const sourceColumn = cacheFieldName(cache, fieldIndex)
    if (!sourceColumn || fieldIndex === null) {
      return []
    }
    const selectedValue = cacheSharedItem(cache, fieldIndex, numberAttribute(entry['item']))
    return [
      {
        sourceColumn,
        ...(selectedValue !== undefined ? { selectedValue } : {}),
      },
    ]
  })
}

function parseHiddenItems(definition: Record<string, unknown>, cache: ParsedPivotCache): NonNullable<WorkbookPivotSnapshot['hiddenItems']> {
  return asArray(recordChild(definition, 'pivotFields')?.['pivotField']).flatMap((fieldEntry, fieldIndex) => {
    if (!isRecord(fieldEntry)) {
      return []
    }
    const sourceColumn = cache.fields[fieldIndex]?.name
    if (!sourceColumn) {
      return []
    }
    const values = asArray(recordChild(fieldEntry, 'items')?.['item']).flatMap((item) => {
      if (!isRecord(item) || stringAttribute(item['h']) !== '1') {
        return []
      }
      const itemIndex = numberAttribute(item['x'])
      const value = cacheSharedItem(cache, fieldIndex, itemIndex)
      return value === undefined ? [] : [value]
    })
    return values.length > 0 ? [{ sourceColumn, values }] : []
  })
}

function parseCalculatedFields(definition: Record<string, unknown>): NonNullable<WorkbookPivotSnapshot['calculatedFields']> {
  return asArray(recordChild(definition, 'calculatedFields')?.['calculatedField']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const name = stringAttribute(entry['name'])
    const formula = stringAttribute(entry['formula'])
    return name && formula ? [{ name, formula, clause: '18.10' }] : []
  })
}

function parsePivotTableXml(sheetName: string, xml: string, caches: ReadonlyMap<number, ParsedPivotCache>): WorkbookPivotSnapshot | null {
  const parsed: unknown = xmlParser.parse(xml)
  const definition = recordChild(parsed, 'pivotTableDefinition')
  const cacheId = numberAttribute(definition?.['cacheId'])
  const cache = cacheId === null ? undefined : caches.get(cacheId)
  const locationRefValue = recordChild(definition, 'location')?.['ref']
  const locationRef = typeof locationRefValue === 'string' ? locationRefValue : null
  if (!definition || !cache || !locationRef) {
    return null
  }
  const location = parseRangeRef(sheetName, locationRef)
  if (!location) {
    return null
  }
  const groupBy = parseFieldNames(definition, 'rowFields', cache)
  const columnFields = parseFieldNames(definition, 'colFields', cache)
  const pageFields = parsePageFields(definition, cache)
  const hiddenItems = parseHiddenItems(definition, cache)
  const calculatedFields = parseCalculatedFields(definition)
  const values = asArray(recordChild(definition, 'dataFields')?.['dataField']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const fieldIndex = numberAttribute(entry['fld'])
    const sourceColumn = cacheFieldName(cache, fieldIndex)
    const summarizeBy = aggregationFromSubtotal(entry['subtotal'])
    if (!sourceColumn || !summarizeBy) {
      return []
    }
    const outputLabel = typeof entry['name'] === 'string' && entry['name'].trim().length > 0 ? entry['name'].trim() : undefined
    const value: WorkbookPivotValueSnapshot = { sourceColumn, summarizeBy }
    if (outputLabel !== undefined) {
      value.outputLabel = outputLabel
    }
    return [value]
  })
  if (values.length === 0) {
    return null
  }
  const name =
    typeof definition['name'] === 'string' && definition['name'].trim().length > 0
      ? definition['name'].trim()
      : `Pivot ${location.startAddress}`
  const start = XLSX.utils.decode_cell(location.startAddress)
  const end = XLSX.utils.decode_cell(location.endAddress)
  return {
    name,
    sheetName,
    address: location.startAddress,
    ...(cache.source ? { source: cache.source } : {}),
    cacheId: cache.cacheId,
    sourceKind: cache.sourceKind,
    ...(cache.sourceKind === 'external-cache-only' ? { cacheOnly: true } : {}),
    cacheFields: cache.fields.map((field) => field.name),
    ...(cache.cachedRecords ? { cachedRecords: cache.cachedRecords.map((row) => Array.from(row)) } : {}),
    groupBy,
    ...(columnFields.length > 0 ? { columnFields } : {}),
    ...(pageFields.length > 0 ? { pageFields } : {}),
    ...(hiddenItems.length > 0 ? { hiddenItems } : {}),
    ...(calculatedFields.length > 0 ? { calculatedFields } : {}),
    values,
    rows: Math.max(1, end.r - start.r + 1),
    cols: Math.max(1, end.c - start.c + 1),
  }
}

function parseUnsupportedPivotTableXml(
  sheetName: string,
  xml: string,
  pivotPath: string,
  unsupportedCaches: ReadonlyMap<number, WorkbookUnsupportedPivotSnapshot>,
): WorkbookUnsupportedPivotSnapshot | null {
  const parsed: unknown = xmlParser.parse(xml)
  const definition = recordChild(parsed, 'pivotTableDefinition')
  const cacheId = numberAttribute(definition?.['cacheId'])
  if (!definition || cacheId === null) {
    return null
  }
  const cached = unsupportedCaches.get(cacheId)
  const locationRefValue = recordChild(definition, 'location')?.['ref']
  const locationRef = typeof locationRefValue === 'string' ? parseRangeRef(sheetName, locationRefValue) : null
  const name = typeof definition['name'] === 'string' && definition['name'].trim().length > 0 ? definition['name'].trim() : undefined
  if (cached) {
    return {
      ...cached,
      ...(locationRef ? { sheetName, address: locationRef.startAddress } : { sheetName }),
      ...(name ? { name } : {}),
      packagePart: pivotPath,
    }
  }
  return {
    kind: 'raw-part',
    cacheId,
    ...(locationRef ? { sheetName, address: locationRef.startAddress } : { sheetName }),
    ...(name ? { name } : {}),
    packagePart: pivotPath,
    reason: 'Pivot table package parts are preserved but could not be projected into the semantic pivot model.',
  }
}

export function readImportedWorkbookPivots(
  source: XlsxZipSource,
  sheetNames: readonly string[],
  tables: readonly WorkbookTableSnapshot[] = [],
  definedNames: readonly WorkbookDefinedNameSnapshot[] = [],
): ImportedWorkbookPivots {
  const zip = readXlsxZipEntries(source)
  const { artifacts, sheetArtifactsByName } = readImportedPivotArtifacts(zip, sheetNames)
  const { caches, hasExternalPivotCaches, unsupportedCaches } = parsePivotCaches(zip, tables, definedNames)
  const pivots: WorkbookPivotSnapshot[] = []
  const unsupportedPivots: WorkbookUnsupportedPivotSnapshot[] = []
  workbookSheetPathEntriesFromSource(zip, sheetNames).forEach((sheet) => {
    const sheetName = sheet.name
    const sheetPath = sheet.path
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml) {
      return
    }
    const sheetRelationships = parseRelationships(getZipText(zip, relationshipPathForPart(sheetPath)))
    const parsedSheet: unknown = xmlParser.parse(sheetXml)
    const pivotRefs = asArray(recordChild(parsedSheet, 'worksheet')?.['pivotTableDefinition'])
    const pivotRelationships =
      pivotRefs.length > 0
        ? pivotRefs.flatMap((entry) => {
            if (!isRecord(entry) || typeof entry['id'] !== 'string') {
              return []
            }
            const relationship = sheetRelationships.find(
              (candidate) => candidate.id === entry['id'] && candidate.type === pivotTableRelationshipType,
            )
            return relationship ? [relationship] : []
          })
        : sheetRelationships.filter((relationship) => relationship.type === pivotTableRelationshipType)
    pivotRelationships.forEach((relationship) => {
      const pivotPath = resolveTargetPath(sheetPath, relationship.target)
      const pivot = parsePivotTableXml(sheetName, getZipText(zip, pivotPath) ?? '', caches)
      if (pivot) {
        pivots.push(pivot)
        return
      }
      const unsupportedPivot = parseUnsupportedPivotTableXml(sheetName, getZipText(zip, pivotPath) ?? '', pivotPath, unsupportedCaches)
      if (unsupportedPivot) {
        unsupportedPivots.push(unsupportedPivot)
      }
    })
  })
  return {
    pivots:
      pivots.length > 0
        ? pivots.toSorted((left, right) =>
            `${left.sheetName}:${left.address}:${left.name}`.localeCompare(`${right.sheetName}:${right.address}:${right.name}`),
          )
        : undefined,
    unsupportedPivots:
      unsupportedPivots.length > 0
        ? unsupportedPivots.toSorted((left, right) =>
            `${left.sheetName ?? ''}:${left.address ?? ''}:${left.name ?? ''}:${String(left.cacheId ?? '')}`.localeCompare(
              `${right.sheetName ?? ''}:${right.address ?? ''}:${right.name ?? ''}:${String(right.cacheId ?? '')}`,
            ),
          )
        : undefined,
    hasExternalPivotCaches,
    artifacts,
    sheetArtifactsByName,
  }
}
