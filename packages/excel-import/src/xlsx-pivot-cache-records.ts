import { XMLParser } from 'fast-xml-parser'
import type { LiteralInput } from '@bilig/protocol'
import { parseRelationships, pivotCacheRecordsRelationshipType, resolveTargetPath } from './xlsx-pivot-artifacts.js'
import { getZipText, type XlsxZipEntries } from './xlsx-zip.js'

interface CacheRecordField {
  readonly sharedItems: readonly LiteralInput[]
}

const orderedXmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
  preserveOrder: true,
})

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function stringAttribute(value: unknown): string | null {
  return typeof value === 'string' && value.trim().length > 0 ? value : null
}

function numberAttribute(value: unknown): number | null {
  const number = typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : Number.NaN
  return Number.isFinite(number) ? number : null
}

function orderedAttributes(node: Record<string, unknown>): Record<string, unknown> {
  const attributes = node[':@']
  return isRecord(attributes) ? attributes : {}
}

function orderedElementName(node: Record<string, unknown>): string | null {
  return Object.keys(node).find((key) => key !== ':@') ?? null
}

function orderedElementChildren(node: Record<string, unknown>, name: string): unknown[] {
  const children = node[name]
  return Array.isArray(children) ? children : []
}

function readInlineCacheRecordValue(name: string, attributes: Record<string, unknown>): LiteralInput | undefined {
  switch (name) {
    case 'm':
      return null
    case 'n': {
      const raw = stringAttribute(attributes['v'])
      if (raw === null) {
        return undefined
      }
      const value = Number(raw)
      return Number.isFinite(value) ? value : raw
    }
    case 'b': {
      const raw = stringAttribute(attributes['v'])
      return raw === null ? undefined : raw === '1' || raw.toLocaleLowerCase('en-US') === 'true'
    }
    case 's':
    case 'e':
    case 'd':
      return stringAttribute(attributes['v']) ?? null
    default:
      return undefined
  }
}

function readSharedCacheRecordValue(
  fields: readonly CacheRecordField[],
  fieldIndex: number,
  attributes: Record<string, unknown>,
): LiteralInput | undefined {
  const itemIndex = numberAttribute(attributes['v'])
  if (itemIndex === null) {
    return undefined
  }
  return fields[fieldIndex]?.sharedItems[itemIndex]
}

function readRecordCellValue(
  fields: readonly CacheRecordField[],
  fieldIndex: number,
  node: Record<string, unknown>,
): LiteralInput | undefined {
  const name = orderedElementName(node)
  if (!name) {
    return undefined
  }
  const attributes = orderedAttributes(node)
  return name === 'x' ? readSharedCacheRecordValue(fields, fieldIndex, attributes) : readInlineCacheRecordValue(name, attributes)
}

function readRecordRowsFromNode(fields: readonly CacheRecordField[], node: Record<string, unknown>): LiteralInput[][] {
  const name = orderedElementName(node)
  if (!name) {
    return []
  }
  if (name === 'r') {
    const row: LiteralInput[] = []
    for (const child of orderedElementChildren(node, name)) {
      if (!isRecord(child)) {
        continue
      }
      const value = readRecordCellValue(fields, row.length, child)
      row.push(value ?? null)
    }
    return row.length > 0 ? [row] : []
  }
  return orderedElementChildren(node, name).flatMap((child) => (isRecord(child) ? readRecordRowsFromNode(fields, child) : []))
}

export function readPivotCacheRecords(xml: string | null | undefined, fields: readonly CacheRecordField[]): LiteralInput[][] {
  if (!xml) {
    return []
  }
  const parsed = orderedXmlParser.parse(xml) as unknown
  return Array.isArray(parsed) ? parsed.flatMap((node) => (isRecord(node) ? readRecordRowsFromNode(fields, node) : [])) : []
}

function pivotCacheDefinitionRelationshipsPath(path: string): string {
  const slashIndex = path.lastIndexOf('/')
  const directory = slashIndex === -1 ? '' : path.slice(0, slashIndex + 1)
  const fileName = slashIndex === -1 ? path : path.slice(slashIndex + 1)
  return `${directory}_rels/${fileName}.rels`
}

export function readPivotCacheRecordsForDefinition(
  zip: XlsxZipEntries,
  path: string,
  fields: readonly CacheRecordField[],
): LiteralInput[][] {
  const relationships = parseRelationships(getZipText(zip, pivotCacheDefinitionRelationshipsPath(path)))
  const recordsRelationship = relationships.find((relationship) => relationship.type === pivotCacheRecordsRelationshipType)
  return recordsRelationship ? readPivotCacheRecords(getZipText(zip, resolveTargetPath(path, recordsRelationship.target)), fields) : []
}
