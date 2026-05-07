import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type { LiteralInput, WorkbookPropertySnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const customPropertiesNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
const documentPropertiesValueTypesNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
const customPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
const customPropertiesContentType = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
const customPropertiesPartPath = 'docProps/custom.xml'
const customPropertiesFormatId = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

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

function textValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  return null
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getZipText(zip: ZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

function setZipText(zip: ZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(xml)
  return asArray(recordChild(parsed, 'Relationships')?.['Relationship']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['Id'] !== 'string' || typeof entry['Target'] !== 'string' || typeof entry['Type'] !== 'string') {
      return []
    }
    return [{ id: entry['Id'], target: entry['Target'], type: entry['Type'] }]
  })
}

function nextRelationshipId(relationships: readonly ParsedRelationship[]): string {
  let next = 1
  for (const relationship of relationships) {
    const match = /^rId(\d+)$/u.exec(relationship.id)
    if (match) {
      next = Math.max(next, Number(match[1]) + 1)
    }
  }
  return `rId${String(next)}`
}

function buildRelationshipsXml(relationships: readonly ParsedRelationship[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) =>
        `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(
          relationship.target,
        )}"/>`,
    ),
    '</Relationships>',
  ].join('')
}

function addContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (contentTypesXml.includes(`PartName="${partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)
}

function normalizableProperty(property: WorkbookPropertySnapshot): WorkbookPropertySnapshot | null {
  if (property.key.trim().length === 0) {
    return null
  }
  if (property.value === null) {
    return null
  }
  if (typeof property.value === 'number' && !Number.isFinite(property.value)) {
    return null
  }
  return { key: property.key, value: property.value }
}

function normalizeWorkbookProperties(properties: readonly WorkbookPropertySnapshot[] | undefined): WorkbookPropertySnapshot[] {
  const byKey = new Map<string, WorkbookPropertySnapshot>()
  for (const property of properties ?? []) {
    const normalized = normalizableProperty(property)
    if (normalized) {
      byKey.set(normalized.key, normalized)
    }
  }
  return [...byKey.values()].toSorted((left, right) => left.key.localeCompare(right.key))
}

function buildCustomPropertyValueXml(value: LiteralInput): string | null {
  if (typeof value === 'string') {
    return `<vt:lpwstr>${escapeXml(value)}</vt:lpwstr>`
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `<vt:r8>${escapeXml(String(value))}</vt:r8>`
  }
  if (typeof value === 'boolean') {
    return `<vt:bool>${value ? 'true' : 'false'}</vt:bool>`
  }
  return null
}

function buildCustomPropertiesXml(properties: readonly WorkbookPropertySnapshot[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Properties xmlns="${customPropertiesNamespace}" xmlns:vt="${documentPropertiesValueTypesNamespace}">`,
    ...properties.flatMap((property, index) => {
      const valueXml = buildCustomPropertyValueXml(property.value)
      if (!valueXml) {
        return []
      }
      return [
        `<property fmtid="${customPropertiesFormatId}" pid="${String(index + 2)}" name="${escapeXml(property.key)}">${valueXml}</property>`,
      ]
    }),
    '</Properties>',
  ].join('')
}

function ensureCustomPropertiesRelationship(zip: ZipEntries): void {
  const relationships = parseRelationships(getZipText(zip, '_rels/.rels'))
  if (!relationships.some((relationship) => relationship.type === customPropertiesRelationshipType)) {
    relationships.push({
      id: nextRelationshipId(relationships),
      target: customPropertiesPartPath,
      type: customPropertiesRelationshipType,
    })
  }
  setZipText(zip, '_rels/.rels', buildRelationshipsXml(relationships))
}

function readCustomPropertiesPartPath(zip: ZipEntries): string | null {
  const relationship = parseRelationships(getZipText(zip, '_rels/.rels')).find((entry) => entry.type === customPropertiesRelationshipType)
  if (relationship) {
    return normalizeZipPath(relationship.target)
  }
  return zip[customPropertiesPartPath] ? customPropertiesPartPath : null
}

function readPropertyValue(property: Record<string, unknown>): LiteralInput | undefined {
  const stringValue = textValue(property['lpwstr'] ?? property['lpstr'])
  if (stringValue !== null) {
    return stringValue
  }

  const numberValue = textValue(property['r8'] ?? property['decimal'] ?? property['i4'] ?? property['int'] ?? property['ui4'])
  if (numberValue !== null) {
    const parsed = Number(numberValue)
    return Number.isFinite(parsed) ? parsed : undefined
  }

  const booleanValue = textValue(property['bool'])
  if (booleanValue !== null) {
    const normalized = booleanValue.trim().toLowerCase()
    if (normalized === 'true' || normalized === '1') {
      return true
    }
    if (normalized === 'false' || normalized === '0') {
      return false
    }
  }
  return undefined
}

export function addExportWorkbookPropertiesToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const properties = normalizeWorkbookProperties(snapshot.workbook.metadata?.properties)
  if (properties.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const contentTypesXml = getZipText(zip, '[Content_Types].xml')
  if (!contentTypesXml) {
    return bytes
  }

  setZipText(zip, customPropertiesPartPath, buildCustomPropertiesXml(properties))
  setZipText(
    zip,
    '[Content_Types].xml',
    addContentTypeOverride(contentTypesXml, `/${customPropertiesPartPath}`, customPropertiesContentType),
  )
  ensureCustomPropertiesRelationship(zip)
  return zipSync(zip)
}

export function readImportedWorkbookProperties(source: XlsxZipSource): WorkbookPropertySnapshot[] | undefined {
  const zip = readXlsxZipEntries(source)
  const partPath = readCustomPropertiesPartPath(zip)
  if (!partPath) {
    return undefined
  }
  const customPropertiesXml = getZipText(zip, partPath)
  if (!customPropertiesXml) {
    return undefined
  }

  const parsed: unknown = xmlParser.parse(customPropertiesXml)
  const properties = asArray(recordChild(parsed, 'Properties')?.['property']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string' || entry['name'].trim().length === 0) {
      return []
    }
    const value = readPropertyValue(entry)
    if (value === undefined) {
      return []
    }
    return [{ key: entry['name'], value }]
  })
  const normalizedProperties = normalizeWorkbookProperties(properties)
  return normalizedProperties.length > 0 ? normalizedProperties : undefined
}
