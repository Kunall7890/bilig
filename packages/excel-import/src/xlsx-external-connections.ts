import { XMLParser } from 'fast-xml-parser'

import type {
  WorkbookExternalConnectionSnapshot,
  WorkbookExternalConnectionSourceKind,
  WorkbookExternalConnectionsSnapshot,
  WorkbookExternalLinkSnapshot,
} from '@bilig/protocol'
import { getZipText, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface ParsedRelationship {
  readonly id: string
  readonly type: string
  readonly target: string
  readonly targetMode?: string
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const externalLinkPathRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'
const oleObjectRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'

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

function stringValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  return null
}

function integerValue(value: unknown): number | null {
  const raw = stringValue(value)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const parsed = Number(raw)
  return Number.isSafeInteger(parsed) ? parsed : null
}

function booleanValue(value: unknown): boolean | undefined {
  const raw = stringValue(value)
  if (raw === null || raw.trim().length === 0) {
    return undefined
  }
  return raw === '1' || raw.toLowerCase() === 'true'
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(xml)
  return asArray(recordChild(parsed, 'Relationships')?.['Relationship']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const id = stringValue(entry['Id'])
    const type = stringValue(entry['Type'])
    const target = stringValue(entry['Target'])
    if (!id || !type || !target) {
      return []
    }
    const targetMode = stringValue(entry['TargetMode'])
    return [{ id, type, target, ...(targetMode ? { targetMode } : {}) }]
  })
}

function resolveTargetPath(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.replace(/^\/+/, '')
  }
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
}

function sourceKindForConnection(type: string | null, connection: Record<string, unknown>): WorkbookExternalConnectionSourceKind {
  if (recordChild(connection, 'dbPr')) {
    return 'database'
  }
  if (recordChild(connection, 'webPr')) {
    return 'web-query'
  }
  if (recordChild(connection, 'textPr')) {
    return 'text'
  }
  if (type === '5') {
    return 'database'
  }
  if (type === '4') {
    return 'web-query'
  }
  return 'unknown'
}

function readConnectionCommand(connection: Record<string, unknown>): {
  readonly command?: string
  readonly commandType?: string
  readonly connection?: string
} {
  const dbPr = recordChild(connection, 'dbPr')
  const webPr = recordChild(connection, 'webPr')
  const textPr = recordChild(connection, 'textPr')
  const source = dbPr ?? webPr ?? textPr ?? connection
  const command = stringValue(source['command'])
  const commandType = stringValue(source['commandType'])
  const connectionText = stringValue(source['connection'])
  return {
    ...(command ? { command } : {}),
    ...(commandType ? { commandType } : {}),
    ...(connectionText ? { connection: connectionText } : {}),
  }
}

function readConnectionsXml(zip: XlsxZipEntries): WorkbookExternalConnectionSnapshot[] {
  const xml = getZipText(zip, 'xl/connections.xml')
  if (!xml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(xml)
  return asArray(recordChild(parsed, 'connections')?.['connection']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const id = integerValue(entry['id'])
    if (id === null) {
      return []
    }
    const type = stringValue(entry['type'])
    const name = stringValue(entry['name'])
    const description = stringValue(entry['description'])
    const refreshOnLoad = booleanValue(entry['refreshOnLoad'])
    const saveData = booleanValue(entry['saveData'])
    return [
      {
        id,
        ...(name ? { name } : {}),
        sourceKind: sourceKindForConnection(type, entry),
        ...(type ? { type } : {}),
        ...(description ? { description } : {}),
        ...readConnectionCommand(entry),
        ...(refreshOnLoad !== undefined ? { refreshOnLoad } : {}),
        ...(saveData !== undefined ? { saveData } : {}),
        clause: '18.13',
      },
    ]
  })
}

function readSheetNames(externalBook: Record<string, unknown>): string[] {
  return asArray(recordChild(externalBook, 'sheetNames')?.['sheetName']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const name = stringValue(entry['val'])
    return name === null || name.trim().length === 0 ? [] : [name]
  })
}

function readDefinedNames(externalBook: Record<string, unknown>): string[] {
  return asArray(recordChild(externalBook, 'definedNames')?.['definedName']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const name = stringValue(entry['name'])
    return name === null || name.trim().length === 0 ? [] : [name]
  })
}

function workbookNameFromExternalTarget(target: string | undefined): string | undefined {
  if (!target) {
    return undefined
  }
  const normalized =
    target
      .split(/[?#]/u)[0]
      ?.replace(/^file:\/+/iu, '')
      .replace(/\\/gu, '/') ?? target
  const lastSegment = normalized.split('/').findLast((segment) => segment.length > 0)
  if (!lastSegment) {
    return undefined
  }
  try {
    return decodeURIComponent(lastSegment)
  } catch {
    return lastSegment
  }
}

function externalLinkRelationshipsPartPath(partPath: string): string {
  const fileName = partPath.slice(partPath.lastIndexOf('/') + 1)
  return `xl/externalLinks/_rels/${fileName}.rels`
}

function readWorkbookExternalLinkTargets(zip: XlsxZipEntries): Map<number, string> {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return new Map()
  }
  const relationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const parsed: unknown = xmlParser.parse(workbookXml)
  const targets = new Map<number, string>()
  let bookIndex = 1
  for (const entry of asArray(recordChild(recordChild(parsed, 'workbook'), 'externalReferences')?.['externalReference'])) {
    if (isRecord(entry) && typeof entry['id'] === 'string') {
      const relationship = relationships.find(
        (candidate) => candidate.id === entry['id'] && candidate.type === externalLinkRelationshipType,
      )
      if (relationship) {
        targets.set(bookIndex, resolveTargetPath('xl/workbook.xml', relationship.target))
      }
    }
    bookIndex += 1
  }
  if (targets.size > 0) {
    return targets
  }
  for (const path of Object.keys(zip)) {
    const match = /^xl\/externalLinks\/externalLink([1-9][0-9]*)\.xml$/u.exec(path)
    if (match) {
      targets.set(Number(match[1]), path)
    }
  }
  return targets
}

function readItemNames(parent: Record<string, unknown>, collectionName: string, itemName: string): string[] {
  return asArray(recordChild(parent, collectionName)?.[itemName]).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const name = stringValue(entry['name'])
    return name === null || name.trim().length === 0 ? [] : [name]
  })
}

function readExternalLinkProvenance(
  bookIndex: number,
  path: string,
  xml: string,
  relationships: readonly ParsedRelationship[],
): WorkbookExternalLinkSnapshot[] {
  const parsed: unknown = xmlParser.parse(xml)
  const externalLink = recordChild(parsed, 'externalLink')
  if (!externalLink) {
    return []
  }
  const output: WorkbookExternalLinkSnapshot[] = []
  const externalBook = recordChild(externalLink, 'externalBook')
  if (externalBook) {
    const relationshipId = stringValue(externalBook['id'])
    const relationship =
      (relationshipId ? relationships.find((candidate) => candidate.id === relationshipId) : undefined) ??
      relationships.find((candidate) => candidate.type === externalLinkPathRelationshipType)
    const workbookName = workbookNameFromExternalTarget(relationship?.target)
    output.push({
      kind: 'external-workbook',
      bookIndex,
      packagePath: path,
      ...(relationship?.target ? { target: relationship.target } : {}),
      ...(relationship?.targetMode ? { targetMode: relationship.targetMode } : {}),
      ...(workbookName ? { workbookName } : {}),
      sheetNames: readSheetNames(externalBook),
      definedNames: readDefinedNames(externalBook),
      clause: '18.14',
    })
  }
  for (const ddeLink of asArray(externalLink['ddeLink'])) {
    if (!isRecord(ddeLink)) {
      continue
    }
    const service = stringValue(ddeLink['ddeService'])
    const topic = stringValue(ddeLink['ddeTopic'])
    output.push({
      kind: 'dde',
      ...(service ? { service } : {}),
      ...(topic ? { topic } : {}),
      itemNames: readItemNames(ddeLink, 'ddeItems', 'ddeItem'),
      refreshExecution: 'disabled',
      packagePath: path,
      clause: '18.14',
    })
  }
  for (const oleLink of asArray(externalLink['oleLink'])) {
    if (!isRecord(oleLink)) {
      continue
    }
    const relationshipId = stringValue(oleLink['id']) ?? undefined
    const relationship = relationshipId
      ? relationships.find((candidate) => candidate.id === relationshipId && candidate.type === oleObjectRelationshipType)
      : undefined
    const progId = stringValue(oleLink['progId'])
    output.push({
      kind: 'ole',
      ...(progId ? { progId } : {}),
      ...(relationshipId ? { relationshipId } : {}),
      ...(relationship?.target ? { target: relationship.target } : {}),
      ...(relationship?.targetMode ? { targetMode: relationship.targetMode } : {}),
      itemNames: readItemNames(oleLink, 'oleItems', 'oleItem'),
      refreshExecution: 'disabled',
      packagePath: path,
      clause: '18.14',
    })
  }
  return output
}

function readExternalLinks(zip: XlsxZipEntries): WorkbookExternalLinkSnapshot[] {
  const links: WorkbookExternalLinkSnapshot[] = []
  for (const [bookIndex, path] of [...readWorkbookExternalLinkTargets(zip).entries()].toSorted((left, right) => left[0] - right[0])) {
    const xml = getZipText(zip, path)
    if (!xml) {
      continue
    }
    links.push(
      ...readExternalLinkProvenance(bookIndex, path, xml, parseRelationships(getZipText(zip, externalLinkRelationshipsPartPath(path)))),
    )
  }
  return links
}

export function readImportedWorkbookExternalConnections(source: XlsxZipSource): WorkbookExternalConnectionsSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const connections = readConnectionsXml(zip)
  const externalLinks = readExternalLinks(zip)
  if (connections.length === 0 && externalLinks.length === 0) {
    return undefined
  }
  return {
    refreshExecution: 'disabled',
    ...(connections.length > 0 ? { connections } : {}),
    ...(externalLinks.length > 0 ? { externalLinks } : {}),
  }
}
