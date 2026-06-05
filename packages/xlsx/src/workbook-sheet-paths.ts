import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipSource } from './zip-reader.js'
import { readXmlAttribute } from './xml.js'

interface WorkbookSheetEntry {
  readonly name: string
  readonly relationshipId: string
}

interface RelationshipEntry {
  readonly id: string
  readonly target: string
  readonly type: string
}

export interface WorkbookSheetPathEntry {
  readonly name: string
  readonly index: number
  readonly path: string
}

const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

function parseRelationships(xml: string | null): RelationshipEntry[] {
  if (!xml) {
    return []
  }
  return [...xml.matchAll(/<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const tag = match[0]
    const id = readXmlAttribute(tag, 'Id')
    const target = readXmlAttribute(tag, 'Target')
    const type = readXmlAttribute(tag, 'Type')
    return id && target && type ? [{ id, target, type }] : []
  })
}

function readWorkbookSheetEntries(workbookXml: string | null): WorkbookSheetEntry[] {
  if (!workbookXml) {
    return []
  }
  return [...workbookXml.matchAll(/<sheet\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const tag = match[0]
    const name = readXmlAttribute(tag, 'name')
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    return name && relationshipId ? [{ name, relationshipId }] : []
  })
}

function resolveTargetPath(fromPath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target)
  }
  const baseParts = normalizeZipPath(fromPath).split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part.length === 0 || part === '.') {
      continue
    }
    if (part === '..') {
      baseParts.pop()
      continue
    }
    baseParts.push(part)
  }
  return normalizeZipPath(baseParts.join('/'))
}

function sortedWorksheetPaths(zip: XlsxZipSource): string[] {
  return Object.keys(readXlsxZipEntries(zip))
    .filter((path) => /^xl\/worksheets\/[^/]+\.xml$/u.test(path))
    .toSorted((left, right) => {
      const leftIndex = Number(/^xl\/worksheets\/sheet([0-9]+)\.xml$/u.exec(left)?.[1] ?? Number.MAX_SAFE_INTEGER)
      const rightIndex = Number(/^xl\/worksheets\/sheet([0-9]+)\.xml$/u.exec(right)?.[1] ?? Number.MAX_SAFE_INTEGER)
      return leftIndex - rightIndex || left.localeCompare(right)
    })
}

export function workbookSheetPathEntriesFromSource(source: XlsxZipSource, sheetNames: readonly string[]): WorkbookSheetPathEntry[] {
  const zip = readXlsxZipEntries(source)
  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const worksheetRelationshipsById = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type === worksheetRelationshipType || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, resolveTargetPath(workbookPath, relationship.target)]),
  )
  const sheetRelationshipsByName = new Map(
    readWorkbookSheetEntries(getZipText(zip, workbookPath)).map((entry) => [entry.name, entry.relationshipId]),
  )
  const fallbackPaths = sortedWorksheetPaths(zip)
  return sheetNames.flatMap((name, index) => {
    const relationshipId = sheetRelationshipsByName.get(name)
    const relationshipPath = relationshipId ? worksheetRelationshipsById.get(relationshipId) : undefined
    if (relationshipPath) {
      return [{ name, index, path: relationshipPath }]
    }
    if (sheetRelationshipsByName.size > 0) {
      return []
    }
    const fallbackPath = fallbackPaths[index]
    return fallbackPath ? [{ name, index, path: fallbackPath }] : []
  })
}
