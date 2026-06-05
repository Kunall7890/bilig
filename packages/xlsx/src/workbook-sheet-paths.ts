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
  const workbookSheets = workbookSheetPathEntriesForSource(zip)
  const workbookSheetsByName = new Map(workbookSheets.map((entry) => [entry.name, entry]))
  return sheetNames.flatMap((name, index) => {
    const entry = workbookSheetsByName.get(name)
    return entry ? [{ ...entry, index }] : []
  })
}

export function workbookSheetPathEntriesForSource(source: XlsxZipSource): WorkbookSheetPathEntry[] {
  const zip = readXlsxZipEntries(source)
  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const worksheetRelationshipsById = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type === worksheetRelationshipType || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, resolveTargetPath(workbookPath, relationship.target)]),
  )
  const workbookSheets = readWorkbookSheetEntries(getZipText(zip, workbookPath))
  if (workbookSheets.length > 0) {
    return workbookSheets.flatMap((entry, index) => {
      const relationshipPath = worksheetRelationshipsById.get(entry.relationshipId)
      return relationshipPath ? [{ name: entry.name, index, path: relationshipPath }] : []
    })
  }
  return sortedWorksheetPaths(zip).map((path, index) => ({ name: `Sheet${String(index + 1)}`, index, path }))
}
