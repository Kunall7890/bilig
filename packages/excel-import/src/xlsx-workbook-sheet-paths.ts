import { XMLParser } from 'fast-xml-parser'
import type * as XLSX from 'xlsx'

import { parseRelationships, resolveTargetPath } from './xlsx-pivot-artifacts.js'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface WorkbookSheetEntry {
  name: string
  relationshipId: string
}

export interface WorkbookSheetPathEntry {
  readonly name: string
  readonly index: number
  readonly path: string
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: true,
})

const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

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

function workbookRecord(workbook: XLSX.WorkBook): Record<string, unknown> | null {
  const value: unknown = workbook
  return isRecord(value) ? value : null
}

function workbookFiles(workbook: XLSX.WorkBook): unknown {
  return workbookRecord(workbook)?.['files']
}

function getFileText(files: unknown, path: string): string | null {
  if (!isRecord(files)) {
    return null
  }
  const file = files[normalizeZipPath(path)]
  if (!isRecord(file)) {
    return null
  }
  const content = file['content']
  if (typeof content === 'string') {
    return content
  }
  if (content instanceof ArrayBuffer) {
    return new TextDecoder().decode(content)
  }
  if (ArrayBuffer.isView(content)) {
    return new TextDecoder().decode(content)
  }
  return null
}

function readWorkbookSheetEntries(workbookXml: string | null): WorkbookSheetEntry[] {
  if (!workbookXml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const workbook = recordChild(parsed, 'workbook')
  const sheets = recordChild(workbook, 'sheets')
  return asArray(sheets?.['sheet']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string' || typeof entry['id'] !== 'string') {
      return []
    }
    return [{ name: entry['name'], relationshipId: entry['id'] }]
  })
}

function sortedWorksheetPaths(zip: XlsxZipSource | undefined): string[] {
  if (!zip) {
    return []
  }
  return Object.keys(readXlsxZipEntries(zip))
    .filter((path) => /^xl\/worksheets\/[^/]+\.xml$/u.test(path))
    .toSorted((left, right) => {
      const leftIndex = Number(/^xl\/worksheets\/sheet([0-9]+)\.xml$/u.exec(left)?.[1] ?? Number.MAX_SAFE_INTEGER)
      const rightIndex = Number(/^xl\/worksheets\/sheet([0-9]+)\.xml$/u.exec(right)?.[1] ?? Number.MAX_SAFE_INTEGER)
      return leftIndex - rightIndex || left.localeCompare(right)
    })
}

export function workbookDirectorySheetPaths(workbook: XLSX.WorkBook, source?: XlsxZipSource): string[] {
  const directory = workbookRecord(workbook)?.['Directory']
  if (isRecord(directory)) {
    const paths = asArray(directory['sheets']).flatMap((entry) => (typeof entry === 'string' ? [entry] : []))
    if (paths.length > 0) {
      return paths
    }
  }
  return sortedWorksheetPaths(source)
}

export function workbookSheetPathsByName(workbook: XLSX.WorkBook, source?: XlsxZipSource): Map<string, string> {
  const sourceZip = source ? readXlsxZipEntries(source) : undefined
  const files = workbookFiles(workbook)
  const workbookRelationships = parseRelationships(
    sourceZip ? getZipText(sourceZip, workbookRelationshipsPath) : getFileText(files, workbookRelationshipsPath),
  )
  const worksheetRelationshipsById = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type === worksheetRelationshipType || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))]),
  )
  const output = new Map<string, string>()
  for (const entry of readWorkbookSheetEntries(sourceZip ? getZipText(sourceZip, workbookPath) : getFileText(files, workbookPath))) {
    const worksheetPath = worksheetRelationshipsById.get(entry.relationshipId)
    if (worksheetPath) {
      output.set(entry.name, worksheetPath)
    }
  }
  return output
}

export function workbookSheetPath(
  pathsByName: ReadonlyMap<string, string>,
  fallbackPaths: readonly string[],
  sheetName: string,
  sheetIndex: number,
): string | undefined {
  if (pathsByName.size > 0) {
    // Relationship-aware workbooks can include non-worksheet sheets such as chartsheets.
    // Missing relationship mappings must not be realigned to worksheet paths by index.
    return pathsByName.get(sheetName)
  }
  return fallbackPaths[sheetIndex]
}

export function workbookSheetPathEntries(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
  source?: XlsxZipSource,
): WorkbookSheetPathEntry[] {
  const pathsByName = workbookSheetPathsByName(workbook, source)
  const fallbackPaths = workbookDirectorySheetPaths(workbook, source)
  return sheetNames.flatMap((name, index) => {
    const path = workbookSheetPath(pathsByName, fallbackPaths, name, index)
    return path ? [{ name, index, path }] : []
  })
}
