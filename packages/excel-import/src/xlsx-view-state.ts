import { unzipSync, zipSync } from 'fflate'

import type { SheetMetadataSnapshot, WorkbookMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import { parseRelationships, resolveTargetPath, setZipText } from './xlsx-pivot-artifacts.js'

const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

interface WorkbookSheetEntry {
  readonly name: string
  readonly relationshipId: string
}

export interface ImportedWorkbookViewState {
  readonly workbookViewState?: WorkbookMetadataSnapshot['viewState']
  readonly sheetViewStateByName: Map<string, NonNullable<SheetMetadataSnapshot['viewState']>>
}

function readXmlAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function readWorkbookSheetEntries(workbookXml: string | null): WorkbookSheetEntry[] {
  if (!workbookXml) {
    return []
  }
  return [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const name = readXmlAttribute(attributes, 'name')
    const relationshipId = readXmlAttribute(attributes, 'r:id') ?? readXmlAttribute(attributes, 'id')
    return name && relationshipId ? [{ name: decodeXmlText(name), relationshipId }] : []
  })
}

function sheetPath(sheetIndex: number): string {
  return `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
}

function worksheetPathsBySheetName(zip: XlsxZipEntries, sheetNames: readonly string[]): Map<string, string> {
  const paths = new Map<string, string>()
  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const worksheetRelationshipsById = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type === worksheetRelationshipType || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))]),
  )
  readWorkbookSheetEntries(getZipText(zip, workbookPath)).forEach((entry) => {
    const worksheetPath = worksheetRelationshipsById.get(entry.relationshipId)
    if (worksheetPath) {
      paths.set(entry.name, worksheetPath)
    }
  })
  sheetNames.forEach((sheetName, sheetIndex) => {
    if (!paths.has(sheetName)) {
      paths.set(sheetName, sheetPath(sheetIndex))
    }
  })
  return paths
}

function readElementXml(xml: string | null, localName: 'bookViews' | 'sheetViews'): string | undefined {
  if (!xml) {
    return undefined
  }
  const elementXml = new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${localName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'u').exec(xml)?.[0]
  return elementXml ? addInheritedNamespaceDeclarations(elementXml, xml) : undefined
}

function addInheritedNamespaceDeclarations(elementXml: string, documentXml: string): string {
  const prefixes = prefixedNamesUsedByXml(elementXml)
  if (prefixes.size === 0) {
    return elementXml
  }
  const rootAttributes = /<[A-Za-z_][\w.-]*\b([^>]*)>/u.exec(documentXml)?.[1] ?? ''
  const namespaceAttributes = new Map(
    [...rootAttributes.matchAll(/\s+xmlns:([A-Za-z_][\w.-]*)=(["'])([\s\S]*?)\2/gu)].map((match) => [
      match[1] ?? '',
      `xmlns:${match[1] ?? ''}=${match[2] ?? '"'}${match[3] ?? ''}${match[2] ?? '"'}`,
    ]),
  )
  const inheritedAttributes = [...prefixes]
    .filter((prefix) => !hasNamespaceDeclaration(elementXml, prefix))
    .flatMap((prefix) => {
      const attribute = namespaceAttributes.get(prefix)
      return attribute ? [attribute] : []
    })
  const ignorablePrefixes = inheritedIgnorablePrefixes(rootAttributes, prefixes)
  if (ignorablePrefixes.length > 0) {
    const mcNamespace = namespaceAttributes.get('mc')
    if (mcNamespace && !hasNamespaceDeclaration(elementXml, 'mc')) {
      inheritedAttributes.push(mcNamespace)
    }
    if (!/\smc:Ignorable=(["'])/u.test(elementXml)) {
      inheritedAttributes.push(`mc:Ignorable="${ignorablePrefixes.join(' ')}"`)
    }
  }
  if (inheritedAttributes.length === 0) {
    return elementXml
  }
  return elementXml.replace(/<([A-Za-z_][\w.-]*)(\s|>|\/>)/u, `<$1 ${inheritedAttributes.join(' ')}$2`)
}

function inheritedIgnorablePrefixes(rootAttributes: string, prefixes: ReadonlySet<string>): string[] {
  const ignorable = /\smc:Ignorable=(["'])([\s\S]*?)\1/u.exec(rootAttributes)?.[2]
  if (!ignorable) {
    return []
  }
  return ignorable.split(/\s+/u).filter((prefix) => prefixes.has(prefix))
}

function prefixedNamesUsedByXml(xml: string): Set<string> {
  const prefixes = new Set<string>()
  for (const match of xml.matchAll(/<\/?([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*/gu)) {
    prefixes.add(match[1] ?? '')
  }
  for (const match of xml.matchAll(/\s([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*=/gu)) {
    prefixes.add(match[1] ?? '')
  }
  prefixes.delete('')
  prefixes.delete('xmlns')
  return prefixes
}

function hasNamespaceDeclaration(xml: string, prefix: string): boolean {
  return new RegExp(`\\sxmlns:${escapeRegExp(prefix)}=(["'])`, 'u').test(xml)
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function insertAfterElement(xml: string, elementName: string, insertedXml: string): string | null {
  const match = new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${elementName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'u').exec(xml)
  return match ? `${xml.slice(0, match.index + match[0].length)}${insertedXml}${xml.slice(match.index + match[0].length)}` : null
}

function replaceOrInsertWorkbookBookViews(workbookXml: string, bookViewsXml: string): string {
  const existing = readElementXml(workbookXml, 'bookViews')
  if (existing) {
    return workbookXml.replace(existing, bookViewsXml)
  }
  return (
    insertAfterElement(workbookXml, 'workbookPr', bookViewsXml) ??
    insertAfterElement(workbookXml, 'fileVersion', bookViewsXml) ??
    workbookXml.replace(/<workbook\b([^>]*)>/u, `<workbook$1>${bookViewsXml}`)
  )
}

function replaceOrInsertWorksheetSheetViews(sheetXml: string, sheetViewsXml: string): string {
  const existing = readElementXml(sheetXml, 'sheetViews')
  if (existing) {
    return sheetXml.replace(existing, sheetViewsXml)
  }
  return (
    insertAfterElement(sheetXml, 'sheetPr', sheetViewsXml) ??
    insertAfterElement(sheetXml, 'dimension', sheetViewsXml) ??
    sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1>${sheetViewsXml}`)
  )
}

export function readImportedWorkbookViewState(source: XlsxZipSource, sheetNames: readonly string[]): ImportedWorkbookViewState {
  const zip = readXlsxZipEntries(source)
  const sheetViewStateByName = new Map<string, NonNullable<SheetMetadataSnapshot['viewState']>>()
  const bookViewsXml = readElementXml(getZipText(zip, workbookPath), 'bookViews')
  const worksheetPaths = worksheetPathsBySheetName(zip, sheetNames)

  sheetNames.forEach((sheetName, sheetIndex) => {
    const worksheetPath = worksheetPaths.get(sheetName) ?? sheetPath(sheetIndex)
    const sheetViewsXml = readElementXml(getZipText(zip, worksheetPath), 'sheetViews')
    if (sheetViewsXml) {
      sheetViewStateByName.set(sheetName, { sheetViewsXml })
    }
  })

  return {
    ...(bookViewsXml ? { workbookViewState: { bookViewsXml } } : {}),
    sheetViewStateByName,
  }
}

export function addExportViewStateToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const workbookViewState = snapshot.workbook.metadata?.viewState
  const sheetsWithViewState = snapshot.sheets.filter((sheet) => sheet.metadata?.viewState)
  if (!workbookViewState && sheetsWithViewState.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  const workbookXml = getZipText(zip, workbookPath)
  if (workbookXml && workbookViewState?.bookViewsXml) {
    setZipText(zip, workbookPath, replaceOrInsertWorkbookBookViews(workbookXml, workbookViewState.bookViewsXml))
    changed = true
  }

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetViewsXml = sheet.metadata?.viewState?.sheetViewsXml
      if (!sheetViewsXml) {
        return
      }
      const path = sheetPath(sheetIndex)
      const sheetXml = getZipText(zip, path)
      if (!sheetXml) {
        return
      }
      setZipText(zip, path, replaceOrInsertWorksheetSheetViews(sheetXml, sheetViewsXml))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}
