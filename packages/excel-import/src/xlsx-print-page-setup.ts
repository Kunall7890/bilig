import { unzipSync, zipSync } from 'fflate'

import type { SheetMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import { parseRelationships, resolveTargetPath, setZipText } from './xlsx-pivot-artifacts.js'

const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'

const printPageSetupElementNames = ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks'] as const

const printPageSetupTailElements = [
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
] as const

type PrintPageSetupElementName = (typeof printPageSetupElementNames)[number]

type PrintPageSetupSnapshot = NonNullable<SheetMetadataSnapshot['printPageSetup']>

interface WorkbookSheetEntry {
  readonly name: string
  readonly relationshipId: string
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

function elementPattern(elementName: PrintPageSetupElementName): RegExp {
  return new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${elementName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'gu')
}

function usedNamespacePrefixes(xml: string): Set<string> {
  const prefixes = new Set<string>()
  for (const match of xml.matchAll(/\b([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\b/gu)) {
    const prefix = match[1]
    if (prefix && prefix !== 'xml' && prefix !== 'xmlns') {
      prefixes.add(prefix)
    }
  }
  return prefixes
}

function worksheetNamespaceDeclaration(sheetXml: string, prefix: string): string | null {
  const worksheetOpening = /<worksheet\b([^>]*)>/u.exec(sheetXml)?.[1] ?? ''
  const declaration = new RegExp(`\\sxmlns:${prefix}=(["'])([\\s\\S]*?)\\1`, 'u').exec(worksheetOpening)
  return declaration ? `xmlns:${prefix}=${declaration[1]}${declaration[2] ?? ''}${declaration[1]}` : null
}

function addMissingNamespaceDeclarations(sheetXml: string, elementXml: string): string {
  const missingDeclarations = [...usedNamespacePrefixes(elementXml)].flatMap((prefix) => {
    if (new RegExp(`\\sxmlns:${prefix}=`, 'u').test(elementXml)) {
      return []
    }
    const declaration = worksheetNamespaceDeclaration(sheetXml, prefix)
    return declaration ? [declaration] : []
  })
  if (missingDeclarations.length === 0) {
    return elementXml
  }
  return elementXml.replace(
    /<((?:[A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*)\b([^>]*?)(\/?)>/u,
    (_match, tagName: string, attributes: string, selfClosing: string) =>
      `<${tagName}${attributes} ${missingDeclarations.join(' ')}${selfClosing}>`,
  )
}

function readElementXml(sheetXml: string, elementName: PrintPageSetupElementName): string | undefined {
  const match = elementPattern(elementName).exec(sheetXml)?.[0]
  return match ? addMissingNamespaceDeclarations(sheetXml, match) : undefined
}

function readSheetPrintPageSetup(sheetXml: string): PrintPageSetupSnapshot | undefined {
  const snapshot: PrintPageSetupSnapshot = {}
  const printOptionsXml = readElementXml(sheetXml, 'printOptions')
  const pageMarginsXml = readElementXml(sheetXml, 'pageMargins')
  const pageSetupXml = readElementXml(sheetXml, 'pageSetup')
  const headerFooterXml = readElementXml(sheetXml, 'headerFooter')
  const rowBreaksXml = readElementXml(sheetXml, 'rowBreaks')
  const colBreaksXml = readElementXml(sheetXml, 'colBreaks')
  if (printOptionsXml) {
    snapshot.printOptionsXml = printOptionsXml
  }
  if (pageMarginsXml) {
    snapshot.pageMarginsXml = pageMarginsXml
  }
  if (pageSetupXml) {
    snapshot.pageSetupXml = pageSetupXml
  }
  if (headerFooterXml) {
    snapshot.headerFooterXml = headerFooterXml
  }
  if (rowBreaksXml) {
    snapshot.rowBreaksXml = rowBreaksXml
  }
  if (colBreaksXml) {
    snapshot.colBreaksXml = colBreaksXml
  }
  return Object.keys(snapshot).length > 0 ? snapshot : undefined
}

function removePrintPageSetupElements(sheetXml: string): string {
  return printPageSetupElementNames.reduce((xml, elementName) => xml.replace(elementPattern(elementName), ''), sheetXml)
}

function printPageSetupXml(snapshot: PrintPageSetupSnapshot): string {
  return [
    snapshot.printOptionsXml,
    snapshot.pageMarginsXml,
    snapshot.pageSetupXml,
    snapshot.headerFooterXml,
    snapshot.rowBreaksXml,
    snapshot.colBreaksXml,
  ]
    .filter((element): element is string => Boolean(element))
    .join('')
}

function insertPrintPageSetupXml(sheetXml: string, snapshot: PrintPageSetupSnapshot): string {
  const payload = printPageSetupXml(snapshot)
  const withoutPrintPageSetup = removePrintPageSetupElements(sheetXml)
  if (payload.length === 0) {
    return withoutPrintPageSetup
  }
  let insertIndex = withoutPrintPageSetup.indexOf('</worksheet>')
  for (const elementName of printPageSetupTailElements) {
    const elementIndex = withoutPrintPageSetup.search(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return withoutPrintPageSetup
  }
  return `${withoutPrintPageSetup.slice(0, insertIndex)}${payload}${withoutPrintPageSetup.slice(insertIndex)}`
}

export function readImportedWorkbookPrintPageSetup(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, PrintPageSetupSnapshot> {
  const zip = readXlsxZipEntries(source)
  const printPageSetupBySheet = new Map<string, PrintPageSetupSnapshot>()
  const worksheetPaths = worksheetPathsBySheetName(zip, sheetNames)

  sheetNames.forEach((sheetName, sheetIndex) => {
    const worksheetPath = worksheetPaths.get(sheetName) ?? sheetPath(sheetIndex)
    const sheetXml = getZipText(zip, worksheetPath)
    if (!sheetXml) {
      return
    }
    const printPageSetup = readSheetPrintPageSetup(sheetXml)
    if (printPageSetup) {
      printPageSetupBySheet.set(sheetName, printPageSetup)
    }
  })

  return printPageSetupBySheet
}

export function addExportPrintPageSetupToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => sheet.metadata?.printPageSetup)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const printPageSetup = sheet.metadata?.printPageSetup
      if (!printPageSetup) {
        return
      }
      const path = sheetPath(sheetIndex)
      const sheetXml = getZipText(zip, path)
      if (!sheetXml) {
        return
      }
      const nextSheetXml = insertPrintPageSetupXml(sheetXml, printPageSetup)
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, path, nextSheetXml)
        changed = true
      }
    })

  return changed ? zipSync(zip) : bytes
}
