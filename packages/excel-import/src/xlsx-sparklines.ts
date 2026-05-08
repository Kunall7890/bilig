import { strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSparklinesSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

const sparklineExtensionUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const extensionElementPattern = /<(?:[A-Za-z_][\w.-]*:)?ext\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?ext>)/gu
const extensionListElementPattern = /<(?:[A-Za-z_][\w.-]*:)?extLst\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?extLst>)/gu
const deterministicZipOptions = { mtime: new Date(1980, 0, 1, 0, 0, 0) } as const

function setZipText(zip: Record<string, Uint8Array>, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function extensionContainsSparklines(extensionXml: string): boolean {
  return extensionXml.includes(sparklineExtensionUri) || /<(?:[A-Za-z_][\w.-]*:)?sparklineGroups\b/u.test(extensionXml)
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

function firstNamespaceDeclaration(sheetXml: string, prefix: string): string | null {
  const worksheetDeclaration = worksheetNamespaceDeclaration(sheetXml, prefix)
  if (worksheetDeclaration) {
    return worksheetDeclaration
  }
  const declaration = new RegExp(`\\sxmlns:${prefix}=(["'])([\\s\\S]*?)\\1`, 'u').exec(sheetXml)
  return declaration ? `xmlns:${prefix}=${declaration[1]}${declaration[2] ?? ''}${declaration[1]}` : null
}

function addMissingNamespaceDeclarations(sheetXml: string, extensionXml: string): string {
  const missingDeclarations = [...usedNamespacePrefixes(extensionXml)].flatMap((prefix) => {
    if (new RegExp(`\\sxmlns:${prefix}=`, 'u').test(extensionXml)) {
      return []
    }
    const declaration = firstNamespaceDeclaration(sheetXml, prefix)
    return declaration ? [declaration] : []
  })
  if (missingDeclarations.length === 0) {
    return extensionXml
  }
  return extensionXml.replace(
    /<((?:[A-Za-z_][\w.-]*:)?ext)\b([^>]*?)(\/?)>/u,
    (_match, tagName: string, attributes: string, selfClosing: string) =>
      `<${tagName}${attributes} ${missingDeclarations.join(' ')}${selfClosing}>`,
  )
}

function readSparklineExtensionXml(sheetXml: string): string | undefined {
  extensionElementPattern.lastIndex = 0
  const extensions = [...sheetXml.matchAll(extensionElementPattern)]
    .map((match) => match[0])
    .filter(extensionContainsSparklines)
    .map((extensionXml) => addMissingNamespaceDeclarations(sheetXml, extensionXml))
  return extensions.length > 0 ? extensions.join('') : undefined
}

function removeSparklineExtensionsFromExtensionList(extensionListXml: string): string {
  extensionElementPattern.lastIndex = 0
  const nextXml = extensionListXml.replace(extensionElementPattern, (extensionXml) =>
    extensionContainsSparklines(extensionXml) ? '' : extensionXml,
  )
  return /<(?:[A-Za-z_][\w.-]*:)?ext\b/u.test(nextXml) ? nextXml : ''
}

function removeSparklineExtensionXml(sheetXml: string): string {
  extensionListElementPattern.lastIndex = 0
  return sheetXml.replace(extensionListElementPattern, (extensionListXml) => removeSparklineExtensionsFromExtensionList(extensionListXml))
}

function insertSparklineExtensionXml(sheetXml: string, sparklineXml: string): string {
  const withoutSparklineXml = removeSparklineExtensionXml(sheetXml)
  const extensionListMatch = /<(?:[A-Za-z_][\w.-]*:)?extLst\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?extLst>/u.exec(withoutSparklineXml)
  if (extensionListMatch) {
    const extensionListXml = extensionListMatch[0]
    const nextExtensionListXml = extensionListXml.replace(/<\/((?:[A-Za-z_][\w.-]*:)?extLst)>/u, `${sparklineXml}</$1>`)
    return `${withoutSparklineXml.slice(0, extensionListMatch.index)}${nextExtensionListXml}${withoutSparklineXml.slice(
      extensionListMatch.index + extensionListXml.length,
    )}`
  }

  const insertIndex = withoutSparklineXml.indexOf('</worksheet>')
  if (insertIndex < 0) {
    return withoutSparklineXml
  }
  return `${withoutSparklineXml.slice(0, insertIndex)}<extLst>${sparklineXml}</extLst>${withoutSparklineXml.slice(insertIndex)}`
}

export function readImportedWorkbookSparklines(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSparklinesSnapshot> {
  const zip = readXlsxZipEntries(source)
  const sparklinesBySheet = new Map<string, WorkbookSparklinesSnapshot>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !extensionContainsSparklines(sheetXml)) {
      return
    }
    const sparklineXml = readSparklineExtensionXml(sheetXml)
    if (sparklineXml) {
      sparklinesBySheet.set(sheetName, { xml: sparklineXml })
    }
  })

  return sparklinesBySheet
}

export function addExportSparklinesToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const zip = unzipSync(bytes)
  let changed = false

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const sparklineXml = sheet.metadata?.sparklines?.xml
      const nextSheetXml = sparklineXml ? insertSparklineExtensionXml(sheetXml, sparklineXml) : removeSparklineExtensionXml(sheetXml)
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, sheetPath, nextSheetXml)
        changed = true
      }
    })

  return changed ? zipSync(zip, deterministicZipOptions) : bytes
}
