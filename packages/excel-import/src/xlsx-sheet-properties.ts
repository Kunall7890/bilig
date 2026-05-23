import { strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookSheetPrSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import { workbookSheetPathEntriesFromSource } from './xlsx-workbook-sheet-paths.js'

function setZipText(zip: XlsxZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function readSheetPrXml(sheetXml: string): string | null {
  return /<sheetPr\b[^>]*(?:\/>|>[\s\S]*?<\/sheetPr>)/u.exec(sheetXml)?.[0] ?? null
}

function removeTabColor(sheetPrXml: string): string {
  return sheetPrXml.replace(/<tabColor\b[^>]*(?:\/>|>[\s\S]*?<\/tabColor>)/gu, '')
}

function hasPreservedSheetPrPayload(sheetPrXml: string): boolean {
  const withoutOpeningAndClosing = sheetPrXml
    .replace(/^<sheetPr\b[^>]*>/u, '')
    .replace(/^<sheetPr\b[^>]*\/>$/u, '')
    .replace(/<\/sheetPr>$/u, '')
    .trim()
  const openingAttributes = /^<sheetPr\b([^>]*)/u.exec(sheetPrXml)?.[1]?.trim() ?? ''
  return openingAttributes.length > 0 || withoutOpeningAndClosing.length > 0
}

export function readImportedWorkbookSheetProperties(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetPrSnapshot> {
  const zip = readXlsxZipEntries(source)
  const propertiesBySheet = new Map<string, WorkbookSheetPrSnapshot>()

  workbookSheetPathEntriesFromSource(zip, sheetNames).forEach((sheet) => {
    const sheetXml = getZipText(zip, sheet.path)
    if (!sheetXml) {
      return
    }
    const sheetPrXml = readSheetPrXml(sheetXml)
    if (!sheetPrXml) {
      return
    }
    const preserved = removeTabColor(sheetPrXml)
    if (hasPreservedSheetPrPayload(preserved)) {
      propertiesBySheet.set(sheet.name, { xml: preserved })
    }
  })

  return propertiesBySheet
}

function tabColorXml(sheetPrXml: string | null): string | null {
  return sheetPrXml ? (/<tabColor\b[^>]*(?:\/>|>[\s\S]*?<\/tabColor>)/u.exec(sheetPrXml)?.[0] ?? null) : null
}

function mergeTabColor(sheetPrXml: string, tabColor: string | null): string {
  if (!tabColor) {
    return sheetPrXml
  }
  const withoutTabColor = removeTabColor(sheetPrXml)
  const selfClosingMatch = /^<sheetPr\b([^>]*)\/>$/u.exec(withoutTabColor)
  if (selfClosingMatch) {
    return `<sheetPr${selfClosingMatch[1] ?? ''}>${tabColor}</sheetPr>`
  }
  const expandedMatch = /^<sheetPr\b([^>]*)>([\s\S]*?)<\/sheetPr>$/u.exec(withoutTabColor)
  if (!expandedMatch) {
    return withoutTabColor
  }
  return `<sheetPr${expandedMatch[1] ?? ''}>${tabColor}${expandedMatch[2] ?? ''}</sheetPr>`
}

function insertSheetPr(sheetXml: string, sheetPrXml: string): string {
  const existing = readSheetPrXml(sheetXml)
  if (existing) {
    return sheetXml.replace(existing, sheetPrXml)
  }
  return sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1>${sheetPrXml}`)
}

export function addExportWorksheetPropertiesToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => sheet.metadata?.sheetPr)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPr = sheet.metadata?.sheetPr
      if (!sheetPr) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const nextSheetPr = mergeTabColor(sheetPr.xml, tabColorXml(readSheetPrXml(sheetXml)))
      setZipText(zip, sheetPath, insertSheetPr(sheetXml, nextSheetPr))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}
