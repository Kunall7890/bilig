import { unzipSync, zipSync } from 'fflate'
import { decodeCellAddress } from '@bilig/xlsx'
import type { SheetJsWorkSheet } from './xlsx-sheetjs-types.js'

import type { SheetMetadataSnapshot, WorkbookHyperlinkSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, setXmlAttribute, setZipText } from './xlsx-export-xml.js'
import { worksheetCellRecords } from './xlsx-worksheet-cells.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readNonEmptyString(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : undefined
}

function readHyperlinkTarget(link: Record<string, unknown>): string | undefined {
  const target = readNonEmptyString(link['Target'])
  if (target) {
    return target
  }
  const location = readNonEmptyString(link['location'])
  return location ? `#${location}` : undefined
}

function readCellDisplayText(cell: Record<string, unknown>): string | undefined {
  const formatted = readNonEmptyString(cell['w'])
  if (formatted) {
    return formatted
  }
  const raw = cell['v']
  if (typeof raw === 'string') {
    return readNonEmptyString(raw)
  }
  if (typeof raw === 'number' || typeof raw === 'boolean') {
    return String(raw)
  }
  return undefined
}

function readXmlAttribute(tag: string, attributeName: string): string | undefined {
  const match = new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(tag)
  return match?.[2]
}

export function readImportedSheetHyperlinks(sheetName: string, sheet: SheetJsWorkSheet): WorkbookHyperlinkSnapshot[] | undefined {
  const hyperlinks: WorkbookHyperlinkSnapshot[] = []
  for (const { address, cell } of worksheetCellRecords(sheet)) {
    const link = cell['l']
    if (!isRecord(link)) {
      continue
    }
    const target = readHyperlinkTarget(link)
    if (!target) {
      continue
    }
    const tooltip = readNonEmptyString(link['Tooltip'])
    const display = readNonEmptyString(link['display']) ?? readCellDisplayText(cell)
    hyperlinks.push({
      sheetName,
      address,
      target,
      ...(tooltip ? { tooltip } : {}),
      ...(display ? { display } : {}),
    })
  }
  return hyperlinks.length > 0
    ? hyperlinks.toSorted(
        (left, right) =>
          decodeCellAddress(left.address).r - decodeCellAddress(right.address).r || left.address.localeCompare(right.address),
      )
    : undefined
}

export function addExportHyperlinksToWorksheet(worksheet: SheetJsWorkSheet, sheet: WorkbookSnapshot['sheets'][number]): void {
  for (const hyperlink of sheet.metadata?.hyperlinks ?? []) {
    if (hyperlink.sheetName !== sheet.name || hyperlink.target.trim().length === 0) {
      continue
    }
    const existingCell: unknown = worksheet[hyperlink.address]
    const cell = isRecord(existingCell) ? existingCell : { t: 'z' }
    cell['l'] = {
      Target: hyperlink.target,
      ...(hyperlink.tooltip ? { Tooltip: hyperlink.tooltip } : {}),
      ...(hyperlink.display ? { display: hyperlink.display } : {}),
    }
    worksheet[hyperlink.address] = cell
  }
}

export function addExportHyperlinkDisplaysToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const sheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  const displayEntries = sheets.map((sheet) => {
    const displaysByAddress = new Map<string, string>()
    for (const hyperlink of sheet.metadata?.hyperlinks ?? []) {
      if (hyperlink.sheetName !== sheet.name || !hyperlink.display || hyperlink.display.trim().length === 0) {
        continue
      }
      displaysByAddress.set(hyperlink.address, hyperlink.display)
    }
    return displaysByAddress
  })
  if (!displayEntries.some((entry) => entry.size > 0)) {
    return bytes
  }
  const zip = unzipSync(bytes)
  let changed = false
  for (const [index, displaysByAddress] of displayEntries.entries()) {
    if (displaysByAddress.size === 0) {
      continue
    }
    const path = `xl/worksheets/sheet${String(index + 1)}.xml`
    const worksheetXml = getZipText(zip, path)
    if (!worksheetXml) {
      continue
    }
    const nextWorksheetXml = worksheetXml.replace(/<(?:[A-Za-z_][\w.-]*:)?hyperlink\b[^>]*\/?>/gu, (tag) => {
      const ref = readXmlAttribute(tag, 'ref')
      const display = ref ? displaysByAddress.get(ref) : undefined
      return display ? setXmlAttribute(tag, 'display', display) : tag
    })
    if (nextWorksheetXml !== worksheetXml) {
      setZipText(zip, path, nextWorksheetXml)
      changed = true
    }
  }
  return changed ? zipSync(zip) : bytes
}

export function hasExportHyperlinks(metadata: SheetMetadataSnapshot | undefined): boolean {
  return (metadata?.hyperlinks?.length ?? 0) > 0
}
