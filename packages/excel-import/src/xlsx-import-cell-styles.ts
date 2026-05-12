import type * as XLSX from 'xlsx'

import type {
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellStyleAlignmentSnapshot,
  CellStyleBordersSnapshot,
  CellStyleFontSnapshot,
  CellStyleProtectionSnapshot,
  CellStyleRecord,
  CellVerticalAlignment,
} from '@bilig/protocol'
import { readImportedAlignmentBoolean, readImportedAlignmentNumber, toLiteralInput } from './workbook-import-helpers.js'
import { worksheetCellRecords } from './xlsx-worksheet-cells.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeRgbColor(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const normalized = value.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return `#${normalized.toLowerCase()}`
  }
  if (/^[0-9a-fA-F]{8}$/.test(normalized)) {
    return `#${normalized.slice(2).toLowerCase()}`
  }
  return null
}

function readRgbColor(value: unknown): string | null {
  if (!isRecord(value)) {
    return null
  }
  return normalizeRgbColor(value['rgb'])
}

function readFiniteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function readImportedNumberFormat(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  if (trimmed.length === 0 || trimmed === 'General') {
    return undefined
  }
  return trimmed
}

function hasImportableXlsxCellPayload(cell: Record<string, unknown>): boolean {
  const formula = cell['f']
  if (typeof formula === 'string' && formula.trim().length > 0) {
    return true
  }
  return toLiteralInput(cell['v']) !== undefined || readImportedNumberFormat(cell['z']) !== undefined
}

export function collectStyleCandidateAddresses(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
  maxCandidateCount: number,
): {
  addressesBySheet: Map<string, Set<string>>
  count: number
} {
  const addressesBySheet = new Map<string, Set<string>>()
  let count = 0
  for (const sheetName of sheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) {
      continue
    }
    const addresses = new Set<string>()
    for (const { address, cell } of worksheetCellRecords(sheet)) {
      if (!hasImportableXlsxCellPayload(cell)) {
        continue
      }
      addresses.add(address)
      count += 1
      if (count > maxCandidateCount) {
        return { addressesBySheet: new Map(), count }
      }
    }
    if (addresses.size > 0) {
      addressesBySheet.set(sheetName, addresses)
    }
  }
  return { addressesBySheet, count }
}

function readImportedFillStyle(style: Record<string, unknown>): CellStyleRecord['fill'] | undefined {
  const fill = isRecord(style['fill']) ? style['fill'] : style
  if (fill['patternType'] !== 'solid') {
    return undefined
  }
  const backgroundColor = readRgbColor(fill['fgColor']) ?? readRgbColor(fill['bgColor'])
  return backgroundColor ? { backgroundColor } : undefined
}

function readImportedFontStyle(style: Record<string, unknown>): CellStyleFontSnapshot | undefined {
  const fontRecord = isRecord(style['font']) ? style['font'] : null
  if (!fontRecord) {
    return undefined
  }
  const font: CellStyleFontSnapshot = {}
  const family = typeof fontRecord['name'] === 'string' ? fontRecord['name'].trim() : ''
  if (family.length > 0) {
    font.family = family
  }
  const size = readFiniteNumber(fontRecord['sz']) ?? readFiniteNumber(fontRecord['size'])
  if (size !== null && size > 0) {
    font.size = size
  }
  if (fontRecord['bold'] === true) {
    font.bold = true
  }
  if (fontRecord['italic'] === true) {
    font.italic = true
  }
  if (fontRecord['underline'] === true || typeof fontRecord['underline'] === 'string') {
    font.underline = true
  }
  const color = readRgbColor(fontRecord['color'])
  if (color) {
    font.color = color
  }
  return Object.keys(font).length > 0 ? font : undefined
}

function readHorizontalAlignment(value: unknown): CellHorizontalAlignment | undefined {
  switch (value) {
    case 'general':
    case 'left':
    case 'center':
    case 'right':
    case 'fill':
    case 'justify':
    case 'centerContinuous':
    case 'distributed':
      return value
    default:
      return undefined
  }
}

function readVerticalAlignment(value: unknown): CellVerticalAlignment | undefined {
  switch (value) {
    case 'top':
      return 'top'
    case 'center':
    case 'middle':
      return 'middle'
    case 'bottom':
    case 'justify':
    case 'distributed':
      return value
    default:
      return undefined
  }
}

function readImportedAlignmentStyle(style: Record<string, unknown>): CellStyleAlignmentSnapshot | undefined {
  const alignmentRecord = isRecord(style['alignment']) ? style['alignment'] : null
  if (!alignmentRecord) {
    return undefined
  }
  const horizontal = readHorizontalAlignment(alignmentRecord['horizontal'])
  const vertical = readVerticalAlignment(alignmentRecord['vertical'])
  const indent = readImportedAlignmentNumber(alignmentRecord['indent'])
  const readingOrder = readImportedAlignmentNumber(alignmentRecord['readingOrder'])
  const textRotation = readImportedAlignmentNumber(alignmentRecord['textRotation'])
  const alignment: CellStyleAlignmentSnapshot = {
    ...(horizontal ? { horizontal } : {}),
    ...(vertical ? { vertical } : {}),
    ...(readImportedAlignmentBoolean(alignmentRecord['wrapText']) === true ? { wrap: true } : {}),
    ...(indent !== null && indent >= 0 ? { indent } : {}),
    ...(readImportedAlignmentBoolean(alignmentRecord['shrinkToFit']) === true ? { shrinkToFit: true } : {}),
    ...(readingOrder !== null ? { readingOrder } : {}),
    ...(textRotation !== null ? { textRotation } : {}),
    ...(readImportedAlignmentBoolean(alignmentRecord['justifyLastLine']) === true ? { justifyLastLine: true } : {}),
  }
  return Object.keys(alignment).length > 0 ? alignment : undefined
}

function readBorderKind(value: unknown): { style: CellBorderStyle; weight: CellBorderWeight } | null {
  switch (value) {
    case 'hair':
    case 'thin':
      return { style: 'solid', weight: 'thin' }
    case 'medium':
      return { style: 'solid', weight: 'medium' }
    case 'thick':
      return { style: 'solid', weight: 'thick' }
    case 'dashed':
    case 'mediumDashed':
    case 'dashDot':
    case 'dashDotDot':
    case 'slantDashDot':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
      return { style: 'dashed', weight: value === 'dashed' ? 'thin' : 'medium' }
    case 'dotted':
      return { style: 'dotted', weight: 'thin' }
    case 'double':
      return { style: 'double', weight: 'medium' }
    default:
      return null
  }
}

function readImportedBorderSide(value: unknown): CellBorderSideSnapshot | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const borderKind = readBorderKind(value['style'])
  if (!borderKind) {
    return undefined
  }
  return {
    ...borderKind,
    color: readRgbColor(value['color']) ?? '#000000',
  }
}

function readImportedBorderStyle(style: Record<string, unknown>): CellStyleBordersSnapshot | undefined {
  const borderRecord = isRecord(style['border']) ? style['border'] : null
  if (!borderRecord) {
    return undefined
  }
  const top = readImportedBorderSide(borderRecord['top'])
  const right = readImportedBorderSide(borderRecord['right'])
  const bottom = readImportedBorderSide(borderRecord['bottom'])
  const left = readImportedBorderSide(borderRecord['left'])
  const borders: CellStyleBordersSnapshot = {
    ...(top ? { top } : {}),
    ...(right ? { right } : {}),
    ...(bottom ? { bottom } : {}),
    ...(left ? { left } : {}),
  }
  return Object.keys(borders).length > 0 ? borders : undefined
}

function readImportedProtectionStyle(style: Record<string, unknown>): CellStyleProtectionSnapshot | undefined {
  const protectionRecord = isRecord(style['protection']) ? style['protection'] : null
  if (!protectionRecord) {
    return undefined
  }
  return {
    ...(typeof protectionRecord['locked'] === 'boolean' ? { locked: protectionRecord['locked'] } : {}),
    ...(typeof protectionRecord['hidden'] === 'boolean' ? { hidden: protectionRecord['hidden'] } : {}),
  }
}

export function readImportedXlsxCellStyle(value: unknown): Omit<CellStyleRecord, 'id'> | null {
  if (!isRecord(value)) {
    return null
  }
  const fill = readImportedFillStyle(value)
  const font = readImportedFontStyle(value)
  const alignment = readImportedAlignmentStyle(value)
  const borders = readImportedBorderStyle(value)
  const protection = readImportedProtectionStyle(value)
  const style: Omit<CellStyleRecord, 'id'> = {
    ...(fill ? { fill } : {}),
    ...(font ? { font } : {}),
    ...(alignment ? { alignment } : {}),
    ...(borders ? { borders } : {}),
    ...(protection !== undefined ? { protection } : {}),
  }
  return Object.keys(style).length > 0 ? style : null
}

export function internImportedStyle(style: Omit<CellStyleRecord, 'id'>, catalog: Map<string, CellStyleRecord>): string {
  const key = JSON.stringify(style)
  const existing = catalog.get(key)
  if (existing) {
    return existing.id
  }
  const record: CellStyleRecord = {
    id: `xlsx-style-${catalog.size + 1}`,
    ...style,
  }
  catalog.set(key, record)
  return record.id
}
