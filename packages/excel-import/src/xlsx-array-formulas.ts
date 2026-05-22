import { unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { readRuntimeImage } from '@bilig/core'
import type {
  CellValue,
  WorkbookSheetArrayFormulaSnapshot,
  WorkbookSheetArrayFormulasSnapshot,
  WorkbookSnapshot,
  WorkbookSpillSnapshot,
} from '@bilig/protocol'
import { ValueTag, formatErrorCode } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import {
  addContentTypeOverride,
  buildRelationshipsXml,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  setZipText,
} from './xlsx-pivot-artifacts.js'
import { addMissingCellsToSheetXml } from './xlsx-cell-insertion.js'
import { encodeFormulaForXlsx } from './xlsx-formula-translation.js'

const cellElementPattern = /<c\b(?<attributes>[^>]*)>(?<body>[\s\S]*?)<\/c>/gu
const formulaElementPattern = /<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>)/u
const formulaElementGlobalPattern = /<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/gu
const valueElementPattern = /<v\b[^>]*>[\s\S]*?<\/v>/u
const sheetMetadataRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata'
const sheetMetadataContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'
const dynamicArrayMetadataXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" ' +
  'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" ' +
  'merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">' +
  '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function isWorksheetCellAddress(value: string): boolean {
  return /^[A-Z]{1,3}[1-9][0-9]*$/u.test(value)
}

function cellColumnIndex(address: string): number {
  try {
    return XLSX.utils.decode_cell(address).c
  } catch {
    return Number.MAX_SAFE_INTEGER
  }
}

function cellRowNumber(address: string): number | null {
  const match = /^[A-Z]+([1-9][0-9]*)$/u.exec(address)
  return match ? Number(match[1]) : null
}

function decodeArrayFormulaRange(value: string): XLSX.Range | undefined {
  try {
    return XLSX.utils.decode_range(value)
  } catch {
    return undefined
  }
}

export function readImportedArrayFormulaSpills(sheetName: string, sheet: XLSX.WorkSheet): WorkbookSpillSnapshot[] | undefined {
  const spills: WorkbookSpillSnapshot[] = []
  for (const address in sheet) {
    const cell: unknown = sheet[address]
    if (!isWorksheetCellAddress(address) || !isRecord(cell)) {
      continue
    }
    const formula = cell['f']
    const arrayRangeText = cell['F']
    if (typeof formula !== 'string' || formula.trim().length === 0 || typeof arrayRangeText !== 'string') {
      continue
    }
    const range = decodeArrayFormulaRange(arrayRangeText.trim())
    if (!range) {
      continue
    }
    const owner = XLSX.utils.decode_cell(address)
    if (range.s.r !== owner.r || range.s.c !== owner.c) {
      continue
    }
    const rows = range.e.r - range.s.r + 1
    const cols = range.e.c - range.s.c + 1
    spills.push({
      sheetName,
      address: XLSX.utils.encode_cell(range.s),
      rows,
      cols,
    })
  }
  return spills.length > 0 ? spills : undefined
}

function readArrayFormulaSnapshots(sheetXml: string | null): WorkbookSheetArrayFormulaSnapshot[] {
  if (!sheetXml) {
    return []
  }
  if (!sheetXml.includes('t="array"') && !sheetXml.includes("t='array'")) {
    return []
  }
  const formulas: WorkbookSheetArrayFormulaSnapshot[] = []
  cellElementPattern.lastIndex = 0
  for (const match of sheetXml.matchAll(cellElementPattern)) {
    const attributes = match.groups?.['attributes'] ?? ''
    const body = match.groups?.['body'] ?? ''
    const address = readAttribute(attributes, 'r')
    formulaElementGlobalPattern.lastIndex = 0
    const formulaXml = [...body.matchAll(formulaElementGlobalPattern)].find((formulaMatch) => {
      return readAttribute(formulaMatch[0], 't') === 'array'
    })?.[0]
    if (address && formulaXml) {
      formulas.push({ address, formulaXml })
    }
  }
  return formulas
}

function insertFormulaIntoCellBody(body: string, formulaXml: string): string {
  if (formulaElementPattern.test(body)) {
    return body.replace(formulaElementPattern, formulaXml)
  }
  const valueIndex = body.search(/<(?:v|is|extLst)\b/u)
  return valueIndex >= 0 ? `${body.slice(0, valueIndex)}${formulaXml}${body.slice(valueIndex)}` : `${formulaXml}${body}`
}

function upsertFormulaInExistingCell(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string | null {
  const addressPattern = escapeRegExp(formula.address)
  const cellPattern = new RegExp(`<c\\b(?<attributes>[^>]*\\br=(["'])${addressPattern}\\2[^>]*)>(?<body>[\\s\\S]*?)<\\/c>`, 'u')
  if (!cellPattern.test(sheetXml)) {
    return null
  }
  return sheetXml.replace(cellPattern, (_cellXml: string, attributes: string, _quote: string, body: string) => {
    return `<c${attributes}>${insertFormulaIntoCellBody(body, formula.formulaXml)}</c>`
  })
}

function insertFormulaCellIntoRow(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string | null {
  const rowNumber = cellRowNumber(formula.address)
  if (rowNumber === null) {
    return null
  }
  const rowPattern = new RegExp(`<row\\b(?<attributes>[^>]*\\br=(["'])${String(rowNumber)}\\2[^>]*)>(?<body>[\\s\\S]*?)<\\/row>`, 'u')
  if (!rowPattern.test(sheetXml)) {
    return null
  }
  const nextCellXml = `<c r="${escapeXml(formula.address)}">${formula.formulaXml}</c>`
  return sheetXml.replace(rowPattern, (_rowXml: string, attributes: string, _quote: string, body: string) => {
    const targetColumn = cellColumnIndex(formula.address)
    let insertIndex = body.length
    for (const match of body.matchAll(/<c\b(?<attributes>[^>]*)>(?:[\s\S]*?)<\/c>/gu)) {
      const address = readAttribute(match.groups?.['attributes'] ?? '', 'r')
      if (address && cellColumnIndex(address) > targetColumn) {
        insertIndex = match.index
        break
      }
    }
    return `<row${attributes}>${body.slice(0, insertIndex)}${nextCellXml}${body.slice(insertIndex)}</row>`
  })
}

function insertFormulaCellIntoSheetData(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string {
  return addMissingCellsToSheetXml(sheetXml, [
    {
      address: formula.address,
      xml: `<c r="${escapeXml(formula.address)}">${formula.formulaXml}</c>`,
    },
  ])
}

function insertCellIntoRow(sheetXml: string, address: string, cellXml: string): string | null {
  const rowNumber = cellRowNumber(address)
  if (rowNumber === null) {
    return null
  }
  const rowPattern = new RegExp(`<row\\b(?<attributes>[^>]*\\br=(["'])${String(rowNumber)}\\2[^>]*)>(?<body>[\\s\\S]*?)<\\/row>`, 'u')
  if (!rowPattern.test(sheetXml)) {
    return null
  }
  return sheetXml.replace(rowPattern, (_rowXml: string, attributes: string, _quote: string, body: string) => {
    const targetColumn = cellColumnIndex(address)
    let insertIndex = body.length
    for (const match of body.matchAll(/<c\b(?<attributes>[^>]*)>(?:[\s\S]*?)<\/c>/gu)) {
      const existingAddress = readAttribute(match.groups?.['attributes'] ?? '', 'r')
      if (existingAddress && cellColumnIndex(existingAddress) > targetColumn) {
        insertIndex = match.index
        break
      }
    }
    return `<row${attributes}>${body.slice(0, insertIndex)}${cellXml}${body.slice(insertIndex)}</row>`
  })
}

function insertCellIntoSheetData(sheetXml: string, address: string, cellXml: string): string {
  return addMissingCellsToSheetXml(sheetXml, [{ address, xml: cellXml }])
}

function upsertArrayFormula(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string {
  return (
    upsertFormulaInExistingCell(sheetXml, formula) ??
    insertFormulaCellIntoRow(sheetXml, formula) ??
    insertFormulaCellIntoSheetData(sheetXml, formula)
  )
}

export function readImportedWorkbookArrayFormulas(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetArrayFormulasSnapshot> {
  const zip = readXlsxZipEntries(source)
  const formulasBySheet = new Map<string, WorkbookSheetArrayFormulasSnapshot>()
  sheetNames.forEach((sheetName, sheetIndex) => {
    const formulas = readArrayFormulaSnapshots(getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`))
    if (formulas.length > 0) {
      formulasBySheet.set(sheetName, { formulas })
    }
  })
  return formulasBySheet
}

export function addExportArrayFormulasToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const sheetsWithArrayFormulas = snapshot.sheets.filter((sheet) => (sheet.metadata?.arrayFormulas?.formulas.length ?? 0) > 0)
  if (sheetsWithArrayFormulas.length === 0) {
    return bytes
  }

  const zip: XlsxZipEntries = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const arrayFormulas = sheet.metadata?.arrayFormulas?.formulas
      if (!arrayFormulas || arrayFormulas.length === 0) {
        return
      }
      const sheetPath = normalizeZipPath(`xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const nextSheetXml = arrayFormulas.reduce(upsertArrayFormula, sheetXml)
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, sheetPath, nextSheetXml)
        changed = true
      }
    })

  return changed ? zipSync(zip) : bytes
}

function runtimeCellValueKey(row: number, col: number): string {
  return `${String(row)}:${String(col)}`
}

function readRuntimeCellValuesBySheet(snapshot: WorkbookSnapshot): Map<string, Map<string, CellValue>> {
  const runtimeImage = readRuntimeImage(snapshot)
  const valuesBySheet = new Map<string, Map<string, CellValue>>()
  const addValue = (sheetName: string, row: number, col: number, value: CellValue): void => {
    let sheetValues = valuesBySheet.get(sheetName)
    if (!sheetValues) {
      sheetValues = new Map()
      valuesBySheet.set(sheetName, sheetValues)
    }
    sheetValues.set(runtimeCellValueKey(row, col), value)
  }
  for (const value of runtimeImage?.cellValues ?? []) {
    addValue(value.sheetName, value.row, value.col, value.value)
  }
  for (const value of runtimeImage?.formulaValues ?? []) {
    addValue(value.sheetName, value.row, value.col, value.value)
  }
  return valuesBySheet
}

function workbookHasNativeDynamicSpills(snapshot: WorkbookSnapshot): boolean {
  const spills = snapshot.workbook.metadata?.spills
  return spills !== undefined && spills.length > 0
}

function cellValueBody(value: CellValue | undefined): { readonly attributes: string; readonly body: string } | undefined {
  if (!value || value.tag === ValueTag.Empty) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Number:
      return { attributes: '', body: `<v>${escapeXml(String(value.value))}</v>` }
    case ValueTag.Boolean:
      return { attributes: ' t="b"', body: `<v>${value.value ? '1' : '0'}</v>` }
    case ValueTag.String:
      return { attributes: ' t="str"', body: `<v>${escapeXml(value.value)}</v>` }
    case ValueTag.Error:
      return { attributes: ' t="e"', body: `<v>${escapeXml(formatErrorCode(value.code))}</v>` }
  }
}

function setCellAttribute(attributes: string, name: string, value: string): string {
  const attribute = `${name}="${escapeXml(value)}"`
  const pattern = new RegExp(`\\s${name}=(["'])[\\s\\S]*?\\1`, 'u')
  return pattern.test(attributes) ? attributes.replace(pattern, ` ${attribute}`) : `${attributes} ${attribute}`
}

function removeCellAttribute(attributes: string, name: string): string {
  return attributes.replace(new RegExp(`\\s${name}=(["'])[\\s\\S]*?\\1`, 'u'), '')
}

function applyCachedValueTypeAttribute(attributes: string, valueXml: { readonly attributes: string; readonly body: string }): string {
  if (!valueXml.attributes) {
    return removeCellAttribute(attributes, 't')
  }
  return setCellAttribute(attributes, 't', valueXml.attributes.slice(' t="'.length, -1))
}

function upsertCellValueBody(body: string, valueBody: string | undefined): string {
  if (!valueBody) {
    return body
  }
  const withoutValue = body.replace(valueElementPattern, '')
  const extIndex = withoutValue.search(/<extLst\b/u)
  return extIndex >= 0 ? `${withoutValue.slice(0, extIndex)}${valueBody}${withoutValue.slice(extIndex)}` : `${withoutValue}${valueBody}`
}

function upsertDynamicArrayOwnerCell(
  sheetXml: string,
  address: string,
  formulaXml: string,
  cachedValue: CellValue | undefined,
  cellMetadataIndex: string | undefined,
): string {
  const addressPattern = escapeRegExp(address)
  const cellPattern = new RegExp(`<c\\b(?<attributes>[^>]*\\br=(["'])${addressPattern}\\2[^>]*)(?:\\/>|>(?<body>[\\s\\S]*?)<\\/c>)`, 'u')
  const valueXml = cellValueBody(cachedValue)
  if (!cellPattern.test(sheetXml)) {
    const attributes = cellMetadataIndex ? ` cm="${escapeXml(cellMetadataIndex)}"` : ''
    const cellXml = `<c r="${escapeXml(address)}"${attributes}${valueXml?.attributes ?? ''}>${formulaXml}${valueXml?.body ?? ''}</c>`
    return insertCellIntoRow(sheetXml, address, cellXml) ?? insertCellIntoSheetData(sheetXml, address, cellXml)
  }
  return sheetXml.replace(cellPattern, (_cellXml: string, attributes: string, _quote: string, body = '') => {
    const nextAttributes = cellMetadataIndex ? setCellAttribute(attributes, 'cm', cellMetadataIndex) : attributes
    const typedAttributes = valueXml ? applyCachedValueTypeAttribute(nextAttributes, valueXml) : nextAttributes
    const nextBody = upsertCellValueBody(insertFormulaIntoCellBody(body, formulaXml), valueXml?.body)
    return `<c${typedAttributes}>${nextBody}</c>`
  })
}

function upsertCachedCell(sheetXml: string, address: string, cachedValue: CellValue | undefined): string {
  const valueXml = cellValueBody(cachedValue)
  if (!valueXml) {
    return sheetXml
  }
  const addressPattern = escapeRegExp(address)
  const cellPattern = new RegExp(`<c\\b(?<attributes>[^>]*\\br=(["'])${addressPattern}\\2[^>]*)(?:\\/>|>(?<body>[\\s\\S]*?)<\\/c>)`, 'u')
  if (!cellPattern.test(sheetXml)) {
    const cellXml = `<c r="${escapeXml(address)}"${valueXml.attributes}>${valueXml.body}</c>`
    return insertCellIntoRow(sheetXml, address, cellXml) ?? insertCellIntoSheetData(sheetXml, address, cellXml)
  }
  return sheetXml.replace(cellPattern, (_cellXml: string, attributes: string, _quote: string, body = '') => {
    const typedAttributes = applyCachedValueTypeAttribute(attributes, valueXml)
    return `<c${typedAttributes}>${upsertCellValueBody(body, valueXml.body)}</c>`
  })
}

function expandRangeForAddress(range: XLSX.Range | null, address: string): XLSX.Range | null {
  try {
    const decoded = XLSX.utils.decode_cell(address)
    return range
      ? {
          s: { r: Math.min(range.s.r, decoded.r), c: Math.min(range.s.c, decoded.c) },
          e: { r: Math.max(range.e.r, decoded.r), c: Math.max(range.e.c, decoded.c) },
        }
      : { s: { r: decoded.r, c: decoded.c }, e: { r: decoded.r, c: decoded.c } }
  } catch {
    return range
  }
}

function worksheetDimensionRange(sheetXml: string): XLSX.Range | null {
  const match = /<dimension\b[^>]*\bref=(["'])([\s\S]*?)\1[^>]*\/?>/u.exec(sheetXml)
  if (!match?.[2]) {
    return null
  }
  try {
    return XLSX.utils.decode_range(match[2])
  } catch {
    return null
  }
}

function updateWorksheetDimension(sheetXml: string, addresses: readonly string[]): string {
  let range = worksheetDimensionRange(sheetXml)
  const before = range ? XLSX.utils.encode_range(range) : undefined
  for (const address of addresses) {
    range = expandRangeForAddress(range, address)
  }
  if (!range) {
    return sheetXml
  }
  const ref = escapeXml(XLSX.utils.encode_range(range))
  if (before === ref) {
    return sheetXml
  }
  if (/<dimension\b/u.test(sheetXml)) {
    return sheetXml.replace(/<dimension\b[^>]*\/?>/u, (dimension) =>
      /\bref=/u.test(dimension) ? dimension.replace(/\bref=(["'])[\s\S]*?\1/u, `ref="${ref}"`) : `<dimension ref="${ref}"/>`,
    )
  }
  return sheetXml.replace(/<worksheet\b[^>]*>/u, (openingTag) => `${openingTag}<dimension ref="${ref}"/>`)
}

function spillRange(spill: WorkbookSpillSnapshot): { readonly range: string; readonly addresses: readonly string[] } | undefined {
  if (spill.rows < 1 || spill.cols < 1) {
    return undefined
  }
  try {
    const owner = XLSX.utils.decode_cell(spill.address)
    const end = { r: owner.r + spill.rows - 1, c: owner.c + spill.cols - 1 }
    const addresses: string[] = []
    for (let row = owner.r; row <= end.r; row += 1) {
      for (let col = owner.c; col <= end.c; col += 1) {
        addresses.push(XLSX.utils.encode_cell({ r: row, c: col }))
      }
    }
    return {
      range: XLSX.utils.encode_range({ s: owner, e: end }),
      addresses,
    }
  } catch {
    return undefined
  }
}

function ensureDynamicArrayMetadata(zip: XlsxZipEntries): string | undefined {
  const metadataPath = 'xl/metadata.xml'
  const existing = getZipText(zip, metadataPath)
  if (existing) {
    return existing.includes('dynamicArrayProperties') ? '1' : undefined
  }
  setZipText(zip, metadataPath, dynamicArrayMetadataXml)
  const relsPath = 'xl/_rels/workbook.xml.rels'
  const relationships = parseRelationships(getZipText(zip, relsPath))
  if (!relationships.some((relationship) => relationship.type === sheetMetadataRelationshipType)) {
    relationships.push({
      id: nextRelationshipId(relationships),
      type: sheetMetadataRelationshipType,
      target: 'metadata.xml',
    })
    setZipText(zip, relsPath, buildRelationshipsXml(relationships))
  }
  const contentTypesXml = getZipText(zip, '[Content_Types].xml')
  if (contentTypesXml) {
    setZipText(zip, '[Content_Types].xml', addContentTypeOverride(contentTypesXml, '/xl/metadata.xml', sheetMetadataContentType))
  }
  return '1'
}

export function addExportNativeSpillsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!workbookHasNativeDynamicSpills(snapshot)) {
    return bytes
  }
  const valuesBySheet = readRuntimeCellValuesBySheet(snapshot)
  if (valuesBySheet.size === 0) {
    return bytes
  }

  const zip: XlsxZipEntries = unzipSync(bytes)
  const cellMetadataIndex = ensureDynamicArrayMetadata(zip)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetSpills = (snapshot.workbook.metadata?.spills ?? []).filter((spill) => spill.sheetName === sheet.name)
      if (sheetSpills.length === 0) {
        return
      }
      const values = valuesBySheet.get(sheet.name)
      if (!values) {
        return
      }
      const cellsByAddress = new Map(sheet.cells.map((cell) => [cell.address, cell]))
      const sheetPath = normalizeZipPath(`xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      let nextSheetXml = sheetXml
      const touchedAddresses: string[] = []
      for (const spill of sheetSpills) {
        const ownerCell = cellsByAddress.get(spill.address)
        const decodedSpill = spillRange(spill)
        if (!ownerCell?.formula || !decodedSpill) {
          continue
        }
        const owner = XLSX.utils.decode_cell(spill.address)
        const formulaXml = `<f t="array" ref="${escapeXml(decodedSpill.range)}">${escapeXml(
          encodeFormulaForXlsx(ownerCell.formula.replace(/^=/u, '')),
        )}</f>`
        nextSheetXml = upsertDynamicArrayOwnerCell(
          nextSheetXml,
          spill.address,
          formulaXml,
          values.get(runtimeCellValueKey(owner.r, owner.c)),
          cellMetadataIndex,
        )
        for (const address of decodedSpill.addresses) {
          const decoded = XLSX.utils.decode_cell(address)
          if (address !== spill.address) {
            nextSheetXml = upsertCachedCell(nextSheetXml, address, values.get(runtimeCellValueKey(decoded.r, decoded.c)))
          }
        }
        touchedAddresses.push(...decodedSpill.addresses)
      }
      nextSheetXml = updateWorksheetDimension(nextSheetXml, touchedAddresses)
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, sheetPath, nextSheetXml)
        changed = true
      }
    })

  return changed ? zipSync(zip) : bytes
}
