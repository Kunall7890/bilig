import * as XLSX from 'xlsx'
import { unzipSync, zipSync } from 'fflate'

import type { LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { escapeXmlAttribute, getZipText, setXmlAttribute, setZipText } from './xlsx-export-xml.js'
import { readImportedXlsxSourceCellPatches, type ImportedXlsxSourceCellPatch } from './xlsx-source-bytes.js'
import { readXmlAttribute, worksheetCellElementPattern, worksheetCellOpeningTagPattern } from './xlsx-style-xml.js'
import { workbookSheetPathEntriesFromSource } from './xlsx-workbook-sheet-paths.js'

type WorkbookSheetSnapshot = WorkbookSnapshot['sheets'][number]

interface WorksheetPatch {
  readonly literals: ReadonlyMap<string, LiteralInput>
  readonly formulaCaches: ReadonlyMap<string, LiteralInput | undefined>
}

const valueElementPattern = /<(?:[A-Za-z_][\w.-]*:)?v\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?v>)/u
const formulaElementPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?f>)/u

function escapeXmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function removeXmlAttribute(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${name}=(["'])[\\s\\S]*?\\1`, 'u'), '')
}

function setCellType(openingTag: string, type: string | null): string {
  return type === null ? removeXmlAttribute(openingTag, 't') : setXmlAttribute(openingTag, 't', type)
}

function cellParts(cellXml: string): { readonly tagName: string; readonly openingTag: string; readonly body: string } | null {
  const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
  const tagName = openingTag ? /^<([^\s/>]+)/u.exec(openingTag)?.[1] : undefined
  if (!openingTag || !tagName) {
    return null
  }
  if (openingTag.endsWith('/>')) {
    return { tagName, openingTag, body: '' }
  }
  const closingTag = `</${tagName}>`
  return {
    tagName,
    openingTag,
    body: cellXml.slice(openingTag.length, cellXml.endsWith(closingTag) ? -closingTag.length : undefined),
  }
}

function inlineStringBody(value: string): string {
  const preserveSpace = /^\s|\s$/u.test(value) ? ' xml:space="preserve"' : ''
  return `<is><t${preserveSpace}>${escapeXmlText(value)}</t></is>`
}

function literalCellBody(value: LiteralInput): { readonly type: string | null; readonly body: string } | null {
  if (value === null) {
    return { type: null, body: '' }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { type: null, body: `<v>${String(value)}</v>` } : null
  }
  if (typeof value === 'boolean') {
    return { type: 'b', body: `<v>${value ? '1' : '0'}</v>` }
  }
  return { type: 'inlineStr', body: inlineStringBody(value) }
}

function cachedValueBody(value: LiteralInput | undefined): { readonly type: string | null; readonly body: string } | null {
  if (value === undefined || value === null) {
    return { type: null, body: '' }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { type: null, body: `<v>${String(value)}</v>` } : null
  }
  if (typeof value === 'boolean') {
    return { type: 'b', body: `<v>${value ? '1' : '0'}</v>` }
  }
  return { type: 'str', body: `<v>${escapeXmlText(value)}</v>` }
}

function rewriteCellXml(cellXml: string, type: string | null, body: string): string | null {
  const parts = cellParts(cellXml)
  if (!parts) {
    return null
  }
  const nextOpeningTag = setCellType(parts.openingTag, type)
  return `${nextOpeningTag.endsWith('/>') ? nextOpeningTag.slice(0, -2) + '>' : nextOpeningTag}${body}</${parts.tagName}>`
}

function patchLiteralCellXml(cellXml: string, value: LiteralInput): string | null {
  const valueBody = literalCellBody(value)
  return valueBody ? rewriteCellXml(cellXml, valueBody.type, valueBody.body) : null
}

function insertCachedValueBody(body: string, valueBody: string): string {
  const withoutValue = body.replace(valueElementPattern, '')
  if (valueBody.length === 0) {
    return withoutValue
  }
  const formulaMatch = formulaElementPattern.exec(withoutValue)
  if (formulaMatch) {
    const insertAt = formulaMatch.index + formulaMatch[0].length
    return `${withoutValue.slice(0, insertAt)}${valueBody}${withoutValue.slice(insertAt)}`
  }
  const extensionIndex = withoutValue.search(/<(?:[A-Za-z_][\w.-]*:)?extLst\b/u)
  return extensionIndex >= 0
    ? `${withoutValue.slice(0, extensionIndex)}${valueBody}${withoutValue.slice(extensionIndex)}`
    : `${withoutValue}${valueBody}`
}

function patchFormulaCacheCellXml(cellXml: string, value: LiteralInput | undefined): string | null {
  const parts = cellParts(cellXml)
  const valueBody = cachedValueBody(value)
  if (!parts || !valueBody) {
    return null
  }
  const nextOpeningTag = setCellType(parts.openingTag, valueBody.type)
  return `${nextOpeningTag.endsWith('/>') ? nextOpeningTag.slice(0, -2) + '>' : nextOpeningTag}${insertCachedValueBody(
    parts.body,
    valueBody.body,
  )}</${parts.tagName}>`
}

function literalCellXml(address: string, value: LiteralInput): string | null {
  const valueBody = literalCellBody(value)
  return valueBody
    ? `<c r="${escapeXmlAttribute(address)}"${valueBody.type === null ? '' : ` t="${valueBody.type}"`}>${valueBody.body}</c>`
    : null
}

function rowNumberForAddress(address: string): number | null {
  const row = XLSX.utils.decode_cell(address).r + 1
  return Number.isSafeInteger(row) && row > 0 ? row : null
}

function insertCellIntoExistingRow(sheetXml: string, address: string, cellXml: string): string | null {
  const rowNumber = rowNumberForAddress(address)
  if (rowNumber === null) {
    return null
  }
  const rowPattern = new RegExp(
    `<((?:[A-Za-z_][\\w.-]*:)?row)\\b(?<attributes>[^>]*\\br=(["'])${rowNumber}\\3[^>]*)(?:\\/>|>(?<body>[\\s\\S]*?)<\\/\\1>)`,
    'u',
  )
  if (!rowPattern.test(sheetXml)) {
    return null
  }
  return sheetXml.replace(rowPattern, (_match: string, tagName: string, attributes: string, _quote: string, body = '') => {
    return `<${tagName}${attributes}>${body}${cellXml}</${tagName}>`
  })
}

function insertRowIntoSheetData(sheetXml: string, address: string, cellXml: string): string | null {
  const rowNumber = rowNumberForAddress(address)
  if (rowNumber === null) {
    return null
  }
  const rowXml = `<row r="${String(rowNumber)}">${cellXml}</row>`
  const nextRowPattern = /<((?:[A-Za-z_][\w.-]*:)?row)\b[^>]*\br=(["'])([0-9]+)\2[^>]*(?:\/>|>[\s\S]*?<\/\1>)/gu
  for (const match of sheetXml.matchAll(nextRowPattern)) {
    const existingRow = Number(match[3])
    if (Number.isSafeInteger(existingRow) && existingRow > rowNumber) {
      return `${sheetXml.slice(0, match.index)}${rowXml}${sheetXml.slice(match.index)}`
    }
  }
  return sheetXml.replace(/<\/((?:[A-Za-z_][\w.-]*:)?sheetData)>/u, `${rowXml}</$1>`)
}

function insertLiteralCell(sheetXml: string, address: string, value: LiteralInput): string | null {
  const cellXml = literalCellXml(address, value)
  return cellXml ? (insertCellIntoExistingRow(sheetXml, address, cellXml) ?? insertRowIntoSheetData(sheetXml, address, cellXml)) : null
}

function updateWorksheetDimension(sheetXml: string, addresses: Iterable<string>): string {
  const decoded = [...addresses].flatMap((address) => {
    try {
      return [XLSX.utils.decode_cell(address)]
    } catch {
      return []
    }
  })
  if (decoded.length === 0 || !/<(?:[A-Za-z_][\w.-]*:)?dimension\b/u.test(sheetXml)) {
    return sheetXml
  }
  return sheetXml.replace(/<((?:[A-Za-z_][\w.-]*:)?dimension)\b([^>]*)\/>/u, (match, tagName: string, attributes: string) => {
    const ref = readXmlAttribute(match, 'ref')
    let range: XLSX.Range | null = null
    try {
      range = ref ? XLSX.utils.decode_range(ref) : null
    } catch {
      range = null
    }
    for (const cell of decoded) {
      range = range
        ? {
            s: { r: Math.min(range.s.r, cell.r), c: Math.min(range.s.c, cell.c) },
            e: { r: Math.max(range.e.r, cell.r), c: Math.max(range.e.c, cell.c) },
          }
        : { s: cell, e: cell }
    }
    return range ? `<${tagName}${setXmlAttribute(attributes, 'ref', XLSX.utils.encode_range(range))}/>` : match
  })
}

function patchWorksheetXml(sheetXml: string, patch: WorksheetPatch): string | null {
  const patchedLiterals = new Set<string>()
  const patchedFormulaCaches = new Set<string>()
  let failed = false
  let nextXml = sheetXml.replace(worksheetCellElementPattern, (cellXml: string) => {
    const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
    const address = openingTag ? readXmlAttribute(openingTag, 'r') : null
    if (!address) {
      return cellXml
    }
    if (patch.literals.has(address)) {
      const patched = patchLiteralCellXml(cellXml, patch.literals.get(address) ?? null)
      if (!patched) {
        failed = true
        return cellXml
      }
      patchedLiterals.add(address)
      return patched
    }
    if (patch.formulaCaches.has(address)) {
      const patched = patchFormulaCacheCellXml(cellXml, patch.formulaCaches.get(address))
      if (!patched) {
        failed = true
        return cellXml
      }
      patchedFormulaCaches.add(address)
      return patched
    }
    return cellXml
  })
  if (failed) {
    return null
  }
  for (const [address, value] of patch.literals.entries()) {
    if (!patchedLiterals.has(address)) {
      const inserted = insertLiteralCell(nextXml, address, value)
      if (inserted === null) {
        return null
      }
      nextXml = inserted
    }
  }
  for (const address of patch.formulaCaches.keys()) {
    if (!patchedFormulaCaches.has(address)) {
      return null
    }
  }
  return updateWorksheetDimension(nextXml, patch.literals.keys())
}

function collectFormulaCaches(sheet: WorkbookSheetSnapshot): Map<string, LiteralInput | undefined> {
  const formulas = new Map<string, LiteralInput | undefined>()
  for (const cell of sheet.cells) {
    if (cell.formula !== undefined) {
      formulas.set(cell.address, cell.value)
    }
  }
  return formulas
}

function literalPatchesBySheet(patches: readonly ImportedXlsxSourceCellPatch[]): Map<string, Map<string, LiteralInput>> {
  const output = new Map<string, Map<string, LiteralInput>>()
  for (const patch of patches) {
    const sheetPatches = output.get(patch.sheetName) ?? new Map<string, LiteralInput>()
    sheetPatches.set(patch.address, patch.value)
    output.set(patch.sheetName, sheetPatches)
  }
  return output
}

function removeCalcChain(zip: Record<string, Uint8Array>): void {
  delete zip['xl/calcChain.xml']
  const contentTypesXml = getZipText(zip, '[Content_Types].xml')
  if (contentTypesXml) {
    setZipText(
      zip,
      '[Content_Types].xml',
      contentTypesXml.replace(/<Override\b(?=[^>]*\bPartName=(["'])\/xl\/calcChain\.xml\1)(?:[^>"']|"[^"]*"|'[^']*')*\/>/u, ''),
    )
  }
  const workbookRelationshipsXml = getZipText(zip, 'xl/_rels/workbook.xml.rels')
  if (workbookRelationshipsXml) {
    setZipText(
      zip,
      'xl/_rels/workbook.xml.rels',
      workbookRelationshipsXml.replace(
        /<Relationship\b(?=[^>]*(?:\bTarget=(["'])calcChain\.xml\1|\bType=(["'])http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain\2))(?:[^>"']|"[^"]*"|'[^']*')*\/>/u,
        '',
      ),
    )
  }
}

function ensureWorkbookRecalculation(zip: Record<string, Uint8Array>): boolean {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return false
  }
  const nextWorkbookXml = /<(?:[A-Za-z_][\w.-]*:)?calcPr\b/u.test(workbookXml)
    ? workbookXml.replace(
        /<((?:[A-Za-z_][\w.-]*:)?calcPr)\b([^>]*)(?:\/>|>[\s\S]*?<\/\1>)/u,
        (_match, tagName: string, attributes: string) => {
          const opening = `<${tagName}${attributes}/>`
          return setXmlAttribute(setXmlAttribute(setXmlAttribute(opening, 'calcMode', 'auto'), 'fullCalcOnLoad', '1'), 'forceFullCalc', '1')
        },
      )
    : workbookXml.replace(/<\/((?:[A-Za-z_][\w.-]*:)?workbook)>/u, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></$1>')
  setZipText(zip, 'xl/workbook.xml', nextWorkbookXml)
  return true
}

export function tryExportSourcePreservingXlsx(snapshot: WorkbookSnapshot, sourceBytes: Uint8Array): Uint8Array | null {
  const literalPatches = readImportedXlsxSourceCellPatches(snapshot)
  if (literalPatches.length === 0) {
    return null
  }
  const zip = unzipSync(sourceBytes)
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  const sheetPathsByName = new Map(
    workbookSheetPathEntriesFromSource(
      zip,
      orderedSheets.map((sheet) => sheet.name),
    ).map((entry) => [entry.name, entry.path]),
  )
  const literalsBySheet = literalPatchesBySheet(literalPatches)
  for (const sheet of orderedSheets) {
    const literals = literalsBySheet.get(sheet.name) ?? new Map<string, LiteralInput>()
    const formulaCaches = collectFormulaCaches(sheet)
    if (literals.size === 0 && formulaCaches.size === 0) {
      continue
    }
    const sheetPath = sheetPathsByName.get(sheet.name)
    const sheetXml = sheetPath ? getZipText(zip, sheetPath) : null
    if (!sheetPath || !sheetXml) {
      return null
    }
    const patchedXml = patchWorksheetXml(sheetXml, { literals, formulaCaches })
    if (patchedXml === null) {
      return null
    }
    setZipText(zip, sheetPath, patchedXml)
  }
  removeCalcChain(zip)
  if (!ensureWorkbookRecalculation(zip)) {
    return null
  }
  return zipSync(zip)
}
