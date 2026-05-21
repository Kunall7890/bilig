import { createHash } from 'node:crypto'

import * as XLSX from 'xlsx'

import { importXlsx, type ImportedWorkbook } from '../packages/excel-import/src/index.js'
import { readImportedExternalWorkbookReferences } from '../packages/excel-import/src/xlsx-external-references.js'
import { readWorkbookSheets, readWorksheetPathsByRelationshipId } from '../packages/excel-import/src/xlsx-large-simple-workbook-metadata.js'
import { decodeCellAddress, encodeCellAddress } from '../packages/excel-import/src/xlsx-large-simple-xml-byte-utils.js'
import { decodeXmlText } from '../packages/excel-import/src/xlsx-large-simple-worksheet-stream-text.js'
import {
  getZipText,
  readXlsxZipEntries,
  readXlsxZipEntriesLazyFromByteSource,
  releaseInflatedLazyXlsxZipEntries,
  type XlsxZipByteSource,
} from '../packages/excel-import/src/xlsx-zip.js'
import { ErrorCode, ValueTag } from '../packages/protocol/src/enums.js'
import type { CellValue, WorkbookExternalWorkbookReferenceSnapshot, WorkbookSnapshot } from '../packages/protocol/src/types.js'
import type { FormulaOracle, PublicWorkbookCorpusCase, PublicWorkbookFeatureCounts } from './public-workbook-corpus-types.ts'
import {
  inspectXlsxWorkbookFootprintLowMemory,
  inspectXlsxWorkbookFootprintLowMemoryAsync,
  isZipWorkbook,
} from './public-workbook-corpus-xlsx-footprint.ts'
import {
  tryInspectLargeSimpleXlsxHeadless,
  type LargeSimpleXlsxHeadlessInspectResult,
} from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'

export interface WorkbookFootprint {
  readonly featureCounts: PublicWorkbookFeatureCounts
  readonly workbookMetadata: PublicWorkbookCorpusCase['workbookMetadata']
  readonly externalWorkbookReferences: readonly WorkbookExternalWorkbookReferenceSnapshot[]
  readonly largeSimpleXlsxImport?: {
    readonly eligible: boolean
    readonly blockers: readonly string[]
  }
}

type WorkbookSheetUsedRange = NonNullable<PublicWorkbookCorpusCase['workbookMetadata']['dimensions'][number]['usedRange']>

interface WorksheetCellEntry {
  readonly address: string
  readonly cell: Record<string, unknown>
  readonly row: number
  readonly column: number
}

const largeSimpleHeadlessFingerprintCellThreshold = 100_000

export function countWorkbookFeatures(snapshot: WorkbookSnapshot, warnings: readonly string[]): PublicWorkbookFeatureCounts {
  return {
    sheetCount: snapshot.sheets.length,
    cellCount: snapshot.sheets.reduce((sum, sheet) => sum + sheet.cells.length, 0),
    formulaCellCount: snapshot.sheets.reduce((sum, sheet) => sum + sheet.cells.filter((cell) => cell.formula !== undefined).length, 0),
    valueCellCount: snapshot.sheets.reduce((sum, sheet) => sum + sheet.cells.filter((cell) => cell.value !== undefined).length, 0),
    definedNameCount: snapshot.workbook.metadata?.definedNames?.length ?? 0,
    tableCount: snapshot.workbook.metadata?.tables?.length ?? 0,
    chartCount: snapshot.workbook.metadata?.charts?.length ?? 0,
    pivotCount: snapshot.workbook.metadata?.pivots?.length ?? 0,
    mergeCount: snapshot.sheets.reduce((sum, sheet) => sum + (sheet.metadata?.merges?.length ?? 0), 0),
    styleRangeCount: snapshot.sheets.reduce((sum, sheet) => sum + (sheet.metadata?.styleRanges?.length ?? 0), 0),
    conditionalFormatCount: snapshot.sheets.reduce((sum, sheet) => sum + (sheet.metadata?.conditionalFormats?.length ?? 0), 0),
    dataValidationCount: snapshot.sheets.reduce((sum, sheet) => sum + (sheet.metadata?.validations?.length ?? 0), 0),
    macroPayloadCount: snapshot.workbook.metadata?.macroPayloads?.length ?? 0,
    warningCount: warnings.length,
  }
}

export function countImportedWorkbookFeatures(imported: ImportedWorkbook): PublicWorkbookFeatureCounts {
  const stats = imported.stats
  if (!stats) {
    return countWorkbookFeatures(imported.snapshot, imported.warnings)
  }
  const metadata = imported.snapshot.workbook.metadata
  return {
    sheetCount: stats.sheetCount,
    cellCount: stats.cellCount,
    formulaCellCount: stats.formulaCellCount,
    valueCellCount: stats.valueCellCount,
    definedNameCount: stats.definedNameCount,
    tableCount: Math.max(stats.tableCount, metadata?.tables?.length ?? 0),
    chartCount: metadata?.charts?.length ?? 0,
    pivotCount: metadata?.pivots?.length ?? 0,
    mergeCount: stats.mergeCount,
    styleRangeCount: imported.snapshot.sheets.reduce((sum, sheet) => sum + (sheet.metadata?.styleRanges?.length ?? 0), 0),
    conditionalFormatCount: stats.conditionalFormatCount,
    dataValidationCount: stats.dataValidationCount,
    macroPayloadCount: metadata?.macroPayloads?.length ?? 0,
    warningCount: imported.warnings.length,
  }
}

export function workbookMetadata(snapshot: WorkbookSnapshot): PublicWorkbookCorpusCase['workbookMetadata'] {
  return {
    workbookName: snapshot.workbook.name,
    sheetNames: snapshot.sheets.toSorted((left, right) => left.order - right.order).map((sheet) => sheet.name),
    dimensions: snapshot.sheets
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => {
        let rowCount = 0
        let columnCount = 0
        let usedRange: WorkbookSheetUsedRange | null = null
        for (const cell of sheet.cells) {
          const row = cell.row ?? rowIndexFromAddress(cell.address)
          const column = cell.col ?? columnIndexFromAddress(cell.address)
          rowCount = Math.max(rowCount, row + 1)
          columnCount = Math.max(columnCount, column + 1)
          usedRange = expandUsedRange(usedRange, row, column)
        }
        return {
          sheetName: sheet.name,
          rowCount,
          columnCount,
          nonEmptyCellCount: sheet.cells.length,
          usedRange,
        }
      }),
  }
}

export function importedWorkbookMetadata(imported: ImportedWorkbook): PublicWorkbookCorpusCase['workbookMetadata'] {
  const stats = imported.stats
  if (!stats) {
    return workbookMetadata(imported.snapshot)
  }
  return {
    workbookName: imported.snapshot.workbook.name,
    sheetNames: imported.snapshot.sheets.toSorted((left, right) => left.order - right.order).map((sheet) => sheet.name),
    dimensions: stats.dimensions.map((dimension) => ({
      sheetName: dimension.sheetName,
      rowCount: dimension.rowCount,
      columnCount: dimension.columnCount,
      nonEmptyCellCount: dimension.nonEmptyCellCount,
      usedRange: dimension.usedRange,
    })),
  }
}

export function emptyFeatureCounts(): PublicWorkbookFeatureCounts {
  return {
    sheetCount: 0,
    cellCount: 0,
    formulaCellCount: 0,
    valueCellCount: 0,
    definedNameCount: 0,
    tableCount: 0,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: 0,
    styleRangeCount: 0,
    conditionalFormatCount: 0,
    dataValidationCount: 0,
    macroPayloadCount: 0,
    warningCount: 0,
  }
}

export function inspectWorkbookFootprint(bytes: Uint8Array, fileName: string): WorkbookFootprint {
  if (isOpenXmlWorkbookFileName(fileName) && isZipWorkbook(bytes)) {
    return inspectXlsxWorkbookFootprintLowMemory(bytes, fileName)
  }
  const workbook = XLSX.read(bytes, {
    type: 'array',
    cellFormula: true,
    cellText: false,
    cellDates: false,
    dense: false,
  })
  const featureCounts = emptyFeatureCounts()
  const dimensions: PublicWorkbookCorpusCase['workbookMetadata']['dimensions'] = []
  featureCounts.sheetCount = workbook.SheetNames.length
  featureCounts.definedNameCount = Array.isArray(workbook.Workbook?.Names) ? workbook.Workbook.Names.length : 0
  featureCounts.pivotCount = countRawPivotTableParts(bytes)
  const externalWorkbookReferences = [...readImportedExternalWorkbookReferences(bytes).values()]
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    let rowCount = 0
    let columnCount = 0
    let nonEmptyCellCount = 0
    let usedRange: WorkbookSheetUsedRange | null = null
    if (sheet) {
      for (const { cell, row, column } of worksheetCellEntries(sheet)) {
        rowCount = Math.max(rowCount, row + 1)
        columnCount = Math.max(columnCount, column + 1)
        nonEmptyCellCount += 1
        usedRange = expandUsedRange(usedRange, row, column)
        featureCounts.cellCount += 1
        if (typeof cell.f === 'string' && cell.f.trim().length > 0) {
          featureCounts.formulaCellCount += 1
        }
        if (cell.v !== undefined) {
          featureCounts.valueCellCount += 1
        }
      }
      featureCounts.mergeCount += Array.isArray(sheet['!merges']) ? sheet['!merges'].length : 0
    }
    dimensions.push({ sheetName, rowCount, columnCount, nonEmptyCellCount, usedRange })
  }
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: fileName.replace(/\.(xlsx|xlsm|csv)$/iu, '') || fileName,
      sheetNames: workbook.SheetNames,
      dimensions,
    },
    externalWorkbookReferences,
  }
}

export async function inspectWorkbookFootprintForWorker(bytes: Uint8Array, fileName: string): Promise<WorkbookFootprint> {
  if (isOpenXmlWorkbookFileName(fileName) && isZipWorkbook(bytes)) {
    return inspectXlsxWorkbookFootprintLowMemoryAsync(bytes, fileName)
  }
  return inspectWorkbookFootprint(bytes, fileName)
}

function isOpenXmlWorkbookFileName(fileName: string): boolean {
  return /\.(xlsx|xlsm|xltx|xltm)$/iu.test(fileName)
}

function countRawPivotTableParts(bytes: Uint8Array): number {
  try {
    return Object.keys(readXlsxZipEntries(bytes)).filter((path) => /^xl\/pivotTables\/pivotTable\d+\.xml$/iu.test(path)).length
  } catch {
    return 0
  }
}

export function fingerprintWorkbookBytes(bytes: Uint8Array, fileName: string): string {
  const headlessFingerprint =
    isOpenXmlWorkbookFileName(fileName) && isZipWorkbook(bytes)
      ? fingerprintLargeSimpleDataOnlyWorkbookSource(byteSourceFromBytes(bytes), fileName)
      : null
  if (headlessFingerprint) {
    return headlessFingerprint
  }
  const preflightFootprint = tryInspectWorkbookFootprint(bytes, fileName)
  if (preflightFootprint) {
    const footprintFingerprint = fingerprintFormulaFreeWorkbookFootprint(preflightFootprint)
    if (footprintFingerprint) {
      return footprintFingerprint
    }
  }
  const imported = importXlsx(bytes, fileName)
  const footprint = preflightFootprint ?? inspectWorkbookFootprint(bytes, fileName)
  const importedCounts = countWorkbookFeatures(imported.snapshot, imported.warnings)
  const counts = {
    ...importedCounts,
    pivotCount: Math.max(importedCounts.pivotCount, footprint.featureCounts.pivotCount),
  }
  const metadata = workbookMetadata(imported.snapshot)
  const formulaShapes = imported.snapshot.sheets.flatMap((sheet) =>
    sheet.cells
      .filter((cell) => cell.formula !== undefined)
      .map((cell) => `${sheet.name}:${cell.address}:${cell.formula ?? ''}`)
      .toSorted(),
  )
  return sha256HexSync(Buffer.from(JSON.stringify({ counts, metadata, formulaShapes })))
}

function tryInspectWorkbookFootprint(bytes: Uint8Array, fileName: string): WorkbookFootprint | null {
  try {
    return inspectWorkbookFootprint(bytes, fileName)
  } catch {
    return null
  }
}

export function fingerprintFormulaFreeWorkbookFootprint(footprint: WorkbookFootprint): string | null {
  if (footprint.featureCounts.formulaCellCount > 0 || footprint.externalWorkbookReferences.length > 0) {
    return null
  }
  return sha256HexSync(
    Buffer.from(
      JSON.stringify({
        counts: footprint.featureCounts,
        metadata: footprint.workbookMetadata,
        formulaShapes: [],
      }),
    ),
  )
}

export function fingerprintLargeSimpleDataOnlyWorkbookSource(source: XlsxZipByteSource, fileName: string): string | null {
  if (!isOpenXmlWorkbookFileName(fileName)) {
    return null
  }
  const zip = readXlsxZipEntriesLazyFromByteSource(source)
  if (!zip) {
    return null
  }
  const inspected = tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, zip, {
    minByteLength: 0,
    releaseZipSource: true,
  })
  if (!inspected || inspected.stats.cellCount <= largeSimpleHeadlessFingerprintCellThreshold || inspected.stats.formulaCellCount > 0) {
    return null
  }
  const counts = countLargeSimpleHeadlessFingerprintFeatures(inspected)
  const metadata: PublicWorkbookCorpusCase['workbookMetadata'] = {
    workbookName: inspected.workbookName,
    sheetNames: inspected.sheetNames,
    dimensions: inspected.stats.dimensions,
  }
  return sha256HexSync(Buffer.from(JSON.stringify({ counts, metadata, formulaShapes: [] })))
}

function byteSourceFromBytes(bytes: Uint8Array): XlsxZipByteSource {
  return {
    byteLength: bytes.byteLength,
    readRange: (start, end) => bytes.subarray(start, end),
  }
}

function countLargeSimpleHeadlessFingerprintFeatures(inspected: LargeSimpleXlsxHeadlessInspectResult): PublicWorkbookFeatureCounts {
  return {
    sheetCount: inspected.stats.sheetCount,
    cellCount: inspected.stats.cellCount,
    formulaCellCount: inspected.stats.formulaCellCount,
    valueCellCount: inspected.stats.valueCellCount,
    definedNameCount: inspected.stats.definedNameCount,
    tableCount: inspected.stats.tableCount,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: inspected.stats.mergeCount,
    styleRangeCount: 0,
    conditionalFormatCount: inspected.stats.conditionalFormatCount,
    dataValidationCount: inspected.stats.dataValidationCount ?? 0,
    macroPayloadCount: 0,
    warningCount: inspected.warnings.length,
  }
}

export function extractFormulaOracles(bytes: Uint8Array): FormulaOracle[] {
  const workbook = XLSX.read(bytes, { type: 'array', cellFormula: true, cellText: false, cellDates: false })
  const oracles: FormulaOracle[] = []
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet?.['!ref']) {
      continue
    }
    for (const { address, cell } of worksheetCellEntries(sheet)) {
      if (typeof cell['f'] !== 'string' || cell['v'] === undefined) {
        continue
      }
      const expected = cellValueFromXlsx(cell)
      if (expected) {
        oracles.push({ sheetName, address, expected })
      }
    }
  }
  return oracles
}

export function extractFormulaOraclesFromXlsxByteSource(source: XlsxZipByteSource, fileName: string): FormulaOracle[] | null {
  if (!isOpenXmlWorkbookFileName(fileName)) {
    return null
  }
  const zip = readXlsxZipEntriesLazyFromByteSource(source)
  if (!zip) {
    return null
  }
  try {
    const workbookXml = getZipText(zip, 'xl/workbook.xml')
    const workbookRelationshipsXml = getZipText(zip, 'xl/_rels/workbook.xml.rels')
    if (!workbookXml || !workbookRelationshipsXml) {
      return null
    }
    const workbookSheets = readWorkbookSheets(workbookXml)
    const worksheetPathsByRelationshipId = readWorksheetPathsByRelationshipId(workbookRelationshipsXml)
    if (workbookSheets.length === 0 || worksheetPathsByRelationshipId.size === 0) {
      return null
    }
    releaseInflatedLazyXlsxZipEntries(zip)
    const oracles: FormulaOracle[] = []
    for (const sheet of workbookSheets) {
      const worksheetPath = worksheetPathsByRelationshipId.get(sheet.relationshipId)
      if (!worksheetPath) {
        return null
      }
      const worksheetXml = getZipText(zip, worksheetPath)
      if (!worksheetXml) {
        return null
      }
      const sheetOracles = extractFormulaOraclesFromWorksheetXml(sheet.name, worksheetXml)
      releaseInflatedLazyXlsxZipEntries(zip)
      if (!sheetOracles) {
        return null
      }
      oracles.push(...sheetOracles)
    }
    return oracles
  } catch {
    return null
  } finally {
    releaseInflatedLazyXlsxZipEntries(zip)
  }
}

export function cellValuesMatchOracle(actual: CellValue, expected: CellValue): boolean {
  if (actual.tag !== expected.tag) {
    return false
  }
  if (actual.tag === ValueTag.Number && expected.tag === ValueTag.Number) {
    const scale = Math.max(1, Math.abs(actual.value), Math.abs(expected.value))
    return Math.abs(actual.value - expected.value) <= Math.max(1e-7, scale * 1e-12)
  }
  if (actual.tag === ValueTag.String && expected.tag === ValueTag.String) {
    return actual.value === expected.value
  }
  if (actual.tag === ValueTag.Boolean && expected.tag === ValueTag.Boolean) {
    return actual.value === expected.value
  }
  return true
}

export function isUnsupportedCycleOracleMismatch(actual: CellValue, expected: CellValue, inCycle: boolean): boolean {
  return inCycle && actual.tag === ValueTag.Error && actual.code === ErrorCode.Cycle && expected.tag !== ValueTag.Error
}

export function formatCellValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return '<empty>'
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return String(value.value)
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return `error:${String(value.code)}`
  }
}

function cellValueFromXlsx(cell: Record<string, unknown>): CellValue | null {
  const value = cell['v']
  switch (cell['t']) {
    case 'n':
      return typeof value === 'number' && Number.isFinite(value) ? { tag: ValueTag.Number, value } : null
    case 'b':
      return typeof value === 'boolean' ? { tag: ValueTag.Boolean, value } : null
    case 's':
    case 'str':
      return typeof value === 'string' ? { tag: ValueTag.String, value, stringId: 0 } : null
    case 'd':
    case 'e':
    case 'z':
      return null
    default:
      return null
  }
}

type XmlFormulaOracleCellValue =
  | { readonly kind: 'value'; readonly value: CellValue }
  | { readonly kind: 'skip' }
  | { readonly kind: 'unsupported' }

function extractFormulaOraclesFromWorksheetXml(sheetName: string, worksheetXml: string): FormulaOracle[] | null {
  const oracles: FormulaOracle[] = []
  for (const match of worksheetXml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?c)\b((?:[^>"']|"[^"]*"|'[^']*')*)>([\s\S]*?)<\/\1>/gu)) {
    const attributes = match[2] ?? ''
    const content = match[3] ?? ''
    if (!/<(?:[A-Za-z_][\w.-]*:)?f\b/u.test(content)) {
      continue
    }
    const rawAddress = readXmlAttribute(attributes, 'r')
    const decodedAddress = rawAddress ? decodeCellAddress(decodeXmlText(rawAddress)) : null
    if (!decodedAddress) {
      return null
    }
    const rawCachedValue = readXmlElementText(content, 'v')
    if (rawCachedValue === null) {
      continue
    }
    const cellValue = cellValueFromWorksheetFormulaCache(readXmlAttribute(attributes, 't'), rawCachedValue)
    if (cellValue.kind === 'unsupported') {
      return null
    }
    if (cellValue.kind === 'skip') {
      continue
    }
    oracles.push({
      sheetName,
      address: encodeCellAddress(decodedAddress.row, decodedAddress.column),
      expected: cellValue.value,
    })
  }
  return oracles
}

function cellValueFromWorksheetFormulaCache(type: string | null, rawCachedValue: string): XmlFormulaOracleCellValue {
  const text = decodeXmlText(rawCachedValue.trim())
  switch (type ?? 'n') {
    case 'n': {
      const value = Number(text)
      return Number.isFinite(value) ? { kind: 'value', value: { tag: ValueTag.Number, value } } : { kind: 'skip' }
    }
    case 'b':
      if (text === '1' || /^true$/iu.test(text)) {
        return { kind: 'value', value: { tag: ValueTag.Boolean, value: true } }
      }
      if (text === '0' || /^false$/iu.test(text)) {
        return { kind: 'value', value: { tag: ValueTag.Boolean, value: false } }
      }
      return { kind: 'skip' }
    case 'str':
      return { kind: 'value', value: { tag: ValueTag.String, value: decodeXmlText(rawCachedValue), stringId: 0 } }
    case 'd':
    case 'e':
    case 'z':
      return { kind: 'skip' }
    default:
      return { kind: 'unsupported' }
  }
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function readXmlElementText(xml: string, elementName: string): string | null {
  const pattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b(?:[^>"']|"[^"]*"|'[^']*')*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${elementName}>`,
    'u',
  )
  return pattern.exec(xml)?.[1] ?? null
}

function expandUsedRange(current: WorkbookSheetUsedRange | null, row: number, column: number): WorkbookSheetUsedRange {
  return current
    ? {
        startRow: Math.min(current.startRow, row),
        startColumn: Math.min(current.startColumn, column),
        endRow: Math.max(current.endRow, row),
        endColumn: Math.max(current.endColumn, column),
      }
    : { startRow: row, startColumn: column, endRow: row, endColumn: column }
}

function worksheetCellEntries(sheet: XLSX.WorkSheet): WorksheetCellEntry[] {
  const denseRows = (sheet as Record<string, unknown>)['!data']
  if (Array.isArray(denseRows)) {
    const denseEntries: WorksheetCellEntry[] = []
    denseRows.forEach((row, rowIndex) => {
      if (!Array.isArray(row)) {
        return
      }
      row.forEach((cell, columnIndex) => {
        if (!isRecord(cell)) {
          return
        }
        denseEntries.push({
          address: XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex }),
          cell,
          row: rowIndex,
          column: columnIndex,
        })
      })
    })
    return denseEntries
  }

  const entries: WorksheetCellEntry[] = []
  for (const [address, value] of Object.entries(sheet)) {
    if (!/^[A-Z]{1,3}[1-9][0-9]*$/u.test(address) || !isRecord(value)) {
      continue
    }
    const decoded = XLSX.utils.decode_cell(address)
    entries.push({
      address,
      cell: value,
      row: decoded.r,
      column: decoded.c,
    })
  }
  return entries.toSorted((left, right) => left.row - right.row || left.column - right.column)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function rowIndexFromAddress(address: string): number {
  const row = Number(/\d+$/u.exec(address)?.[0] ?? '1')
  return Number.isInteger(row) && row > 0 ? row - 1 : 0
}

function columnIndexFromAddress(address: string): number {
  const letters = /^[A-Z]+/iu.exec(address)?.[0].toUpperCase() ?? 'A'
  let column = 0
  for (const letter of letters) {
    column = column * 26 + letter.charCodeAt(0) - 64
  }
  return Math.max(0, column - 1)
}

export function sha256HexSync(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}
