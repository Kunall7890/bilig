import { statSync } from 'node:fs'

import { translateFormulaReferences } from '@bilig/formula'

import { decodeCellAddress, encodeCellAddress } from './address.js'
import { createFileXlsxSourceReader } from './file-source.js'
import { decodeXmlText, getXmlElementText, readXmlAttribute } from './xml.js'
import { workbookSheetPathEntriesForSource } from './workbook-sheet-paths.js'
import { forEachInflatedXlsxZipEntryChunk, readXlsxZipEntriesLazyFromByteSource, type XlsxZipEntries } from './zip-reader.js'

export type XlsxFormulaCacheInspectionLimit = number | 'all'
export type XlsxFormulaCacheLiteral = string | number | boolean | null
export type XlsxFormulaCacheReadPhase = 'open-source' | 'workbook-metadata' | 'formula-cache-scan'
export type XlsxFormulaCacheReadErrorReason = 'invalid-or-zip64-xlsx' | 'worksheet-stream-unavailable'

export interface XlsxFormulaCacheCell {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: XlsxFormulaCacheLiteral
}

export interface XlsxFormulaCacheScanResult {
  readonly inputBytes: number
  readonly sheetNames: readonly string[]
  readonly formulaCellCount: number
  readonly cells: readonly XlsxFormulaCacheCell[]
}

export class XlsxFormulaCacheReadError extends Error {
  readonly reason: XlsxFormulaCacheReadErrorReason
  readonly inputBytes: number

  constructor(message: string, reason: XlsxFormulaCacheReadErrorReason, inputBytes: number) {
    super(message)
    this.name = 'XlsxFormulaCacheReadError'
    this.reason = reason
    this.inputBytes = inputBytes
  }
}

interface FormulaInfo {
  readonly formula: string | null
  readonly sharedFormulaIndex: string | null
}

interface SharedFormulaMaster {
  readonly formula: string
  readonly row: number
  readonly col: number
}

interface SharedStringReference {
  readonly kind: 'shared-string'
  readonly index: number
}

type PendingCellValue =
  | { readonly kind: 'empty' }
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'string'; readonly value: string }
  | { readonly kind: 'error'; readonly value: string }
  | SharedStringReference

interface PendingFormulaCacheCell {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: PendingCellValue
}

const textDecoder = new TextDecoder()
const formulaElementPattern = /<((?:[A-Za-z_][\w.-]*:)?f)\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/\1>)/u
const worksheetCellStartTagPattern = /<(?:[A-Za-z_][\w.-]*:)?c\b/u
const worksheetCellCarryLength = 128

export function readXlsxFormulaCacheCellsFromFile(
  inputPath: string,
  options: {
    readonly inspectLimit?: XlsxFormulaCacheInspectionLimit
    readonly onPhase?: (phase: XlsxFormulaCacheReadPhase) => void
  } = {},
): XlsxFormulaCacheScanResult {
  const inputBytes = statSync(inputPath).size
  const source = createFileXlsxSourceReader(inputPath)
  try {
    options.onPhase?.('open-source')
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      throw new XlsxFormulaCacheReadError(
        'XLSX formula cache scan requires a ZIP central directory it can read lazily',
        'invalid-or-zip64-xlsx',
        inputBytes,
      )
    }
    const sheetEntries = workbookSheetPathEntriesForSource(zip)
    options.onPhase?.('workbook-metadata')
    const scan = collectXlsxFormulaCacheCells(
      zip,
      sheetEntries,
      normalizeXlsxFormulaCacheInspectionLimit(options.inspectLimit ?? 'all'),
      inputBytes,
    )
    options.onPhase?.('formula-cache-scan')
    return {
      inputBytes,
      sheetNames: sheetEntries.map((sheet) => sheet.name),
      formulaCellCount: scan.formulaCellCount,
      cells: scan.cells,
    }
  } finally {
    source.release?.()
  }
}

export function normalizeXlsxFormulaCacheInspectionLimit(limit: XlsxFormulaCacheInspectionLimit): XlsxFormulaCacheInspectionLimit {
  if (limit === 'all') {
    return limit
  }
  if (Number.isInteger(limit) && limit > 0) {
    return limit
  }
  throw new Error(`Expected inspectLimit to be "all" or a positive integer, received: ${String(limit)}`)
}

function collectXlsxFormulaCacheCells(
  zip: XlsxZipEntries,
  sheetEntries: readonly { readonly name: string; readonly path: string }[],
  inspectionLimit: XlsxFormulaCacheInspectionLimit,
  inputBytes: number,
): { readonly formulaCellCount: number; readonly cells: readonly XlsxFormulaCacheCell[] } {
  const pendingCells: PendingFormulaCacheCell[] = []
  const sharedStringIndexes = new Set<number>()
  let formulaCellCount = 0
  let collectedCellCount = 0
  const shouldCollect = (): boolean => inspectionLimit === 'all' || collectedCellCount < inspectionLimit
  const markCollected = (): void => {
    collectedCellCount += 1
  }
  for (const sheet of sheetEntries) {
    const scan = collectXlsxFormulaCacheCellsForSheet(
      zip,
      sheet.name,
      sheet.path,
      shouldCollect,
      markCollected,
      sharedStringIndexes,
      inputBytes,
    )
    formulaCellCount += scan.formulaCellCount
    for (const cell of scan.cells) {
      pendingCells.push(cell)
    }
  }
  const sharedStrings = readTargetSharedStrings(zip, sharedStringIndexes)
  return {
    formulaCellCount,
    cells: pendingCells.map((cell) => {
      const output: {
        target: string
        formula: string
        cachedValue?: XlsxFormulaCacheLiteral
      } = {
        target: cell.target,
        formula: cell.formula,
      }
      if (cell.cachedValue !== undefined) {
        output.cachedValue = literalValueForPendingCacheInspection(cell.cachedValue, sharedStrings)
      }
      return output
    }),
  }
}

function collectXlsxFormulaCacheCellsForSheet(
  zip: XlsxZipEntries,
  sheetName: string,
  sheetPath: string,
  shouldCollect: () => boolean,
  markCollected: () => void,
  sharedStringIndexes: Set<number>,
  inputBytes: number,
): { readonly formulaCellCount: number; readonly cells: readonly PendingFormulaCacheCell[] } {
  const sharedFormulaMasters = new Map<string, SharedFormulaMaster>()
  const cells: PendingFormulaCacheCell[] = []
  let formulaCellCount = 0
  let buffer = ''
  const processBuffer = (final: boolean): void => {
    let cell = shiftWorksheetCellXml(buffer, final)
    while (cell.status === 'cell') {
      buffer = cell.buffer
      const { cellXml, openingTag } = cell
      const addressText = openingTag ? readXmlAttribute(openingTag, 'r') : null
      if (addressText) {
        let decodedAddress: ReturnType<typeof decodeCellAddress> | null
        try {
          decodedAddress = decodeCellAddress(addressText)
        } catch {
          decodedAddress = null
        }
        const formula = decodedAddress ? readFormulaInfo(cellXml) : null
        if (decodedAddress && formula) {
          formulaCellCount += 1
          if (formula.sharedFormulaIndex && formula.formula) {
            sharedFormulaMasters.set(formula.sharedFormulaIndex, {
              formula: formula.formula,
              row: decodedAddress.r,
              col: decodedAddress.c,
            })
          }
          if (shouldCollect()) {
            const formulaText = formulaSourceForCacheInspection(formula, sharedFormulaMasters, decodedAddress.r, decodedAddress.c)
            const cachedValue = cachedValueForFormulaCacheInspection(cellXml, openingTag, sharedStringIndexes)
            cells.push({
              target: formatQualifiedTarget(sheetName, encodeCellAddress(decodedAddress)),
              formula: formulaText.startsWith('=') ? formulaText : `=${formulaText}`,
              ...(cachedValue === undefined ? {} : { cachedValue }),
            })
            markCollected()
          }
        }
      }
      cell = shiftWorksheetCellXml(buffer, final)
    }
    buffer = cell.status === 'pending' ? cell.buffer : ''
  }
  const streamed = forEachInflatedXlsxZipEntryChunk(
    zip,
    sheetPath,
    (chunk) => {
      buffer += textDecoder.decode(chunk, { stream: true })
      processBuffer(false)
    },
    { chunkSize: 64 * 1024, forceStreamingInflate: true },
  )
  if (!streamed) {
    throw new XlsxFormulaCacheReadError(`Unable to stream worksheet XML for ${sheetName}`, 'worksheet-stream-unavailable', inputBytes)
  }
  buffer += textDecoder.decode()
  processBuffer(true)
  return { formulaCellCount, cells }
}

type WorksheetCellShift =
  | { readonly status: 'cell'; readonly cellXml: string; readonly openingTag: string; readonly buffer: string }
  | { readonly status: 'pending'; readonly buffer: string }
  | { readonly status: 'done' }

function shiftWorksheetCellXml(buffer: string, final: boolean): WorksheetCellShift {
  const startMatch = worksheetCellStartTagPattern.exec(buffer)
  if (!startMatch) {
    return {
      status: final ? 'done' : 'pending',
      buffer: final ? '' : buffer.slice(Math.max(0, buffer.length - worksheetCellCarryLength)),
    }
  }
  const startIndex = startMatch.index
  const openingEnd = findXmlTagEnd(buffer, startIndex)
  if (openingEnd < 0) {
    return { status: 'pending', buffer: buffer.slice(startIndex) }
  }
  const openingTag = buffer.slice(startIndex, openingEnd + 1)
  const endIndex = openingTag.endsWith('/>') ? openingEnd + 1 : findXmlElementEnd(buffer, openingEnd + 1, openingTag)
  if (endIndex < 0) {
    return { status: 'pending', buffer: buffer.slice(startIndex) }
  }
  return {
    status: 'cell',
    cellXml: buffer.slice(startIndex, endIndex),
    openingTag,
    buffer: buffer.slice(endIndex),
  }
}

function findXmlTagEnd(xml: string, startIndex: number): number {
  let quote: '"' | "'" | null = null
  for (let index = startIndex; index < xml.length; index += 1) {
    const char = xml[index]
    if (quote) {
      if (char === quote) {
        quote = null
      }
    } else if (char === '"' || char === "'") {
      quote = char
    } else if (char === '>') {
      return index
    }
  }
  return -1
}

function findXmlElementEnd(xml: string, startIndex: number, openingTag: string): number {
  const tagName = /^<((?:[A-Za-z_][\w.-]*:)?c)\b/u.exec(openingTag)?.[1]
  if (!tagName) {
    return -1
  }
  const closingPattern = new RegExp(`</${escapeRegExp(tagName)}\\s*>`, 'u')
  const relativeMatch = closingPattern.exec(xml.slice(startIndex))
  return relativeMatch ? startIndex + relativeMatch.index + relativeMatch[0].length : -1
}

function readFormulaInfo(cellXml: string): FormulaInfo | null {
  const formulaXml = formulaElementPattern.exec(cellXml)?.[0]
  if (!formulaXml) {
    return null
  }
  const openingEnd = formulaXml.indexOf('>')
  const openingTag = openingEnd >= 0 ? formulaXml.slice(0, openingEnd + 1) : formulaXml
  const sharedFormulaIndex = readXmlAttribute(openingTag, 'si')
  const formula =
    formulaXml.endsWith('/>') || openingEnd < 0
      ? null
      : decodeXmlText(formulaXml.slice(openingEnd + 1, formulaXml.replace(/<\/(?:[A-Za-z_][\w.-]*:)?f>\s*$/u, '').length))
  return {
    formula: formula && formula.trim().length > 0 ? formula : null,
    sharedFormulaIndex,
  }
}

function formulaSourceForCacheInspection(
  formula: FormulaInfo,
  sharedFormulaMasters: ReadonlyMap<string, SharedFormulaMaster>,
  row: number,
  col: number,
): string {
  if (formula.formula) {
    return formula.formula
  }
  if (!formula.sharedFormulaIndex) {
    return ''
  }
  const master = sharedFormulaMasters.get(formula.sharedFormulaIndex)
  return master ? translateFormulaReferences(master.formula, row - master.row, col - master.col) : ''
}

function cachedValueForFormulaCacheInspection(
  cellXml: string,
  openingTag: string,
  sharedStringIndexes: Set<number>,
): PendingCellValue | undefined {
  if (getXmlElementText(cellXml, 'v') === null && readXmlAttribute(openingTag, 't') !== 'inlineStr') {
    return undefined
  }
  const value = readCellValue(cellXml, openingTag)
  if (isSharedStringReference(value)) {
    sharedStringIndexes.add(value.index)
  }
  return value
}

function readCellValue(cellXml: string, openingTag: string): PendingCellValue {
  const type = readXmlAttribute(openingTag, 't')
  if (type === 'inlineStr') {
    return { kind: 'string', value: readTextRuns(cellXml) }
  }
  const rawValue = getXmlElementText(cellXml, 'v')
  if (rawValue === null) {
    return { kind: 'empty' }
  }
  if (type === 's') {
    const index = Number(rawValue)
    return Number.isSafeInteger(index) && index >= 0 ? { kind: 'shared-string', index } : { kind: 'empty' }
  }
  if (type === 'str') {
    return { kind: 'string', value: decodeXmlText(rawValue) }
  }
  if (type === 'b') {
    return { kind: 'boolean', value: rawValue === '1' || rawValue.toLowerCase() === 'true' }
  }
  if (type === 'e') {
    return { kind: 'error', value: normalizeErrorText(decodeXmlText(rawValue)) }
  }
  const numeric = Number(rawValue)
  return Number.isFinite(numeric) ? { kind: 'number', value: numeric } : { kind: 'string', value: decodeXmlText(rawValue) }
}

function literalValueForPendingCacheInspection(
  value: PendingCellValue,
  sharedStrings: ReadonlyMap<number, string>,
): XlsxFormulaCacheLiteral {
  switch (value.kind) {
    case 'shared-string':
      return sharedStrings.get(value.index) ?? ''
    case 'empty':
      return null
    case 'number':
      return value.value
    case 'boolean':
      return value.value
    case 'string':
      return value.value
    case 'error':
      return value.value
  }
}

function readTargetSharedStrings(zip: XlsxZipEntries, targetIndexes: ReadonlySet<number>): ReadonlyMap<number, string> {
  const values = new Map<number, string>()
  if (targetIndexes.size === 0) {
    return values
  }
  let buffer = ''
  let index = 0
  const processBuffer = (final: boolean): boolean => {
    const safeEnd = final ? buffer.length : Math.max(0, buffer.lastIndexOf('<si'))
    if (safeEnd === 0 && !final) {
      return true
    }
    const safeXml = buffer.slice(0, safeEnd)
    for (const match of safeXml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?si)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu)) {
      if (targetIndexes.has(index)) {
        values.set(index, readTextRuns(match[0]))
      }
      index += 1
    }
    buffer = buffer.slice(safeEnd)
    return values.size < targetIndexes.size
  }
  const streamed = forEachInflatedXlsxZipEntryChunk(
    zip,
    'xl/sharedStrings.xml',
    (chunk) => {
      buffer += textDecoder.decode(chunk, { stream: true })
      return processBuffer(false)
    },
    { chunkSize: 64 * 1024, forceStreamingInflate: true },
  )
  if (!streamed) {
    return values
  }
  buffer += textDecoder.decode()
  processBuffer(true)
  return values
}

function isSharedStringReference(value: PendingCellValue): value is SharedStringReference {
  return value.kind === 'shared-string'
}

function formatQualifiedTarget(sheetName: string, address: string): string {
  return `${quoteSheetName(sheetName)}!${address}`
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/u.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function normalizeErrorText(value: string): string {
  switch (value.toUpperCase()) {
    case '#DIV/0!':
      return '#DIV/0!'
    case '#REF!':
      return '#REF!'
    case '#VALUE!':
      return '#VALUE!'
    case '#NAME?':
      return '#NAME?'
    case '#N/A':
      return '#N/A'
    case '#NUM!':
      return '#NUM!'
    case '#FIELD!':
      return '#FIELD!'
    case '#NULL!':
      return '#NULL!'
    default:
      return '#VALUE!'
  }
}

function readTextRuns(xml: string): string {
  const runs = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)].map(
    (match) => decodeXmlText(match[1] ?? ''),
  )
  return runs.length > 0 ? runs.join('') : ''
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
