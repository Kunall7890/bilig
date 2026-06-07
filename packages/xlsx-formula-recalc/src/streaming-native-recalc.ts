import { statSync } from 'node:fs'

import { parseFormula, translateFormulaReferences, type FormulaNode } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import {
  createFileXlsxSourceReader,
  decodeCellAddress,
  decodeCellRange,
  encodeCellAddress,
  exportXlsxSourceLiteralPatchesToFileAsync,
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  normalizeZipPath,
  readXlsxZipEntriesLazyFromByteSource,
  readXmlAttribute,
  workbookSheetPathEntriesForSource,
  worksheetCellElementPattern,
  worksheetCellOpeningTagPattern,
  type XlsxCellRange,
  type XlsxSourceLiteralPatch,
  type XlsxZipEntries,
} from '@bilig/xlsx'

import { evaluateStreamingNativeWasmFormulas, expandStreamingNativeFormulaDependencyRows } from './streaming-native-row-chain-wasm.js'

export type XlsxFormulaRecalcEngineMode = 'streaming-native' | 'workpaper'
export type XlsxFormulaRecalcEngine = 'auto' | XlsxFormulaRecalcEngineMode
export type XlsxFormulaRecalcFallbackPolicy = 'error' | 'workpaper'

export interface XlsxFormulaRecalcPhaseRss {
  readonly phase: string
  readonly rssBytes: number
}

export interface StreamingNativeFormulaCounts {
  readonly scannedFormulaCellCount: number
  readonly targetedFormulaCellCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patchedFormulaCacheCount: number
  readonly unsupportedFormulaCellCount: number
  readonly nativeKernelFormulaCellCount: number
  readonly nativeKernelBatchCount: number
}

export interface XlsxFormulaRecalcNativeDiagnostics {
  readonly engineMode: XlsxFormulaRecalcEngineMode
  readonly inputBytes: number
  readonly phaseRssPeaks: readonly XlsxFormulaRecalcPhaseRss[]
  readonly maxObservedRssBytes: number
  readonly maxRssBytes?: number
  readonly sheetCount: number
  readonly targetRowCount: number
  readonly editCount: number
  readonly readCount: number
  readonly formulaCounts: StreamingNativeFormulaCounts
  readonly patchedCacheCount: number
  readonly unsupportedReason?: string
}

export interface XlsxFormulaRecalcDiagnostics extends XlsxFormulaRecalcNativeDiagnostics {
  readonly externalWorkbookHydration?: unknown
}

export interface StreamingNativeXlsxFormulaRecalcOptions {
  readonly outputPath: string
  readonly edits?: readonly { readonly target: string; readonly value: LiteralInput }[]
  readonly reads?: readonly string[]
  readonly maxRssBytes?: number
}

export interface StreamingNativeXlsxFormulaRecalcResult {
  readonly bytesWritten: number
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, CellValue>>
  readonly changes: readonly []
  readonly diagnostics: XlsxFormulaRecalcNativeDiagnostics
}

interface QualifiedTarget {
  readonly source: string
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
}

interface QualifiedEdit extends QualifiedTarget {
  readonly value: LiteralInput
}

export interface SheetScanState {
  readonly sheetName: string
  readonly sheetPath: string
  readonly targetRows: Set<number>
  readonly rows: Map<number, Map<number, PendingCellValue>>
  readonly formulaCells: NativeFormulaCell[]
  readonly sharedFormulaMasters: Map<string, SharedFormulaMaster>
  readonly tableRelationshipIds: Set<string>
  scannedFormulaCellCount: number
}

export interface NativeFormulaCell {
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
  readonly formula: string | null
  readonly sharedFormulaIndex: string | null
}

interface SharedFormulaMaster {
  readonly formula: string
  readonly row: number
  readonly col: number
}

interface FormulaInfo {
  readonly formula: string | null
  readonly sharedFormulaIndex: string | null
  readonly sharedFormulaRef: string | null
}

export interface NativeTable {
  readonly name: string
  readonly displayName: string
  readonly sheetName: string
  readonly range: XlsxCellRange
  readonly headerRowCount: number
  readonly totalsRowCount: number
  readonly columns: readonly string[]
}

interface EvaluationContext {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly rowValues: Map<number, PendingCellValue>
  readonly sheetRows: ReadonlyMap<number, Map<number, PendingCellValue>>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
}

interface SharedStringReference {
  readonly kind: 'shared-string'
  readonly index: number
}

export type PendingCellValue = CellValue | SharedStringReference

class UnsupportedStreamingNativeFormulaError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'UnsupportedStreamingNativeFormulaError'
  }
}

export class StreamingNativeXlsxRecalcError extends Error {
  readonly diagnostics: XlsxFormulaRecalcNativeDiagnostics

  constructor(message: string, diagnostics: XlsxFormulaRecalcNativeDiagnostics) {
    super(message)
    this.name = 'StreamingNativeXlsxRecalcError'
    this.diagnostics = diagnostics
  }
}

const emptyCellValue: CellValue = Object.freeze({ tag: ValueTag.Empty })
const textDecoder = new TextDecoder()
const formulaElementPattern = /<((?:[A-Za-z_][\w.-]*:)?f)\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/\1>)/u
const tablePartPattern = /<(?:[A-Za-z_][\w.-]*:)?tablePart\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu
const relationshipPattern = /<Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu
const tableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'

export async function recalculateXlsxFileToFileStreamingNative(
  inputPath: string,
  options: StreamingNativeXlsxFormulaRecalcOptions,
): Promise<StreamingNativeXlsxFormulaRecalcResult> {
  const inputBytes = statSync(inputPath).size
  const phaseRssPeaks: XlsxFormulaRecalcPhaseRss[] = []
  let maxObservedRssBytes = 0
  const recordPhase = (phase: string): void => {
    const rssBytes = process.memoryUsage().rss
    maxObservedRssBytes = Math.max(maxObservedRssBytes, rssBytes)
    phaseRssPeaks.push({ phase, rssBytes })
    if (options.maxRssBytes !== undefined && rssBytes > options.maxRssBytes) {
      throw streamingError(`streaming-native exceeded maxRssBytes during ${phase}`, phaseRssPeaks, {
        inputBytes,
        maxObservedRssBytes,
        maxRssBytes: options.maxRssBytes,
      })
    }
  }

  const source = createFileXlsxSourceReader(inputPath)
  try {
    recordPhase('open-source')
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      throw streamingError('streaming-native requires a ZIP central directory it can read lazily', phaseRssPeaks, {
        inputBytes,
        maxObservedRssBytes,
        ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
        unsupportedReason: 'invalid-or-zip64-xlsx',
      })
    }
    const sheetEntries = workbookSheetPathEntriesForSource(zip)
    const sheetNames = sheetEntries.map((sheet) => sheet.name)
    const sheetPathsByName = new Map(sheetEntries.map((sheet) => [sheet.name, sheet.path]))
    const edits = (options.edits ?? []).map((edit) => Object.assign(parseQualifiedTarget(edit.target), { value: edit.value }))
    const reads = (options.reads ?? []).map(parseQualifiedTarget)
    validateTargets([...edits, ...reads], sheetPathsByName)
    const targetRowsBySheet = targetRowsForTargets([...edits, ...reads])
    recordPhase('workbook-metadata')

    const sheetScans = new Map<string, SheetScanState>()
    for (const sheet of sheetEntries) {
      const targetRows = targetRowsBySheet.get(sheet.name)
      if (!targetRows || targetRows.size === 0) {
        continue
      }
      const scan = scanWorksheet(zip, sheet.name, sheet.path, targetRows)
      sheetScans.set(sheet.name, scan)
    }
    recordPhase('worksheet-scan')

    const expandedDependencySheets = expandStreamingNativeFormulaDependencyRows({
      sheetScans,
      resolveFormulaSource,
    })
    for (const sheetName of expandedDependencySheets) {
      const scan = sheetScans.get(sheetName)
      const sheetPath = sheetPathsByName.get(sheetName)
      if (!scan || !sheetPath) {
        continue
      }
      const nextScan = scanWorksheet(zip, sheetName, sheetPath, scan.targetRows)
      const originalTargetRows = targetRowsBySheet.get(sheetName) ?? new Set<number>()
      nextScan.formulaCells.splice(
        0,
        nextScan.formulaCells.length,
        ...nextScan.formulaCells.filter((cell) => originalTargetRows.has(cell.row)),
      )
      sheetScans.set(sheetName, nextScan)
    }
    recordPhase('dependency-row-scan')

    const sharedStrings = readTargetSharedStrings(zip, collectSharedStringReferences(sheetScans))
    hydrateSharedStringReferences(sheetScans, sharedStrings)
    recordPhase('shared-strings')

    const tablesBySheet = readNativeTablesBySheet(zip, sheetScans)
    recordPhase('table-metadata')

    const patches: XlsxSourceLiteralPatch[] = []
    applyEdits(edits, sheetScans, patches)
    const formulaCounts = evaluateFormulaCells(sheetScans, tablesBySheet, patches)
    const readValues = readTargets(reads, sheetScans)
    const targetRowCount = [...targetRowsBySheet.values()].reduce((sum, rows) => sum + rows.size, 0)
    const editCount = edits.length
    const readCount = reads.length
    const patchedSheetNames = [...new Set(patches.map((patch) => patch.sheetName))]
    recordPhase('evaluate')

    sheetScans.clear()
    if (tablesBySheet instanceof Map) {
      tablesBySheet.clear()
    }
    targetRowsBySheet.clear()
    sheetPathsByName.clear()

    const output = await exportXlsxSourceLiteralPatchesToFileAsync({
      source,
      patches,
      sheetNames: patchedSheetNames,
      outputPath: options.outputPath,
    })
    recordPhase('write-output')

    const diagnostics: XlsxFormulaRecalcNativeDiagnostics = {
      engineMode: 'streaming-native',
      inputBytes,
      phaseRssPeaks,
      maxObservedRssBytes,
      ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
      sheetCount: sheetNames.length,
      targetRowCount,
      editCount,
      readCount,
      formulaCounts,
      patchedCacheCount: formulaCounts.patchedFormulaCacheCount,
    }
    return {
      bytesWritten: output.bytesWritten,
      warnings: [],
      sheetNames,
      reads: readValues,
      changes: [],
      diagnostics,
    }
  } catch (error) {
    if (error instanceof StreamingNativeXlsxRecalcError) {
      throw error
    }
    const unsupportedReason = error instanceof Error ? error.message : String(error)
    throw streamingError(`streaming-native could not recalculate this workbook: ${unsupportedReason}`, phaseRssPeaks, {
      inputBytes,
      maxObservedRssBytes,
      ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
      unsupportedReason,
    })
  } finally {
    source.release?.()
  }
}

function streamingError(
  message: string,
  phaseRssPeaks: readonly XlsxFormulaRecalcPhaseRss[],
  args: {
    readonly inputBytes: number
    readonly maxObservedRssBytes: number
    readonly maxRssBytes?: number
    readonly unsupportedReason?: string
  },
): StreamingNativeXlsxRecalcError {
  return new StreamingNativeXlsxRecalcError(message, {
    engineMode: 'streaming-native',
    inputBytes: args.inputBytes,
    phaseRssPeaks,
    maxObservedRssBytes: args.maxObservedRssBytes,
    ...(args.maxRssBytes === undefined ? {} : { maxRssBytes: args.maxRssBytes }),
    sheetCount: 0,
    targetRowCount: 0,
    editCount: 0,
    readCount: 0,
    formulaCounts: {
      scannedFormulaCellCount: 0,
      targetedFormulaCellCount: 0,
      evaluatedFormulaCellCount: 0,
      patchedFormulaCacheCount: 0,
      unsupportedFormulaCellCount: args.unsupportedReason ? 1 : 0,
      nativeKernelFormulaCellCount: 0,
      nativeKernelBatchCount: 0,
    },
    patchedCacheCount: 0,
    ...(args.unsupportedReason === undefined ? {} : { unsupportedReason: args.unsupportedReason }),
  })
}

function parseQualifiedTarget(target: string): QualifiedTarget {
  const trimmed = target.trim()
  const separator = findSheetSeparator(trimmed)
  if (separator <= 0 || separator >= trimmed.length - 1) {
    throw new Error(`Expected a sheet-qualified A1 target such as Data!R57152, received: ${target}`)
  }
  const sheetName = unquoteSheetName(trimmed.slice(0, separator))
  const address = encodeCellAddress(
    decodeCellAddress(
      trimmed
        .slice(separator + 1)
        .replaceAll('$', '')
        .toUpperCase(),
    ),
  )
  const decoded = decodeCellAddress(address)
  return {
    source: target,
    sheetName,
    address,
    row: decoded.r,
    col: decoded.c,
  }
}

function findSheetSeparator(value: string): number {
  let quoted = false
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index]
    if (character === "'") {
      if (quoted && value[index + 1] === "'") {
        index += 1
        continue
      }
      quoted = !quoted
      continue
    }
    if (character === '!' && !quoted) {
      return index
    }
  }
  return -1
}

function unquoteSheetName(value: string): string {
  if (value.startsWith("'") && value.endsWith("'")) {
    return value.slice(1, -1).replaceAll("''", "'")
  }
  return value
}

function validateTargets(targets: readonly QualifiedTarget[], sheetPathsByName: ReadonlyMap<string, string>): void {
  for (const target of targets) {
    if (!sheetPathsByName.has(target.sheetName)) {
      throw new Error(`Unknown sheet in XLSX formula recalculation target: ${target.sheetName}`)
    }
  }
}

function targetRowsForTargets(targets: readonly QualifiedTarget[]): Map<string, Set<number>> {
  const output = new Map<string, Set<number>>()
  for (const target of targets) {
    const rows = output.get(target.sheetName) ?? new Set<number>()
    rows.add(target.row)
    output.set(target.sheetName, rows)
  }
  return output
}

function scanWorksheet(zip: XlsxZipEntries, sheetName: string, sheetPath: string, targetRows: Set<number>): SheetScanState {
  const state: SheetScanState = {
    sheetName,
    sheetPath,
    targetRows,
    rows: new Map(),
    formulaCells: [],
    sharedFormulaMasters: new Map(),
    tableRelationshipIds: new Set(),
    scannedFormulaCellCount: 0,
  }
  let buffer = ''
  const processBuffer = (final: boolean): void => {
    const safeEnd = final ? buffer.length : Math.max(0, buffer.lastIndexOf('<'))
    if (safeEnd === 0 && !final) {
      return
    }
    const safeXml = buffer.slice(0, safeEnd)
    for (const match of safeXml.matchAll(new RegExp(worksheetCellElementPattern.source, 'gu'))) {
      scanWorksheetCell(state, match[0])
    }
    for (const match of safeXml.matchAll(tablePartPattern)) {
      const id = readXmlAttribute(match[0], 'r:id') ?? readXmlAttribute(match[0], 'id')
      if (id) {
        state.tableRelationshipIds.add(id)
      }
    }
    buffer = buffer.slice(safeEnd)
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
    throw new Error(`Unable to stream worksheet XML for ${sheetName}`)
  }
  buffer += textDecoder.decode()
  processBuffer(true)
  return state
}

function scanWorksheetCell(state: SheetScanState, cellXml: string): void {
  const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
  const addressText = openingTag ? readXmlAttribute(openingTag, 'r') : null
  if (!openingTag || !addressText) {
    return
  }
  let address
  try {
    address = decodeCellAddress(addressText)
  } catch {
    return
  }
  const formula = readFormulaInfo(cellXml)
  if (formula) {
    state.scannedFormulaCellCount += 1
    if (formula.sharedFormulaIndex && formula.formula) {
      state.sharedFormulaMasters.set(formula.sharedFormulaIndex, {
        formula: formula.formula,
        row: address.r,
        col: address.c,
      })
    }
  }
  if (!state.targetRows.has(address.r)) {
    return
  }
  const row = state.rows.get(address.r) ?? new Map<number, PendingCellValue>()
  row.set(address.c, readCellValue(cellXml, openingTag))
  state.rows.set(address.r, row)
  if (formula) {
    state.formulaCells.push({
      sheetName: state.sheetName,
      address: encodeCellAddress(address),
      row: address.r,
      col: address.c,
      formula: formula.formula,
      sharedFormulaIndex: formula.sharedFormulaIndex,
    })
  }
}

function readFormulaInfo(cellXml: string): FormulaInfo | null {
  const formulaXml = formulaElementPattern.exec(cellXml)?.[0]
  if (!formulaXml) {
    return null
  }
  const openingEnd = formulaXml.indexOf('>')
  const openingTag = openingEnd >= 0 ? formulaXml.slice(0, openingEnd + 1) : formulaXml
  const sharedFormulaIndex = readXmlAttribute(openingTag, 'si')
  const sharedFormulaRef = readXmlAttribute(openingTag, 'ref')
  const formula =
    formulaXml.endsWith('/>') || openingEnd < 0
      ? null
      : decodeXmlText(formulaXml.slice(openingEnd + 1, formulaXml.replace(/<\/(?:[A-Za-z_][\w.-]*:)?f>\s*$/u, '').length))
  return {
    formula: formula && formula.trim().length > 0 ? formula : null,
    sharedFormulaIndex,
    sharedFormulaRef,
  }
}

function readCellValue(cellXml: string, openingTag: string): PendingCellValue {
  const type = readXmlAttribute(openingTag, 't')
  if (type === 'inlineStr') {
    return stringCellValue(readTextRuns(cellXml))
  }
  const rawValue = readElementText(cellXml, 'v')
  if (rawValue === null) {
    return emptyCellValue
  }
  if (type === 's') {
    const index = Number(rawValue)
    return Number.isSafeInteger(index) && index >= 0 ? { kind: 'shared-string', index } : emptyCellValue
  }
  if (type === 'str') {
    return stringCellValue(decodeXmlText(rawValue))
  }
  if (type === 'b') {
    return { tag: ValueTag.Boolean, value: rawValue === '1' || rawValue.toLowerCase() === 'true' }
  }
  if (type === 'e') {
    return { tag: ValueTag.Error, code: errorCodeForText(decodeXmlText(rawValue)) }
  }
  const numeric = Number(rawValue)
  return Number.isFinite(numeric) ? { tag: ValueTag.Number, value: numeric } : stringCellValue(decodeXmlText(rawValue))
}

function errorCodeForText(value: string): ErrorCode {
  switch (value.toUpperCase()) {
    case '#DIV/0!':
      return ErrorCode.Div0
    case '#REF!':
      return ErrorCode.Ref
    case '#VALUE!':
      return ErrorCode.Value
    case '#NAME?':
      return ErrorCode.Name
    case '#N/A':
      return ErrorCode.NA
    case '#NUM!':
      return ErrorCode.Num
    case '#FIELD!':
      return ErrorCode.Field
    case '#NULL!':
      return ErrorCode.Null
    default:
      return ErrorCode.Value
  }
}

function stringCellValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function collectSharedStringReferences(sheetScans: ReadonlyMap<string, SheetScanState>): ReadonlySet<number> {
  const indexes = new Set<number>()
  for (const scan of sheetScans.values()) {
    for (const row of scan.rows.values()) {
      for (const value of row.values()) {
        if (isSharedStringReference(value)) {
          indexes.add(value.index)
        }
      }
    }
  }
  return indexes
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
        if (values.size === targetIndexes.size) {
          return false
        }
      }
      index += 1
    }
    buffer = buffer.slice(safeEnd)
    return true
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

function hydrateSharedStringReferences(sheetScans: ReadonlyMap<string, SheetScanState>, sharedStrings: ReadonlyMap<number, string>): void {
  for (const scan of sheetScans.values()) {
    for (const row of scan.rows.values()) {
      for (const [col, value] of row.entries()) {
        if (!isSharedStringReference(value)) {
          continue
        }
        const sharedString = sharedStrings.get(value.index)
        if (sharedString === undefined) {
          throw new Error(`Unable to resolve shared string ${String(value.index)}`)
        }
        row.set(col, stringCellValue(sharedString))
      }
    }
  }
}

function isSharedStringReference(value: PendingCellValue): value is SharedStringReference {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}

function readNativeTablesBySheet(
  zip: XlsxZipEntries,
  sheetScans: ReadonlyMap<string, SheetScanState>,
): ReadonlyMap<string, readonly NativeTable[]> {
  const output = new Map<string, NativeTable[]>()
  for (const scan of sheetScans.values()) {
    if (scan.tableRelationshipIds.size === 0) {
      continue
    }
    const relationships = readWorksheetRelationships(zip, scan.sheetPath)
    const tables: NativeTable[] = []
    for (const relationshipId of scan.tableRelationshipIds) {
      const target = relationships.get(relationshipId)
      if (!target) {
        continue
      }
      const tableXml = getZipText(zip, target)
      const table = tableXml ? parseNativeTableXml(scan.sheetName, tableXml) : null
      if (table) {
        tables.push(table)
      }
    }
    if (tables.length > 0) {
      output.set(scan.sheetName, tables)
    }
  }
  return output
}

function readWorksheetRelationships(zip: XlsxZipEntries, worksheetPath: string): ReadonlyMap<string, string> {
  const relationshipsXml = getZipText(zip, worksheetRelationshipsPath(worksheetPath))
  const relationships = new Map<string, string>()
  if (!relationshipsXml) {
    return relationships
  }
  for (const match of relationshipsXml.matchAll(relationshipPattern)) {
    const tag = match[0]
    const id = readXmlAttribute(tag, 'Id')
    const target = readXmlAttribute(tag, 'Target')
    const type = readXmlAttribute(tag, 'Type')
    if (id && target && type && (type === tableRelationshipType || type.endsWith('/table'))) {
      relationships.set(id, resolveTargetPath(worksheetPath, target))
    }
  }
  return relationships
}

function worksheetRelationshipsPath(worksheetPath: string): string {
  const slash = worksheetPath.lastIndexOf('/')
  const directory = slash >= 0 ? worksheetPath.slice(0, slash) : ''
  const fileName = slash >= 0 ? worksheetPath.slice(slash + 1) : worksheetPath
  return `${directory}/_rels/${fileName}.rels`
}

function resolveTargetPath(fromPath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target)
  }
  const baseParts = normalizeZipPath(fromPath).split('/').slice(0, -1)
  for (const part of target.split('/')) {
    if (part.length === 0 || part === '.') {
      continue
    }
    if (part === '..') {
      baseParts.pop()
      continue
    }
    baseParts.push(part)
  }
  return normalizeZipPath(baseParts.join('/'))
}

function parseNativeTableXml(sheetName: string, tableXml: string): NativeTable | null {
  const tableTag = /<(?:[A-Za-z_][\w.-]*:)?table\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(tableXml)?.[0]
  if (!tableTag) {
    return null
  }
  const name = readXmlAttribute(tableTag, 'name') ?? readXmlAttribute(tableTag, 'displayName')
  const displayName = readXmlAttribute(tableTag, 'displayName') ?? name
  const ref = readXmlAttribute(tableTag, 'ref')
  if (!name || !displayName || !ref) {
    return null
  }
  let range
  try {
    range = decodeCellRange(ref)
  } catch {
    return null
  }
  const headerRowCount = parseNonNegativeInteger(readXmlAttribute(tableTag, 'headerRowCount'), 1)
  const totalsRowCount = parseNonNegativeInteger(readXmlAttribute(tableTag, 'totalsRowCount'), 0)
  const columns = [...tableXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?tableColumn\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/gu)].flatMap((match) => {
    const columnName = readXmlAttribute(match[0], 'name')
    return columnName ? [normalizeStructuredReferenceColumnName(columnName)] : []
  })
  return columns.length > 0
    ? {
        name,
        displayName,
        sheetName,
        range,
        headerRowCount,
        totalsRowCount,
        columns,
      }
    : null
}

function parseNonNegativeInteger(raw: string | null, fallback: number): number {
  if (raw === null) {
    return fallback
  }
  const value = Number(raw)
  return Number.isSafeInteger(value) && value >= 0 ? value : fallback
}

function applyEdits(
  edits: readonly QualifiedEdit[],
  sheetScans: ReadonlyMap<string, SheetScanState>,
  patches: XlsxSourceLiteralPatch[],
): void {
  for (const edit of edits) {
    if (typeof edit.value === 'string' && edit.value.startsWith('=')) {
      throw new Error(`streaming-native does not support formula edits: ${edit.sheetName}!${edit.address}`)
    }
    const scan = sheetScans.get(edit.sheetName)
    if (!scan) {
      throw new Error(`Missing scan state for edited sheet ${edit.sheetName}`)
    }
    const row = scan.rows.get(edit.row) ?? new Map<number, PendingCellValue>()
    row.set(edit.col, cellValueFromLiteralInput(edit.value))
    scan.rows.set(edit.row, row)
    patches.push({
      sheetName: edit.sheetName,
      address: edit.address,
      value: edit.value,
    })
  }
}

function cellValueFromLiteralInput(value: LiteralInput): CellValue {
  if (value === null) {
    return emptyCellValue
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { tag: ValueTag.Number, value } : { tag: ValueTag.Error, code: ErrorCode.Num }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  return stringCellValue(value)
}

function evaluateFormulaCells(
  sheetScans: ReadonlyMap<string, SheetScanState>,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
  patches: XlsxSourceLiteralPatch[],
): StreamingNativeFormulaCounts {
  const nativeFormulaCells = evaluateStreamingNativeWasmFormulas({
    sheetScans,
    tablesBySheet,
    resolveFormulaSource,
  })
  let evaluatedFormulaCellCount = nativeFormulaCells.evaluatedFormulaCellCount
  let patchedFormulaCacheCount = nativeFormulaCells.patches.length
  let unsupportedFormulaCellCount = 0
  patches.push(...nativeFormulaCells.patches)
  for (const scan of sheetScans.values()) {
    const formulaCells = scan.formulaCells.toSorted((left, right) => left.row - right.row || left.col - right.col)
    const formulaPatches = new Map<string, XlsxSourceLiteralPatch>()
    const maxPasses = Math.max(1, formulaCells.length + 1)
    let converged = formulaCells.length === 0
    for (let pass = 0; pass < maxPasses; pass += 1) {
      let changed = false
      for (const cell of formulaCells) {
        if (nativeFormulaCells.processedCells.has(`${cell.sheetName}!${cell.address}`)) {
          continue
        }
        const rowValues = scan.rows.get(cell.row)
        if (!rowValues) {
          continue
        }
        try {
          const formula = resolveFormulaSource(scan, cell)
          const ast = parseFormula(formula)
          const value = evaluateFormulaAst(ast, {
            sheetName: cell.sheetName,
            row: cell.row,
            col: cell.col,
            rowValues,
            sheetRows: scan.rows,
            tablesBySheet,
          })
          changed = changed || !cellValuesEqual(resolvedCellValue(rowValues.get(cell.col)), value)
          rowValues.set(cell.col, value)
          const patchValue = literalInputForFormulaCache(value)
          if (patchValue === undefined) {
            throw new UnsupportedStreamingNativeFormulaError(`unsupported formula result at ${cell.sheetName}!${cell.address}`)
          }
          formulaPatches.set(`${cell.sheetName}!${cell.address}`, {
            sheetName: cell.sheetName,
            address: cell.address,
            value: patchValue,
            preserveFormula: true,
          })
        } catch (error) {
          unsupportedFormulaCellCount += 1
          throw error
        }
      }
      if (!changed) {
        converged = true
        break
      }
    }
    if (!converged) {
      throw new UnsupportedStreamingNativeFormulaError(`row-local formulas did not converge for sheet ${scan.sheetName}`)
    }
    evaluatedFormulaCellCount += formulaPatches.size
    patchedFormulaCacheCount += formulaPatches.size
    patches.push(...formulaPatches.values())
  }
  return {
    scannedFormulaCellCount: [...sheetScans.values()].reduce((sum, scan) => sum + scan.scannedFormulaCellCount, 0),
    targetedFormulaCellCount: [...sheetScans.values()].reduce((sum, scan) => sum + scan.formulaCells.length, 0),
    evaluatedFormulaCellCount,
    patchedFormulaCacheCount,
    unsupportedFormulaCellCount,
    nativeKernelFormulaCellCount: nativeFormulaCells.evaluatedFormulaCellCount,
    nativeKernelBatchCount: nativeFormulaCells.batchCount,
  }
}

function cellValuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true
    case ValueTag.Number:
      return right.tag === ValueTag.Number && Object.is(left.value, right.value)
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code
  }
}

function resolveFormulaSource(scan: SheetScanState, cell: NativeFormulaCell): string {
  if (cell.formula) {
    return cell.formula
  }
  if (!cell.sharedFormulaIndex) {
    throw new UnsupportedStreamingNativeFormulaError(`missing formula text at ${cell.sheetName}!${cell.address}`)
  }
  const master = scan.sharedFormulaMasters.get(cell.sharedFormulaIndex)
  if (!master) {
    throw new UnsupportedStreamingNativeFormulaError(
      `missing shared formula master ${cell.sharedFormulaIndex} at ${cell.sheetName}!${cell.address}`,
    )
  }
  return translateFormulaReferences(master.formula, cell.row - master.row, cell.col - master.col)
}

function evaluateFormulaAst(node: FormulaNode, context: EvaluationContext): CellValue {
  switch (node.kind) {
    case 'NumberLiteral':
      return { tag: ValueTag.Number, value: node.value }
    case 'BooleanLiteral':
      return { tag: ValueTag.Boolean, value: node.value }
    case 'StringLiteral':
      return stringCellValue(node.value)
    case 'CellRef':
      return readCellReference(node, context)
    case 'StructuredRef':
      return readStructuredReference(node, context)
    case 'UnaryExpr': {
      const value = evaluateFormulaAst(node.argument, context)
      const number = coerceNumber(value)
      return { tag: ValueTag.Number, value: node.operator === '-' ? -number : number }
    }
    case 'BinaryExpr':
      return evaluateBinaryExpr(node.operator, evaluateFormulaAst(node.left, context), evaluateFormulaAst(node.right, context))
    case 'CallExpr':
      return evaluateCallExpr(node, context)
    case 'ErrorLiteral':
      return { tag: ValueTag.Error, code: node.code }
    case 'OmittedArgument':
    case 'ArrayConstant':
    case 'NameRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
    case 'InvokeExpr':
      throw new UnsupportedStreamingNativeFormulaError(`unsupported formula node: ${node.kind}`)
  }
}

function readCellReference(node: Extract<FormulaNode, { readonly kind: 'CellRef' }>, context: EvaluationContext): CellValue {
  if (node.sheetName && node.sheetName !== context.sheetName) {
    throw new UnsupportedStreamingNativeFormulaError(`cross-sheet direct reference is not row-local: ${node.sheetName}!${node.ref}`)
  }
  const address = decodeCellAddress(node.ref)
  const rowValues = address.r === context.row ? context.rowValues : context.sheetRows.get(address.r)
  return resolvedCellValue(rowValues?.get(address.c))
}

function readStructuredReference(node: Extract<FormulaNode, { readonly kind: 'StructuredRef' }>, context: EvaluationContext): CellValue {
  if (node.endColumnName) {
    throw new UnsupportedStreamingNativeFormulaError(
      `multi-column structured reference is not scalar: ${node.columnName}:${node.endColumnName}`,
    )
  }
  if (node.section !== undefined && node.section !== 'this-row') {
    throw new UnsupportedStreamingNativeFormulaError(`unsupported structured reference section: ${node.section}`)
  }
  const table = findCurrentRowTable(node.tableName, context)
  const columnName = normalizeStructuredReferenceColumnName(node.columnName)
  const columnIndex = table.columns.findIndex((column) => column === columnName)
  const normalizedColumnIndex =
    columnIndex >= 0
      ? columnIndex
      : table.columns.findIndex((column) => column.toLocaleLowerCase('en-US') === columnName.toLocaleLowerCase('en-US'))
  if (normalizedColumnIndex < 0) {
    throw new UnsupportedStreamingNativeFormulaError(`unknown structured reference column: ${node.columnName}`)
  }
  return resolvedCellValue(context.rowValues.get(table.range.s.c + normalizedColumnIndex))
}

function findCurrentRowTable(tableName: string, context: EvaluationContext): NativeTable {
  const tables = context.tablesBySheet.get(context.sheetName) ?? []
  const matching = tables.filter((table) => {
    const nameMatches =
      tableName.length === 0 ||
      table.name.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US') ||
      table.displayName.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US')
    return nameMatches && rowIsInTableDataBody(table, context.row)
  })
  const containingFormulaCell = matching.find((table) => context.col >= table.range.s.c && context.col <= table.range.e.c)
  const table = containingFormulaCell ?? matching[0]
  if (!table) {
    throw new UnsupportedStreamingNativeFormulaError(
      tableName.length === 0 ? 'unable to resolve current-row table' : `unable to resolve current-row table: ${tableName}`,
    )
  }
  return table
}

function rowIsInTableDataBody(table: NativeTable, row: number): boolean {
  const start = table.range.s.r + table.headerRowCount
  const end = table.range.e.r - table.totalsRowCount
  return row >= start && row <= end
}

function evaluateBinaryExpr(operator: string, left: CellValue, right: CellValue): CellValue {
  switch (operator) {
    case '+':
      return { tag: ValueTag.Number, value: coerceNumber(left) + coerceNumber(right) }
    case '-':
      return { tag: ValueTag.Number, value: coerceNumber(left) - coerceNumber(right) }
    case '*':
      return { tag: ValueTag.Number, value: coerceNumber(left) * coerceNumber(right) }
    case '/': {
      const divisor = coerceNumber(right)
      return divisor === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : { tag: ValueTag.Number, value: coerceNumber(left) / divisor }
    }
    case '&':
      return stringCellValue(cellValueText(left) + cellValueText(right))
    case '=':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) === 0 }
    case '<>':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) !== 0 }
    case '>':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) > 0 }
    case '>=':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) >= 0 }
    case '<':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) < 0 }
    case '<=':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) <= 0 }
    case '^':
    case ':':
      throw new UnsupportedStreamingNativeFormulaError(`unsupported binary operator: ${operator}`)
    default:
      throw new UnsupportedStreamingNativeFormulaError(`unknown binary operator: ${operator}`)
  }
}

function evaluateCallExpr(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  const callee = node.callee.toUpperCase().replace(/^_XLFN\./u, '')
  if (callee === 'IF') {
    if (node.args.length < 2 || node.args.length > 3) {
      throw new UnsupportedStreamingNativeFormulaError('IF requires 2 or 3 arguments')
    }
    return coerceBoolean(evaluateFormulaAst(node.args[0]!, context))
      ? evaluateFormulaAst(node.args[1]!, context)
      : node.args[2]
        ? evaluateFormulaAst(node.args[2], context)
        : { tag: ValueTag.Boolean, value: false }
  }
  if (callee === 'IFS') {
    if (node.args.length < 2 || node.args.length % 2 !== 0) {
      throw new UnsupportedStreamingNativeFormulaError('IFS requires condition/value pairs')
    }
    for (let index = 0; index < node.args.length; index += 2) {
      if (coerceBoolean(evaluateFormulaAst(node.args[index]!, context))) {
        return evaluateFormulaAst(node.args[index + 1]!, context)
      }
    }
    return { tag: ValueTag.Error, code: ErrorCode.NA }
  }
  throw new UnsupportedStreamingNativeFormulaError(`unsupported function: ${node.callee}`)
}

function coerceNumber(value: CellValue): number {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String: {
      const numeric = Number(value.value)
      if (Number.isFinite(numeric)) {
        return numeric
      }
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce string to number: ${value.value}`)
    }
    case ValueTag.Error:
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce error to number: ${String(value.code)}`)
  }
}

function coerceBoolean(value: CellValue): boolean {
  switch (value.tag) {
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Number:
      return value.value !== 0
    case ValueTag.Empty:
      return false
    case ValueTag.String:
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce string to boolean: ${value.value}`)
    case ValueTag.Error:
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce error to boolean: ${String(value.code)}`)
  }
}

function compareCellValues(left: CellValue, right: CellValue): number {
  if (isNumericComparable(left) && isNumericComparable(right)) {
    return coerceNumber(left) - coerceNumber(right)
  }
  const leftText = cellValueText(left).toLocaleLowerCase('en-US')
  const rightText = cellValueText(right).toLocaleLowerCase('en-US')
  return leftText < rightText ? -1 : leftText > rightText ? 1 : 0
}

function isNumericComparable(value: CellValue): boolean {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean || value.tag === ValueTag.Empty
}

function cellValueText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      throw new UnsupportedStreamingNativeFormulaError(`cannot use error as text: ${String(value.code)}`)
  }
}

function readTargets(
  reads: readonly QualifiedTarget[],
  sheetScans: ReadonlyMap<string, SheetScanState>,
): Readonly<Record<string, CellValue>> {
  const output: Record<string, CellValue> = {}
  for (const read of reads) {
    const row = sheetScans.get(read.sheetName)?.rows.get(read.row)
    output[read.source] = resolvedCellValue(row?.get(read.col))
  }
  return output
}

function resolvedCellValue(value: PendingCellValue | undefined): CellValue {
  if (value === undefined || isSharedStringReference(value)) {
    return emptyCellValue
  }
  return value
}

function literalInputForFormulaCache(value: CellValue): LiteralInput | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return Number.isFinite(value.value) ? value.value : undefined
    case ValueTag.Boolean:
      return value.value
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return undefined
  }
}

function readTextRuns(xml: string): string {
  const runs = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b(?:[^>"']|"[^"]*"|'[^']*')*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)].map(
    (match) => decodeXmlText(match[1] ?? ''),
  )
  return runs.length > 0 ? runs.join('') : ''
}

function readElementText(xml: string, name: string): string | null {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  return (
    new RegExp(
      `<(?:[A-Za-z_][\\w.-]*:)?${escapedName}\\b(?:[^>"']|"[^"]*"|'[^']*')*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${escapedName}>`,
      'u',
    ).exec(xml)?.[1] ?? null
  )
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function normalizeStructuredReferenceColumnName(value: string): string {
  return decodeExcelEscapedText(value).replace(/\r\n?/gu, '\n').trim()
}

function decodeExcelEscapedText(value: string): string {
  const escapedUnderscore = '\uE000'
  return value
    .replace(/_x005F_/giu, escapedUnderscore)
    .replace(/_x([0-9a-fA-F]{4})_/gu, (_match, code: string) => {
      const codePoint = Number.parseInt(code, 16)
      return Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : ''
    })
    .replaceAll(escapedUnderscore, '_')
}
