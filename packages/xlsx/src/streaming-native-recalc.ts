import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'

import { decodeCellAddress, decodeCellRange, encodeCellAddress, type XlsxCellRange } from './address.js'
import type { XlsxExternalWorkbookHydrationDiagnostics, XlsxExternalWorkbookInput } from './external-workbook-types.js'
import { exportXlsxSourceLiteralPatchesToFileAsync, type XlsxSourceLiteralPatch } from './source-preserving-literal-patches.js'
import { normalizeExternalWorkbookReferences, readStreamingNativeExternalCachedRowsByAlias } from './streaming-native-external-cache.js'
import {
  StreamingNativeSheetCellArena,
  type StreamingNativeMutablePendingCellRow,
  type StreamingNativePendingCellRows,
  type StreamingNativePendingCellValue,
  type StreamingNativeSharedStringReference,
} from './streaming-native-cell-arena.js'
import { expandStreamingNativeFormulaDependencyRows } from './streaming-native-row-chain-wasm.js'
import {
  closeStreamingNativeWorkbookCore,
  openStreamingNativeWorkbookCore,
  StreamingNativeWorkbookOpenError,
  type StreamingNativeWorkbookCore,
} from './streaming-native-workbook-core.js'
import { decodeXmlText, readXmlAttribute, worksheetCellElementPattern, worksheetCellOpeningTagPattern } from './xml.js'
import { evaluateFormulaCells, readTargets, resolveFormulaSource } from './streaming-native-recalc-evaluator.js'
import { normalizeStructuredReferenceColumnName } from './streaming-native-text.js'
import { forEachInflatedXlsxZipEntryChunk, getZipText, normalizeZipPath, type XlsxZipEntries } from './zip-reader.js'

export type StreamingNativeXlsxFormulaRecalcEngineMode = 'streaming-native'

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
  readonly engineMode: StreamingNativeXlsxFormulaRecalcEngineMode
  readonly fallbackUsed: false
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
  readonly externalWorkbookHydration?: XlsxExternalWorkbookHydrationDiagnostics
}

export interface XlsxFormulaRecalcDiagnostics extends XlsxFormulaRecalcNativeDiagnostics {
  readonly externalWorkbookHydration?: XlsxExternalWorkbookHydrationDiagnostics
}

export interface StreamingNativeXlsxFormulaRecalcOptions {
  readonly outputPath: string
  readonly edits?: readonly { readonly target: string; readonly value: LiteralInput }[]
  readonly reads?: readonly string[]
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly maxRssBytes?: number
  readonly dryRun?: boolean
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
  readonly rows: StreamingNativeSheetCellArena
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

export type PendingCellValue = StreamingNativePendingCellValue
export type PendingCellRow = StreamingNativeMutablePendingCellRow
export type PendingCellRows = StreamingNativePendingCellRows

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
  if (!options.dryRun && options.outputPath.length === 0) {
    throw new Error('streaming-native outputPath is required unless dryRun is enabled')
  }
  let inputBytes = 0
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

  let core: StreamingNativeWorkbookCore | null = null
  try {
    core = openStreamingNativeWorkbookCore(inputPath, {
      onPhase: (phase, info) => {
        inputBytes = info.inputBytes
        recordPhase(phase)
      },
    })
    inputBytes = core.inputBytes
    const { source, zip, sheetEntries, sheetNames } = core
    const sheetPathsByName = new Map(sheetEntries.map((sheet) => [sheet.name, sheet.path]))
    const edits = (options.edits ?? []).map((edit) => Object.assign(parseQualifiedTarget(edit.target), { value: edit.value }))
    const reads = (options.reads ?? []).map(parseQualifiedTarget)
    validateTargets([...edits, ...reads], sheetPathsByName)
    const targetRowsBySheet = targetRowsForTargets([...edits, ...reads])

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

    const resolveParserFormulaSource = (scan: SheetScanState, cell: NativeFormulaCell): string =>
      normalizeExternalWorkbookReferences(resolveFormulaSource(scan, cell))
    const expandedDependencySheets = expandStreamingNativeFormulaDependencyRows({
      sheetScans,
      resolveFormulaSource: resolveParserFormulaSource,
      targetRowsForSheet: (sheetName) => {
        if (!sheetPathsByName.has(sheetName)) {
          return undefined
        }
        const targetRows = targetRowsBySheet.get(sheetName) ?? new Set<number>()
        targetRowsBySheet.set(sheetName, targetRows)
        return targetRows
      },
    })
    for (const sheetName of expandedDependencySheets) {
      const sheetPath = sheetPathsByName.get(sheetName)
      const targetRows = targetRowsBySheet.get(sheetName)
      if (!sheetPath || !targetRows || targetRows.size === 0) {
        continue
      }
      const nextScan = scanWorksheet(zip, sheetName, sheetPath, targetRows)
      nextScan.formulaCells.splice(0, nextScan.formulaCells.length, ...nextScan.formulaCells.filter((cell) => targetRows.has(cell.row)))
      sheetScans.set(sheetName, nextScan)
    }
    recordPhase('dependency-row-scan')

    const sharedStrings = readTargetSharedStrings(zip, collectSharedStringReferences(sheetScans))
    hydrateSharedStringReferences(sheetScans, sharedStrings)
    recordPhase('shared-strings')

    const expandedHydratedDependencySheets = expandStreamingNativeFormulaDependencyRows({
      sheetScans,
      resolveFormulaSource: resolveParserFormulaSource,
      targetRowsForSheet: (sheetName) => {
        if (!sheetPathsByName.has(sheetName)) {
          return undefined
        }
        const targetRows = targetRowsBySheet.get(sheetName) ?? new Set<number>()
        targetRowsBySheet.set(sheetName, targetRows)
        return targetRows
      },
    })
    for (const sheetName of expandedHydratedDependencySheets) {
      const sheetPath = sheetPathsByName.get(sheetName)
      const targetRows = targetRowsBySheet.get(sheetName)
      if (!sheetPath || !targetRows || targetRows.size === 0) {
        continue
      }
      const nextScan = scanWorksheet(zip, sheetName, sheetPath, targetRows)
      nextScan.formulaCells.splice(0, nextScan.formulaCells.length, ...nextScan.formulaCells.filter((cell) => targetRows.has(cell.row)))
      sheetScans.set(sheetName, nextScan)
    }
    if (expandedHydratedDependencySheets.size > 0) {
      recordPhase('hydrated-dependency-row-scan')
      const hydratedSharedStrings = readTargetSharedStrings(zip, collectSharedStringReferences(sheetScans))
      hydrateSharedStringReferences(sheetScans, hydratedSharedStrings)
      recordPhase('hydrated-dependency-shared-strings')
    }

    const tablesBySheet = readNativeTablesBySheet(zip, sheetScans)
    recordPhase('table-metadata')
    const externalCache = readStreamingNativeExternalCachedRowsByAlias(zip, sheetScans, resolveFormulaSource, options.externalWorkbooks)
    const externalCachedRowsByAlias = externalCache.rowsByAlias
    if (externalCachedRowsByAlias.size > 0) {
      recordPhase('external-link-cache')
    }

    const patches: XlsxSourceLiteralPatch[] = []
    applyEdits(edits, sheetScans, patches)
    const formulaCounts = evaluateFormulaCells(sheetScans, tablesBySheet, externalCachedRowsByAlias, patches)
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

    const bytesWritten = options.dryRun
      ? 0
      : (
          await exportXlsxSourceLiteralPatchesToFileAsync({
            source,
            patches,
            textPatches: externalCache.textPatches,
            sheetNames: patchedSheetNames,
            outputPath: options.outputPath,
          })
        ).bytesWritten
    patches.length = 0
    recordPhase(options.dryRun ? 'dry-run-output' : 'write-output')

    const diagnostics: XlsxFormulaRecalcNativeDiagnostics = {
      engineMode: 'streaming-native',
      fallbackUsed: false,
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
      ...(externalCache.diagnostics === undefined ? {} : { externalWorkbookHydration: externalCache.diagnostics }),
    }
    return {
      bytesWritten,
      warnings: externalCache.warnings,
      sheetNames,
      reads: readValues,
      changes: [],
      diagnostics,
    }
  } catch (error) {
    if (error instanceof StreamingNativeXlsxRecalcError) {
      throw error
    }
    const unsupportedReason =
      error instanceof StreamingNativeWorkbookOpenError ? error.reason : error instanceof Error ? error.message : String(error)
    const errorInputBytes = error instanceof StreamingNativeWorkbookOpenError ? error.inputBytes : inputBytes
    throw streamingError(`streaming-native could not recalculate this workbook: ${unsupportedReason}`, phaseRssPeaks, {
      inputBytes: errorInputBytes,
      maxObservedRssBytes,
      ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
      unsupportedReason,
    })
  } finally {
    if (core) {
      closeStreamingNativeWorkbookCore(core)
    }
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
    fallbackUsed: false,
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
    rows: new StreamingNativeSheetCellArena(),
    formulaCells: [],
    sharedFormulaMasters: new Map(),
    tableRelationshipIds: new Set(),
    scannedFormulaCellCount: 0,
  }
  let buffer = ''
  const processBuffer = (final: boolean): void => {
    const safeEnd = final ? buffer.length : lastWorksheetCellStartIndex(buffer)
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

function lastWorksheetCellStartIndex(xml: string): number {
  let index = -1
  for (const match of xml.matchAll(/<(?:(?:[A-Za-z_][\w.-]*):)?c\b/gu)) {
    index = match.index ?? index
  }
  return Math.max(0, index)
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
  const row = state.rows.getOrCreate(address.r)
  row.set(address.c, readCellValue(cellXml, openingTag))
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

function isSharedStringReference(value: PendingCellValue): value is StreamingNativeSharedStringReference {
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
    const row = scan.rows.getOrCreate(edit.row)
    row.set(edit.col, cellValueFromLiteralInput(edit.value))
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
