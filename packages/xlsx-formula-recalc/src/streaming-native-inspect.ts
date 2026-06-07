import { statSync } from 'node:fs'

import { translateFormulaReferences } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import {
  createFileXlsxSourceReader,
  decodeCellAddress,
  encodeCellAddress,
  forEachInflatedXlsxZipEntryChunk,
  readXmlAttribute,
  readXlsxZipEntriesLazyFromByteSource,
  workbookSheetPathEntriesForSource,
  worksheetCellElementPattern,
  worksheetCellOpeningTagPattern,
  type XlsxZipEntries,
} from '@bilig/xlsx'

import {
  recalculateXlsxFileToFileStreamingNative,
  StreamingNativeXlsxRecalcError,
  type XlsxFormulaRecalcNativeDiagnostics,
  type XlsxFormulaRecalcPhaseRss,
} from './streaming-native-recalc.js'

export type StreamingNativeXlsxCacheInspectionLimit = number | 'all'
export type StreamingNativeXlsxCacheStatus = 'fresh' | 'stale' | 'missing-cache' | 'unsupported-recalculation'
export type StreamingNativeXlsxCacheLiteral = string | number | boolean | null

export interface StreamingNativeXlsxCacheStatusSummary {
  readonly inspected: number
  readonly stale: number
  readonly fresh: number
  readonly missingCache: number
  readonly unsupportedRecalculation: number
}

export interface StreamingNativeXlsxCacheFormulaInspection {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: StreamingNativeXlsxCacheLiteral
  readonly literalRecalculatedValue?: StreamingNativeXlsxCacheLiteral
  readonly cacheStatus: StreamingNativeXlsxCacheStatus
  readonly staleCachedValue: boolean | null
}

export interface StreamingNativeXlsxCacheInspectionResult {
  readonly schemaVersion: 'xlsx-cache-doctor.v1'
  readonly sheetNames: readonly string[]
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly uninspectedFormulaCellCount: number
  readonly inspectionLimit: StreamingNativeXlsxCacheInspectionLimit
  readonly staleCachedFormulaCount: number
  readonly cacheStatusSummary: StreamingNativeXlsxCacheStatusSummary
  readonly suggestedReads: readonly string[]
  readonly formulas: readonly StreamingNativeXlsxCacheFormulaInspection[]
  readonly warnings: readonly string[]
  readonly diagnostics: XlsxFormulaRecalcNativeDiagnostics
  readonly inspectionCompleted: true
  readonly recalculationCompleted: boolean
  readonly excelParity: 'not_proven'
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

type PendingCellValue = CellValue | SharedStringReference

interface NativeFormulaCacheCell {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: StreamingNativeXlsxCacheLiteral
}

interface PendingNativeFormulaCacheCell {
  readonly target: string
  readonly formula: string
  readonly cachedValue?: PendingCellValue
}

interface NativeFormulaCacheCells {
  readonly formulaCellCount: number
  readonly cells: readonly NativeFormulaCacheCell[]
}

const emptyCellValue: CellValue = Object.freeze({ tag: ValueTag.Empty })
const textDecoder = new TextDecoder()
const formulaElementPattern = /<((?:[A-Za-z_][\w.-]*:)?f)\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/\1>)/u

export async function inspectXlsxCacheFileStreamingNative(
  inputPath: string,
  options: {
    readonly inspectLimit?: StreamingNativeXlsxCacheInspectionLimit
    readonly maxRssBytes?: number
  } = {},
): Promise<StreamingNativeXlsxCacheInspectionResult> {
  const inputBytes = statSync(inputPath).size
  const phaseRssPeaks: XlsxFormulaRecalcPhaseRss[] = []
  let maxObservedRssBytes = 0
  const recordPhase = (phase: string): void => {
    const rssBytes = process.memoryUsage().rss
    maxObservedRssBytes = Math.max(maxObservedRssBytes, rssBytes)
    phaseRssPeaks.push({ phase, rssBytes })
    if (options.maxRssBytes !== undefined && rssBytes > options.maxRssBytes) {
      throw streamingInspectionError(`streaming-native inspection exceeded maxRssBytes during ${phase}`, phaseRssPeaks, {
        inputBytes,
        maxObservedRssBytes,
        maxRssBytes: options.maxRssBytes,
      })
    }
  }

  const source = createFileXlsxSourceReader(inputPath)
  let sheetNames: readonly string[] = []
  let formulaCacheScan: NativeFormulaCacheCells = { formulaCellCount: 0, cells: [] }
  try {
    recordPhase('inspect-open-source')
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      throw streamingInspectionError('streaming-native inspection requires a ZIP central directory it can read lazily', phaseRssPeaks, {
        inputBytes,
        maxObservedRssBytes,
        ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
        unsupportedReason: 'invalid-or-zip64-xlsx',
      })
    }
    const sheetEntries = workbookSheetPathEntriesForSource(zip)
    sheetNames = sheetEntries.map((sheet) => sheet.name)
    recordPhase('inspect-workbook-metadata')
    formulaCacheScan = collectNativeFormulaCacheCells(
      zip,
      sheetEntries,
      normalizeStreamingNativeInspectionLimit(options.inspectLimit ?? 'all'),
    )
    recordPhase('inspect-formula-cache-scan')
  } finally {
    source.release?.()
  }

  const inspectionLimit = normalizeStreamingNativeInspectionLimit(options.inspectLimit ?? 'all')
  const inspectedFormulaCells = formulaCacheScan.cells
  const suggestedReads = inspectedFormulaCells.map((cell) => cell.target)
  if (suggestedReads.length === 0) {
    const diagnostics = inspectionDiagnostics({
      inputBytes,
      phaseRssPeaks,
      maxObservedRssBytes,
      ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
      sheetCount: sheetNames.length,
      inspectedFormulaCellCount: 0,
      formulaCellCount: formulaCacheScan.formulaCellCount,
    })
    return {
      schemaVersion: 'xlsx-cache-doctor.v1',
      sheetNames,
      formulaCellCount: formulaCacheScan.formulaCellCount,
      inspectedFormulaCellCount: 0,
      uninspectedFormulaCellCount: formulaCacheScan.formulaCellCount,
      inspectionLimit,
      staleCachedFormulaCount: 0,
      cacheStatusSummary: { inspected: 0, stale: 0, fresh: 0, missingCache: 0, unsupportedRecalculation: 0 },
      suggestedReads,
      formulas: [],
      warnings: [],
      diagnostics,
      inspectionCompleted: true,
      recalculationCompleted: true,
      excelParity: 'not_proven',
    }
  }

  let recalculationCompleted = true
  let reads: Readonly<Record<string, CellValue>> = {}
  let recalcDiagnostics: XlsxFormulaRecalcNativeDiagnostics | undefined
  try {
    const recalculated = await recalculateXlsxFileToFileStreamingNative(inputPath, {
      outputPath: `${inputPath}.native-inspect-dry-run`,
      reads: suggestedReads,
      ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
      dryRun: true,
    })
    reads = recalculated.reads
    recalcDiagnostics = recalculated.diagnostics
  } catch (error) {
    recalculationCompleted = false
    if (error instanceof StreamingNativeXlsxRecalcError) {
      recalcDiagnostics = error.diagnostics
    } else {
      throw error
    }
  }

  const diagnostics = mergeInspectionRecalcDiagnostics({
    inputBytes,
    phaseRssPeaks,
    maxObservedRssBytes,
    ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
    sheetCount: sheetNames.length,
    inspectedFormulaCellCount: inspectedFormulaCells.length,
    formulaCellCount: formulaCacheScan.formulaCellCount,
    ...(recalcDiagnostics === undefined ? {} : { recalcDiagnostics }),
  })
  const formulas = inspectedFormulaCells.map((cell) => {
    const literalRecalculatedValue = recalculationCompleted ? literalValueForCacheInspection(reads[cell.target]) : undefined
    const cacheStatus = recalculationCompleted
      ? streamingNativeCacheStatus(cell.cachedValue, literalRecalculatedValue)
      : 'unsupported-recalculation'
    return {
      target: cell.target,
      formula: cell.formula,
      ...(cell.cachedValue !== undefined ? { cachedValue: cell.cachedValue } : {}),
      ...(literalRecalculatedValue !== undefined ? { literalRecalculatedValue } : {}),
      cacheStatus,
      staleCachedValue: streamingNativeStaleCachedValue(cacheStatus),
    }
  })

  return {
    schemaVersion: 'xlsx-cache-doctor.v1',
    sheetNames,
    formulaCellCount: formulaCacheScan.formulaCellCount,
    inspectedFormulaCellCount: inspectedFormulaCells.length,
    uninspectedFormulaCellCount: formulaCacheScan.formulaCellCount - inspectedFormulaCells.length,
    inspectionLimit,
    staleCachedFormulaCount: formulas.filter((formula) => formula.staleCachedValue === true).length,
    cacheStatusSummary: streamingNativeCacheStatusSummary(formulas),
    suggestedReads,
    formulas,
    warnings: [],
    diagnostics,
    inspectionCompleted: true,
    recalculationCompleted,
    excelParity: 'not_proven',
  }
}

function collectNativeFormulaCacheCells(
  zip: XlsxZipEntries,
  sheetEntries: readonly { readonly name: string; readonly path: string }[],
  inspectionLimit: StreamingNativeXlsxCacheInspectionLimit,
): NativeFormulaCacheCells {
  const pendingCells: PendingNativeFormulaCacheCell[] = []
  const sharedStringIndexes = new Set<number>()
  let formulaCellCount = 0
  let collectedCellCount = 0
  const shouldCollect = (): boolean => inspectionLimit === 'all' || collectedCellCount < inspectionLimit
  const markCollected = (): void => {
    collectedCellCount += 1
  }
  for (const sheet of sheetEntries) {
    const scan = collectNativeFormulaCacheCellsForSheet(zip, sheet.name, sheet.path, shouldCollect, markCollected, sharedStringIndexes)
    formulaCellCount += scan.formulaCellCount
    pendingCells.push(...scan.cells)
  }
  const sharedStrings = readTargetSharedStrings(zip, sharedStringIndexes)
  return {
    formulaCellCount,
    cells: pendingCells.map((cell) => {
      const output: {
        target: string
        formula: string
        cachedValue?: StreamingNativeXlsxCacheLiteral
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

function collectNativeFormulaCacheCellsForSheet(
  zip: XlsxZipEntries,
  sheetName: string,
  sheetPath: string,
  shouldCollect: () => boolean,
  markCollected: () => void,
  sharedStringIndexes: Set<number>,
): { readonly formulaCellCount: number; readonly cells: readonly PendingNativeFormulaCacheCell[] } {
  const sharedFormulaMasters = new Map<string, SharedFormulaMaster>()
  const cells: PendingNativeFormulaCacheCell[] = []
  let formulaCellCount = 0
  let buffer = ''
  const processBuffer = (final: boolean): void => {
    const safeEnd = final ? buffer.length : lastWorksheetCellStartIndex(buffer)
    if (safeEnd === 0 && !final) {
      return
    }
    const safeXml = buffer.slice(0, safeEnd)
    for (const match of safeXml.matchAll(new RegExp(worksheetCellElementPattern.source, 'gu'))) {
      const cellXml = match[0]
      const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
      const addressText = openingTag ? readXmlAttribute(openingTag, 'r') : null
      if (!openingTag || !addressText) {
        continue
      }
      let decodedAddress
      try {
        decodedAddress = decodeCellAddress(addressText)
      } catch {
        continue
      }
      const formula = readFormulaInfo(cellXml)
      if (!formula) {
        continue
      }
      formulaCellCount += 1
      if (formula.sharedFormulaIndex && formula.formula) {
        sharedFormulaMasters.set(formula.sharedFormulaIndex, {
          formula: formula.formula,
          row: decodedAddress.r,
          col: decodedAddress.c,
        })
      }
      if (!shouldCollect()) {
        continue
      }
      const formulaText = formulaSourceForCacheInspection(formula, sharedFormulaMasters, decodedAddress.r, decodedAddress.c)
      const cachedValue = cachedValueForFormulaCacheInspection(cellXml, openingTag, sharedStringIndexes)
      cells.push({
        target: formatNativeQualifiedTarget(sheetName, encodeCellAddress(decodedAddress)),
        formula: formulaText.startsWith('=') ? formulaText : `=${formulaText}`,
        ...(cachedValue === undefined ? {} : { cachedValue }),
      })
      markCollected()
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
  return { formulaCellCount, cells }
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
  if (readElementText(cellXml, 'v') === null && readXmlAttribute(openingTag, 't') !== 'inlineStr') {
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

function literalValueForPendingCacheInspection(
  value: PendingCellValue,
  sharedStrings: ReadonlyMap<number, string>,
): StreamingNativeXlsxCacheLiteral {
  if (isSharedStringReference(value)) {
    return sharedStrings.get(value.index) ?? ''
  }
  return literalValueForCacheInspection(value) ?? null
}

function literalValueForCacheInspection(value: CellValue | undefined): StreamingNativeXlsxCacheLiteral | undefined {
  if (value === undefined) {
    return undefined
  }
  if (value.tag === ValueTag.Error) {
    return errorTextForCode(value.code)
  }
  return literalInputForFormulaCache(value)
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
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}

function lastWorksheetCellStartIndex(xml: string): number {
  let index = -1
  for (const match of xml.matchAll(/<(?:(?:[A-Za-z_][\w.-]*):)?c\b/gu)) {
    index = match.index ?? index
  }
  return Math.max(0, index)
}

function streamingNativeCacheStatus(
  cachedValue: StreamingNativeXlsxCacheLiteral | undefined,
  literalRecalculatedValue: StreamingNativeXlsxCacheLiteral | undefined,
): StreamingNativeXlsxCacheStatus {
  if (cachedValue === undefined) {
    return 'missing-cache'
  }
  if (literalRecalculatedValue === undefined) {
    return 'unsupported-recalculation'
  }
  return JSON.stringify(cachedValue) === JSON.stringify(literalRecalculatedValue) ? 'fresh' : 'stale'
}

function streamingNativeStaleCachedValue(cacheStatus: StreamingNativeXlsxCacheStatus): boolean | null {
  switch (cacheStatus) {
    case 'stale':
      return true
    case 'fresh':
      return false
    case 'missing-cache':
    case 'unsupported-recalculation':
      return null
  }
}

function streamingNativeCacheStatusSummary(
  formulas: readonly StreamingNativeXlsxCacheFormulaInspection[],
): StreamingNativeXlsxCacheStatusSummary {
  return {
    inspected: formulas.length,
    stale: formulas.filter((formula) => formula.cacheStatus === 'stale').length,
    fresh: formulas.filter((formula) => formula.cacheStatus === 'fresh').length,
    missingCache: formulas.filter((formula) => formula.cacheStatus === 'missing-cache').length,
    unsupportedRecalculation: formulas.filter((formula) => formula.cacheStatus === 'unsupported-recalculation').length,
  }
}

function normalizeStreamingNativeInspectionLimit(limit: StreamingNativeXlsxCacheInspectionLimit): StreamingNativeXlsxCacheInspectionLimit {
  if (limit === 'all') {
    return limit
  }
  if (Number.isInteger(limit) && limit > 0) {
    return limit
  }
  throw new Error(`Expected inspectLimit to be "all" or a positive integer, received: ${String(limit)}`)
}

function inspectionDiagnostics(args: {
  readonly inputBytes: number
  readonly phaseRssPeaks: readonly XlsxFormulaRecalcPhaseRss[]
  readonly maxObservedRssBytes: number
  readonly maxRssBytes?: number
  readonly sheetCount: number
  readonly inspectedFormulaCellCount: number
  readonly formulaCellCount: number
}): XlsxFormulaRecalcNativeDiagnostics {
  return {
    engineMode: 'streaming-native',
    inputBytes: args.inputBytes,
    phaseRssPeaks: args.phaseRssPeaks,
    maxObservedRssBytes: args.maxObservedRssBytes,
    ...(args.maxRssBytes === undefined ? {} : { maxRssBytes: args.maxRssBytes }),
    sheetCount: args.sheetCount,
    targetRowCount: args.inspectedFormulaCellCount,
    editCount: 0,
    readCount: args.inspectedFormulaCellCount,
    formulaCounts: {
      scannedFormulaCellCount: args.formulaCellCount,
      targetedFormulaCellCount: args.inspectedFormulaCellCount,
      evaluatedFormulaCellCount: 0,
      patchedFormulaCacheCount: 0,
      unsupportedFormulaCellCount: 0,
      nativeKernelFormulaCellCount: 0,
      nativeKernelBatchCount: 0,
    },
    patchedCacheCount: 0,
  }
}

function mergeInspectionRecalcDiagnostics(args: {
  readonly inputBytes: number
  readonly phaseRssPeaks: readonly XlsxFormulaRecalcPhaseRss[]
  readonly maxObservedRssBytes: number
  readonly maxRssBytes?: number
  readonly sheetCount: number
  readonly inspectedFormulaCellCount: number
  readonly formulaCellCount: number
  readonly recalcDiagnostics?: XlsxFormulaRecalcNativeDiagnostics
}): XlsxFormulaRecalcNativeDiagnostics {
  const base = inspectionDiagnostics(args)
  if (!args.recalcDiagnostics) {
    return base
  }
  const phaseRssPeaks = [
    ...args.phaseRssPeaks,
    ...args.recalcDiagnostics.phaseRssPeaks.map((phase) => ({
      phase: `inspect-recalc:${phase.phase}`,
      rssBytes: phase.rssBytes,
    })),
  ]
  return {
    ...args.recalcDiagnostics,
    phaseRssPeaks,
    maxObservedRssBytes: Math.max(args.maxObservedRssBytes, args.recalcDiagnostics.maxObservedRssBytes),
    sheetCount: Math.max(args.sheetCount, args.recalcDiagnostics.sheetCount),
    formulaCounts: {
      ...args.recalcDiagnostics.formulaCounts,
      scannedFormulaCellCount: Math.max(args.formulaCellCount, args.recalcDiagnostics.formulaCounts.scannedFormulaCellCount),
      targetedFormulaCellCount: args.inspectedFormulaCellCount,
    },
  }
}

function formatNativeQualifiedTarget(sheetName: string, address: string): string {
  return `${quoteNativeSheetName(sheetName)}!${address}`
}

function quoteNativeSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/u.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function errorTextForCode(code: ErrorCode): string {
  switch (code) {
    case ErrorCode.Div0:
      return '#DIV/0!'
    case ErrorCode.Ref:
      return '#REF!'
    case ErrorCode.Value:
      return '#VALUE!'
    case ErrorCode.Name:
      return '#NAME?'
    case ErrorCode.NA:
      return '#N/A'
    case ErrorCode.Num:
      return '#NUM!'
    case ErrorCode.Field:
      return '#FIELD!'
    case ErrorCode.Null:
      return '#NULL!'
    case ErrorCode.None:
    case ErrorCode.Cycle:
    case ErrorCode.Spill:
    case ErrorCode.Blocked:
      return '#VALUE!'
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

function streamingInspectionError(
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
