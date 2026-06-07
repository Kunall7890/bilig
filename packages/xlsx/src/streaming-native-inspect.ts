import { statSync } from 'node:fs'

import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'

import {
  normalizeXlsxFormulaCacheInspectionLimit,
  readXlsxFormulaCacheCellsFromFile,
  XlsxFormulaCacheReadError,
  type XlsxFormulaCacheCell,
  type XlsxFormulaCacheInspectionLimit,
  type XlsxFormulaCacheLiteral,
} from './formula-cache-reader.js'
import {
  recalculateXlsxFileToFileStreamingNative,
  StreamingNativeXlsxRecalcError,
  type XlsxFormulaRecalcNativeDiagnostics,
  type XlsxFormulaRecalcPhaseRss,
} from './streaming-native-recalc.js'

export type StreamingNativeXlsxCacheInspectionLimit = number | 'all'
export type StreamingNativeXlsxCacheStatus = 'fresh' | 'stale' | 'missing-cache' | 'unsupported-recalculation'
export type StreamingNativeXlsxCacheLiteral = XlsxFormulaCacheLiteral

export const defaultStreamingNativeXlsxCacheInspectionLimit: StreamingNativeXlsxCacheInspectionLimit = 2000

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

  let sheetNames: readonly string[] = []
  let formulaCacheScan: { readonly formulaCellCount: number; readonly cells: readonly XlsxFormulaCacheCell[] } = {
    formulaCellCount: 0,
    cells: [],
  }
  const inspectLimit = options.inspectLimit ?? defaultStreamingNativeXlsxCacheInspectionLimit
  try {
    const scan = readXlsxFormulaCacheCellsFromFile(inputPath, {
      inspectLimit,
      onPhase: (phase) => recordPhase(`inspect-${phase}`),
    })
    sheetNames = scan.sheetNames
    formulaCacheScan = scan
  } catch (error) {
    if (error instanceof XlsxFormulaCacheReadError) {
      throw streamingInspectionError(error.message, phaseRssPeaks, {
        inputBytes: error.inputBytes,
        maxObservedRssBytes,
        ...(options.maxRssBytes === undefined ? {} : { maxRssBytes: options.maxRssBytes }),
        unsupportedReason: error.reason,
      })
    }
    throw error
  }

  const inspectionLimit = normalizeStreamingNativeInspectionLimit(inspectLimit)
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
  return normalizeXlsxFormulaCacheInspectionLimit(limit as XlsxFormulaCacheInspectionLimit) as StreamingNativeXlsxCacheInspectionLimit
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
    fallbackUsed: false,
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
    default:
      return '#VALUE!'
  }
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
