import { createRequire } from 'node:module'

import type { ImportedWorkbook } from './workbook-import-result.js'
import { XLSX_CONTENT_TYPE } from './workbook-import-content-types.js'
import {
  assertXlsxInspectionWithinMaterializationLimits,
  denseSheetJsByteThreshold,
  largeCalcChainStreamingFormulaThreshold,
  resolveXlsxImportLimits,
  shouldRetryDataOnlyLargeSimpleImport,
  type XlsxImportOptions,
} from './xlsx-import-limits.js'
import { tryInspectLargeSimpleXlsxHeadless } from './xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import {
  hasFullImporterOnlyPackageMetadata,
  shouldBypassLargeSimpleByteThresholdForPackageArtifacts,
} from './xlsx-large-simple-package-artifact-threshold.js'
import { attachImportedXlsxSourceReader, detachImportedXlsxSourceBytes } from './xlsx-source-bytes.js'
import { prepareSheetJsParserXlsxBytesFromZip } from './xlsx-style-only-blank-cells.js'
import { readXlsxZipEntriesLazyFromByteSource, type XlsxZipByteSource, type XlsxZipEntries } from './xlsx-zip.js'

const largeCalcChainStreamingByteThreshold = 5_000_000
const sheetJsBlankStyleStripMinCellCount = 1_000
const requireModule = createRequire(import.meta.url)
const vitestEagerModules = readVitestEagerModules()

type ImportMetaGlob = (patterns: readonly string[], options: { readonly eager: true }) => Readonly<Record<string, unknown>>

declare global {
  interface ImportMeta {
    readonly glob?: ImportMetaGlob
  }
}

interface SheetJsImporterModule {
  readonly importSheetJsWorkbook: (
    bytes: Uint8Array,
    fileName: string,
    contentType: typeof XLSX_CONTENT_TYPE,
    workbookZip: XlsxZipEntries | null,
    sourceBytesForUntouchedExport?: Uint8Array,
  ) => ImportedWorkbook
  readonly importXlsxFromPreparedSheetJsParserData: (
    parserData: Uint8Array,
    fileName: string,
    contentType: typeof XLSX_CONTENT_TYPE,
    workbookZip: XlsxZipEntries,
    sourceByteLength: number,
  ) => ImportedWorkbook
}

let sheetJsImporterModule: SheetJsImporterModule | undefined

export interface XlsxByteSourceImportOptions extends XlsxImportOptions {
  readonly attachSourceReaderForUntouchedExport?: boolean
}

export function importXlsxFromZipByteSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: XlsxByteSourceImportOptions = {},
): ImportedWorkbook {
  const workbookZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
  if (!workbookZip) {
    return importXlsxFromMaterializedSource(source, fileName, options)
  }
  const limits = resolveXlsxImportLimits(options)
  const hasCalcChain = Object.hasOwn(workbookZip, 'xl/calcChain.xml')
  const bypassLargeSimpleByteThreshold =
    shouldBypassLargeSimpleByteThresholdForPackageArtifacts(workbookZip) && !hasFullImporterOnlyPackageMetadata(workbookZip)
  const needsCalcChainFormulaCountInspection =
    hasCalcChain && source.byteLength >= denseSheetJsByteThreshold && source.byteLength < largeCalcChainStreamingByteThreshold
  const inspection =
    limits || needsCalcChainFormulaCountInspection
      ? inspectLargeSimpleXlsxSource(source, fileName, options.limits ? { minByteLength: 0 } : undefined)
      : null
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const hasLargeCalcChainFormulaSet = hasCalcChain && (inspection?.stats.formulaCellCount ?? 0) >= largeCalcChainStreamingFormulaThreshold
  const allowCachedUnsupportedFormulaText =
    hasCalcChain && (source.byteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet)
  const shouldTryLargeSimpleImport =
    !hasCalcChain ||
    source.byteLength >= largeCalcChainStreamingByteThreshold ||
    hasLargeCalcChainFormulaSet ||
    bypassLargeSimpleByteThreshold
  const largeSimpleImportOptions = {
    ...(options.limits || bypassLargeSimpleByteThreshold ? { minByteLength: 0 } : {}),
    allowUnsupportedFormulaText: allowCachedUnsupportedFormulaText,
    allowUnsupportedCellMetadata: allowCachedUnsupportedFormulaText,
    allowPreReleaseSheetFinalization: true,
    releaseArenaAfterMaterialization: true,
    releaseZipSource: true,
    maxMaterializedLazyPackageArtifactBytes: 8 * 1024 * 1024,
  }
  let largeSimpleImport = !shouldTryLargeSimpleImport
    ? null
    : tryImportLargeSimpleXlsx({ byteLength: source.byteLength }, fileName, workbookZip, largeSimpleImportOptions)
  if (!largeSimpleImport && shouldRetryDataOnlyLargeSimpleImport(inspection, source.byteLength, allowCachedUnsupportedFormulaText)) {
    const retryZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
    largeSimpleImport = retryZip
      ? tryImportLargeSimpleXlsx({ byteLength: source.byteLength }, fileName, retryZip, {
          ...largeSimpleImportOptions,
          materializeMetadata: false,
        })
      : null
  }
  if (largeSimpleImport) {
    if (options.attachSourceReaderForUntouchedExport !== false) {
      attachImportedXlsxSourceReader(largeSimpleImport.snapshot, {
        byteLength: source.byteLength,
        readBytes: () => readAllSourceBytes(source),
      })
    }
    return largeSimpleImport
  }
  if (options.attachSourceReaderForUntouchedExport === false) {
    const preparedFallback = importXlsxFromPreparedByteSourceFallback(source, fileName)
    if (preparedFallback) {
      return preparedFallback
    }
  }
  return importXlsxFromMaterializedSource(source, fileName, options)
}

function importXlsxFromPreparedByteSourceFallback(source: XlsxZipByteSource, fileName: string): ImportedWorkbook | null {
  const workbookZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
  if (!workbookZip) {
    return null
  }
  const parserData = prepareSheetJsParserXlsxBytesFromZip(workbookZip, {
    minBlankCellCount: sheetJsBlankStyleStripMinCellCount,
    omitParserIgnoredPackageParts: true,
  })
  if (!parserData) {
    return null
  }
  return loadSheetJsImporterModule().importXlsxFromPreparedSheetJsParserData(
    parserData,
    fileName,
    XLSX_CONTENT_TYPE,
    workbookZip,
    source.byteLength,
  )
}

function importXlsxFromMaterializedSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: XlsxByteSourceImportOptions,
): ImportedWorkbook {
  const imported = loadSheetJsImporterModule().importSheetJsWorkbook(readAllSourceBytes(source), fileName, XLSX_CONTENT_TYPE, null)
  if (options.attachSourceReaderForUntouchedExport === false) {
    detachImportedXlsxSourceBytes(imported.snapshot)
  }
  return imported
}

function loadSheetJsImporterModule(): SheetJsImporterModule {
  sheetJsImporterModule ??= requireSheetJsImporterModule()
  return sheetJsImporterModule
}

function requireSheetJsImporterModule(): SheetJsImporterModule {
  const vitestModule = readVitestEagerModule('./xlsx-sheetjs-import.js')
  if (vitestModule) {
    return readSheetJsImporterModule(vitestModule)
  }
  try {
    return readSheetJsImporterModule(requireModule('./xlsx-sheetjs-import.js'))
  } catch (error) {
    if (!isModuleNotFoundError(error)) {
      throw error
    }
  }
  try {
    return readSheetJsImporterModule(requireModule('../dist/xlsx-sheetjs-import.js'))
  } catch (error) {
    if (!isModuleNotFoundError(error)) {
      throw error
    }
    return readSheetJsImporterModule(requireModule('./xlsx-sheetjs-import.ts'))
  }
}

function readVitestEagerModules(): Readonly<Record<string, unknown>> {
  if (process.env['VITEST'] !== 'true') {
    return {}
  }
  if (!isImportMetaGlob(import.meta.glob)) {
    return {}
  }
  return import.meta.glob(['./xlsx-sheetjs-import.ts'], { eager: true })
}

function readVitestEagerModule(path: string): unknown {
  return vitestEagerModules[path] ?? vitestEagerModules[path.replace(/\.js$/u, '.ts')]
}

function isImportMetaGlob(value: unknown): value is ImportMetaGlob {
  return typeof value === 'function'
}

function isModuleNotFoundError(error: unknown): boolean {
  return isRecord(error) && error['code'] === 'MODULE_NOT_FOUND'
}

function readSheetJsImporterModule(value: unknown): SheetJsImporterModule {
  if (isSheetJsImporterModule(value)) {
    return value
  }
  throw new Error('SheetJS XLSX importer module is missing required exports')
}

function isSheetJsImporterModule(value: unknown): value is SheetJsImporterModule {
  return (
    isRecord(value) &&
    typeof value['importSheetJsWorkbook'] === 'function' &&
    typeof value['importXlsxFromPreparedSheetJsParserData'] === 'function'
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function inspectLargeSimpleXlsxSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: { readonly minByteLength?: number } = {},
): ReturnType<typeof tryInspectLargeSimpleXlsxHeadless> {
  const inspectionZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
  return inspectionZip
    ? tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, inspectionZip, {
        allowUnsupportedWorksheetFeaturesForMetrics: true,
        ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
        releaseZipSource: true,
      })
    : null
}

function readAllSourceBytes(source: XlsxZipByteSource): Uint8Array {
  return source.readRange(0, source.byteLength)
}

export function borrowXlsxZipByteSource(source: XlsxZipByteSource): XlsxZipByteSource {
  return {
    byteLength: source.byteLength,
    readRange: (start, end) => source.readRange(start, end),
    ...(source.readRangeInto
      ? {
          readRangeInto: (start: number, end: number, target: Uint8Array) => source.readRangeInto!(start, end, target),
        }
      : {}),
  }
}
