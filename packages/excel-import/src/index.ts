import { createRequire } from 'node:module'
import type { Unzipped } from 'fflate'
import type { CsvParseOptions } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
  type ExcelWorkbookImportContentType,
} from './workbook-import-content-types.js'
import type { ImportedWorkbook } from './workbook-import-result.js'
import {
  assertXlsxInspectionWithinMaterializationLimits,
  denseSheetJsByteThreshold,
  largeCalcChainStreamingFormulaThreshold,
  resolveXlsxImportLimits,
  shouldRetryDataOnlyLargeSimpleImport,
  type XlsxImportOptions,
} from './xlsx-import-limits.js'
import { tryInspectLargeSimpleXlsxHeadless, type LargeSimpleXlsxHeadlessInspectResult } from './xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import {
  hasFullImporterOnlyPackageMetadata,
  shouldBypassLargeSimpleByteThresholdForPackageArtifacts,
} from './xlsx-large-simple-package-artifact-threshold.js'
import { releaseOwnedXlsxSourceBytes, type OwnedXlsxSourceBytes } from './xlsx-owned-source-release.js'
import {
  attachImportedXlsxSourceBytes,
  attachImportedXlsxSourceReader,
  createTempFileImportedXlsxSourceReader,
  type ImportedXlsxSourceReader,
} from './xlsx-source-bytes.js'
import { readLazyXlsxZipSource, readXlsxZipEntries, readXlsxZipEntriesLazy, type XlsxZipByteSource } from './xlsx-zip.js'

export { manualCalculationModeWarning, precisionAsDisplayedCalculationWarning } from './xlsx-calculation-settings.js'
export {
  dataTableFormulasWarning,
  externalPivotCachesWarning,
  externalWorkbookReferencesWarning,
  macroExecutionDeclinedWarning,
  volatileFormulasWarning,
} from './xlsx-import-warnings.js'
export { readImportedXlsxCellStyle } from './xlsx-import-cell-styles.js'
export { XlsxImportSizeLimitExceededError } from './xlsx-import-limits.js'
export type { ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'
export type { ImportedWorkbookPreview } from './workbook-import-preview.js'
export type { ImportedWorkbook } from './workbook-import-result.js'
export type { XlsxImportLimits, XlsxImportOptions } from './xlsx-import-limits.js'
export {
  CSV_CONTENT_TYPE,
  EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES,
  LEGACY_XLS_CONTENT_TYPE,
  WORKBOOK_IMPORT_CONTENT_TYPES,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
} from './workbook-import-content-types.js'
export type { ExcelWorkbookImportContentType, WorkbookImportContentType } from './workbook-import-content-types.js'

const largeCalcChainStreamingByteThreshold = 5_000_000
const largeSimpleInMemoryUntouchedExportSourceLimit = 8 * 1024 * 1024
const requireModule = createRequire(import.meta.url)
const vitestEagerModules = readVitestEagerModules()

type ImportMetaGlob = (patterns: readonly string[], options: { readonly eager: true }) => Readonly<Record<string, unknown>>

declare global {
  interface ImportMeta {
    readonly glob?: ImportMetaGlob
  }
}

interface XlsxExportModule {
  readonly exportXlsx: (snapshot: WorkbookSnapshot) => Uint8Array
}

interface CsvImportModule {
  readonly importCsv: (csv: string, fileName: string, options?: CsvParseOptions) => ImportedWorkbook
}

interface SheetJsImportModule {
  readonly importSheetJsWorkbook: (
    data: Uint8Array,
    fileName: string,
    contentType: ExcelWorkbookImportContentType,
    workbookZip: Unzipped | null,
    sourceBytesForUntouchedExport?: Uint8Array,
  ) => ImportedWorkbook
  readonly importXlsxFromPreparedSheetJsParserData: (
    parserData: Uint8Array,
    fileName: string,
    contentType: ExcelWorkbookImportContentType,
    workbookZip: Unzipped | null,
    sourceFileSizeBytes: number,
  ) => ImportedWorkbook
}

let xlsxExportModule: XlsxExportModule | undefined
let csvImportModule: CsvImportModule | undefined
let sheetJsImportModule: SheetJsImportModule | undefined

function loadXlsxExportModule(): XlsxExportModule {
  xlsxExportModule ??= readXlsxExportModule(readVitestEagerModule('./xlsx-export.js') ?? requireLocalModule('./xlsx-export.js'))
  return xlsxExportModule
}

function loadCsvImportModule(): CsvImportModule {
  csvImportModule ??= readCsvImportModule(readVitestEagerModule('./csv-import.js') ?? requireLocalModule('./csv-import.js'))
  return csvImportModule
}

function loadSheetJsImportModule(): SheetJsImportModule {
  sheetJsImportModule ??= readSheetJsImportModule(
    readVitestEagerModule('./xlsx-sheetjs-import.js') ?? requireLocalModule('./xlsx-sheetjs-import.js'),
  )
  return sheetJsImportModule
}

function readVitestEagerModules(): Readonly<Record<string, unknown>> {
  if (process.env['VITEST'] !== 'true') {
    return {}
  }
  if (!isImportMetaGlob(import.meta.glob)) {
    return {}
  }
  return import.meta.glob(['./xlsx-export.ts', './csv-import.ts', './xlsx-sheetjs-import.ts'], { eager: true })
}

function readVitestEagerModule(path: string): unknown {
  return vitestEagerModules[path] ?? vitestEagerModules[path.replace(/\.js$/u, '.ts')]
}

function isImportMetaGlob(value: unknown): value is ImportMetaGlob {
  return typeof value === 'function'
}

function requireLocalModule(jsPath: string): unknown {
  try {
    return requireModule(jsPath)
  } catch (error) {
    if (!isModuleNotFoundError(error)) {
      throw error
    }
    return requireModule(`../dist/${jsPath.slice(2)}`)
  }
}

function isModuleNotFoundError(error: unknown): boolean {
  return isRecord(error) && error['code'] === 'MODULE_NOT_FOUND'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isXlsxExportFunction(value: unknown): value is XlsxExportModule['exportXlsx'] {
  return typeof value === 'function'
}

function isCsvImportFunction(value: unknown): value is CsvImportModule['importCsv'] {
  return typeof value === 'function'
}

function isSheetJsImportFunction(value: unknown): value is SheetJsImportModule['importSheetJsWorkbook'] {
  return typeof value === 'function'
}

function isPreparedSheetJsImportFunction(value: unknown): value is SheetJsImportModule['importXlsxFromPreparedSheetJsParserData'] {
  return typeof value === 'function'
}

function readXlsxExportModule(value: unknown): XlsxExportModule {
  const loadedExportXlsx = isRecord(value) ? value['exportXlsx'] : undefined
  if (isXlsxExportFunction(loadedExportXlsx)) {
    return { exportXlsx: loadedExportXlsx }
  }
  throw new Error('XLSX export module is missing required exports')
}

function readCsvImportModule(value: unknown): CsvImportModule {
  const loadedImportCsv = isRecord(value) ? value['importCsv'] : undefined
  if (isCsvImportFunction(loadedImportCsv)) {
    return { importCsv: loadedImportCsv }
  }
  throw new Error('CSV import module is missing required exports')
}

function readSheetJsImportModule(value: unknown): SheetJsImportModule {
  const loadedImportSheetJsWorkbook = isRecord(value) ? value['importSheetJsWorkbook'] : undefined
  const loadedPreparedImport = isRecord(value) ? value['importXlsxFromPreparedSheetJsParserData'] : undefined
  if (isSheetJsImportFunction(loadedImportSheetJsWorkbook) && isPreparedSheetJsImportFunction(loadedPreparedImport)) {
    return {
      importSheetJsWorkbook: loadedImportSheetJsWorkbook,
      importXlsxFromPreparedSheetJsParserData: loadedPreparedImport,
    }
  }
  throw new Error('SheetJS XLSX import module is missing required exports')
}

export function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array {
  return loadXlsxExportModule().exportXlsx(snapshot)
}

export function importCsv(csv: string, fileName: string, options?: CsvParseOptions): ImportedWorkbook {
  return loadCsvImportModule().importCsv(csv, fileName, options)
}

function importSheetJsWorkbook(
  data: Uint8Array,
  fileName: string,
  contentType: ExcelWorkbookImportContentType,
  workbookZip: Unzipped | null,
  sourceBytesForUntouchedExport?: Uint8Array,
): ImportedWorkbook {
  return loadSheetJsImportModule().importSheetJsWorkbook(data, fileName, contentType, workbookZip, sourceBytesForUntouchedExport)
}

export function importXlsxFromPreparedSheetJsParserData(
  parserData: Uint8Array,
  fileName: string,
  contentType: ExcelWorkbookImportContentType,
  workbookZip: Unzipped | null,
  sourceFileSizeBytes: number,
): ImportedWorkbook {
  return loadSheetJsImportModule().importXlsxFromPreparedSheetJsParserData(
    parserData,
    fileName,
    contentType,
    workbookZip,
    sourceFileSizeBytes,
  )
}

export type CsvImportOptions = CsvParseOptions
export type XlsxHeadlessInspectResult = LargeSimpleXlsxHeadlessInspectResult

export interface WorkbookImportFileOptions {
  csv?: CsvImportOptions
  xlsx?: XlsxImportOptions
}

export class InvalidXlsxZipContainerError extends Error {
  constructor() {
    super('Invalid or corrupt XLSX zip container')
    this.name = 'InvalidXlsxZipContainerError'
  }
}

function readValidXlsxZipContainer(bytes: Uint8Array, mode: 'eager' | 'lazy' = 'eager'): Unzipped {
  try {
    const zip = mode === 'lazy' ? readXlsxZipEntriesLazy(bytes) : readXlsxZipEntries(bytes)
    void zip['xl/workbook.xml']
    return zip
  } catch {
    throw new InvalidXlsxZipContainerError()
  }
}

function inspectLargeSimpleXlsxSource(
  data: Uint8Array,
  fileName: string,
  options: { readonly minByteLength?: number } = {},
): LargeSimpleXlsxHeadlessInspectResult | null {
  const inspectionZip = readValidXlsxZipContainer(data, 'lazy')
  return tryInspectLargeSimpleXlsxHeadless({ byteLength: data.byteLength }, fileName, inspectionZip, {
    allowUnsupportedWorksheetFeaturesForMetrics: true,
    ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
    releaseZipSource: true,
  })
}

export function inspectXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): XlsxHeadlessInspectResult | null {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return inspectLargeSimpleXlsxSource(data, fileName, { minByteLength: 0 })
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const ownedSource: OwnedXlsxSourceBytes = { bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes) }
  const sourceByteLength = ownedSource.bytes.byteLength
  let spooledUntouchedExportSource: (ImportedXlsxSourceReader & XlsxZipByteSource) | undefined =
    sourceByteLength > largeSimpleInMemoryUntouchedExportSourceLimit ? createTempFileImportedXlsxSourceReader(ownedSource.bytes) : undefined
  const limits = resolveXlsxImportLimits(options)
  const inspectionOptions = options.limits ? { minByteLength: 0 } : undefined
  const workbookZip = readValidXlsxZipContainer(ownedSource.bytes, 'lazy')
  const hasCalcChain = Object.hasOwn(workbookZip, 'xl/calcChain.xml')
  const bypassLargeSimpleByteThreshold =
    shouldBypassLargeSimpleByteThresholdForPackageArtifacts(workbookZip) && !hasFullImporterOnlyPackageMetadata(workbookZip)
  const needsCalcChainFormulaCountInspection =
    hasCalcChain && sourceByteLength >= denseSheetJsByteThreshold && sourceByteLength < largeCalcChainStreamingByteThreshold
  const inspection =
    limits || needsCalcChainFormulaCountInspection ? inspectLargeSimpleXlsxSource(ownedSource.bytes, fileName, inspectionOptions) : null
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const hasLargeCalcChainFormulaSet = hasCalcChain && (inspection?.stats.formulaCellCount ?? 0) >= largeCalcChainStreamingFormulaThreshold
  const allowCachedUnsupportedFormulaText =
    hasCalcChain && (sourceByteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet)
  const shouldTryLargeSimpleImport =
    !hasCalcChain ||
    sourceByteLength >= largeCalcChainStreamingByteThreshold ||
    hasLargeCalcChainFormulaSet ||
    bypassLargeSimpleByteThreshold
  const releaseOwnedSourceBytesForLargeSimpleImport =
    spooledUntouchedExportSource || (bypassLargeSimpleByteThreshold && sourceByteLength < denseSheetJsByteThreshold)
      ? () => releaseOwnedXlsxSourceBytes(ownedSource, (releasedBytes) => (bytes = releasedBytes))
      : undefined
  const largeSimpleImportOptions = {
    ...(options.limits || bypassLargeSimpleByteThreshold ? { minByteLength: 0 } : {}),
    allowUnsupportedFormulaText: allowCachedUnsupportedFormulaText,
    allowUnsupportedCellMetadata: allowCachedUnsupportedFormulaText,
    allowPreReleaseSheetFinalization: releaseOwnedSourceBytesForLargeSimpleImport === undefined,
    releaseArenaAfterMaterialization: true,
    releaseZipSource: true,
    maxMaterializedLazyPackageArtifactBytes: 8 * 1024 * 1024,
    ...(spooledUntouchedExportSource ? { replacementZipSource: spooledUntouchedExportSource } : {}),
    ...(releaseOwnedSourceBytesForLargeSimpleImport ? { releaseOwnedSourceBytes: releaseOwnedSourceBytesForLargeSimpleImport } : {}),
  }
  let largeSimpleImport = shouldTryLargeSimpleImport
    ? tryImportLargeSimpleXlsx({ byteLength: sourceByteLength }, fileName, workbookZip, largeSimpleImportOptions)
    : null
  if (!largeSimpleImport && shouldRetryDataOnlyLargeSimpleImport(inspection, sourceByteLength, allowCachedUnsupportedFormulaText)) {
    largeSimpleImport = tryImportLargeSimpleXlsx(
      { byteLength: sourceByteLength },
      fileName,
      readValidXlsxZipContainer(ownedSource.bytes, 'lazy'),
      {
        ...largeSimpleImportOptions,
        materializeMetadata: false,
      },
    )
  }
  if (largeSimpleImport) {
    if (ownedSource.bytes.byteLength > 0) {
      spooledUntouchedExportSource?.release?.()
      spooledUntouchedExportSource = undefined
      attachImportedXlsxSourceBytes(largeSimpleImport.snapshot, ownedSource.bytes)
    } else if (spooledUntouchedExportSource) {
      attachImportedXlsxSourceReader(largeSimpleImport.snapshot, spooledUntouchedExportSource)
      spooledUntouchedExportSource = undefined
    }
    return largeSimpleImport
  }
  const fallbackData =
    ownedSource.bytes.byteLength > 0
      ? ownedSource.bytes
      : spooledUntouchedExportSource
        ? spooledUntouchedExportSource.readBytes()
        : readLazyXlsxZipSource(workbookZip)
  spooledUntouchedExportSource?.release?.()
  spooledUntouchedExportSource = undefined
  if (!fallbackData || fallbackData.byteLength === 0) {
    throw new InvalidXlsxZipContainerError()
  }
  const imported = importSheetJsWorkbook(
    fallbackData,
    fileName,
    XLSX_CONTENT_TYPE,
    readValidXlsxZipContainer(fallbackData, 'lazy'),
    ownedSource.bytes,
  )
  return imported
}

export function importXlsm(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const workbookZip = readValidXlsxZipContainer(data, 'lazy')
  return importSheetJsWorkbook(data, fileName, XLSM_CONTENT_TYPE, workbookZip, data)
}

export function importXlsb(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return importSheetJsWorkbook(data, fileName, XLSB_CONTENT_TYPE, null)
}

export function importXls(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return importSheetJsWorkbook(data, fileName, LEGACY_XLS_CONTENT_TYPE, null)
}

export function importWorkbookFile(
  bytes: Uint8Array | ArrayBuffer,
  fileName: string,
  contentType: string,
  options: WorkbookImportFileOptions = {},
): ImportedWorkbook {
  const normalizedContentType = normalizeWorkbookImportContentType(contentType)
  if (normalizedContentType === XLSX_CONTENT_TYPE) {
    return importXlsx(bytes, fileName, options.xlsx)
  }
  if (normalizedContentType === XLSM_CONTENT_TYPE) {
    return importXlsm(bytes, fileName)
  }
  if (normalizedContentType === XLSB_CONTENT_TYPE) {
    return importXlsb(bytes, fileName)
  }
  if (normalizedContentType === LEGACY_XLS_CONTENT_TYPE) {
    return importXls(bytes, fileName)
  }
  if (normalizedContentType === CSV_CONTENT_TYPE) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    return importCsv(new TextDecoder().decode(data), fileName, options.csv)
  }
  throw new Error('Unsupported workbook import content type')
}
