import type { CsvParseOptions } from '@bilig/core'
import { importCsv as importCsvSnapshot } from './csv-import.js'
import type { ImportedWorkbook } from './workbook-import-result.js'
import {
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
} from './workbook-import-content-types.js'
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
import { importSheetJsWorkbook } from './xlsx-sheetjs-import.js'
import { attachImportedXlsxSourceBytes } from './xlsx-source-bytes.js'
import { readXlsxZipEntriesLazy, type XlsxZipEntries } from './xlsx-zip.js'

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

export function importCsv(csv: string, fileName: string, options?: CsvParseOptions): ImportedWorkbook {
  return importCsvSnapshot(csv, fileName, options)
}

export function inspectXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): XlsxHeadlessInspectResult | null {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return inspectLargeSimpleXlsxSource(data, fileName, { minByteLength: 0 })
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const sourceByteLength = data.byteLength
  const limits = resolveXlsxImportLimits(options)
  const workbookZip = readValidXlsxZipContainer(data)
  const hasCalcChain = Object.hasOwn(workbookZip, 'xl/calcChain.xml')
  const bypassLargeSimpleByteThreshold =
    shouldBypassLargeSimpleByteThresholdForPackageArtifacts(workbookZip) && !hasFullImporterOnlyPackageMetadata(workbookZip)
  const needsCalcChainFormulaCountInspection =
    hasCalcChain && sourceByteLength >= denseSheetJsByteThreshold && sourceByteLength < largeCalcChainStreamingByteThreshold
  const inspection =
    limits || needsCalcChainFormulaCountInspection
      ? inspectLargeSimpleXlsxSource(data, fileName, options.limits ? { minByteLength: 0 } : undefined)
      : null
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const hasLargeCalcChainFormulaSet = hasCalcChain && (inspection?.stats.formulaCellCount ?? 0) >= largeCalcChainStreamingFormulaThreshold
  const allowCachedUnsupportedFormulaText =
    hasCalcChain && (sourceByteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet)
  const shouldTryLargeSimpleImport =
    !hasCalcChain ||
    sourceByteLength >= largeCalcChainStreamingByteThreshold ||
    hasLargeCalcChainFormulaSet ||
    bypassLargeSimpleByteThreshold
  const largeSimpleImportOptions = {
    ...(options.limits || bypassLargeSimpleByteThreshold ? { minByteLength: 0 } : {}),
    allowUnsupportedFormulaText: allowCachedUnsupportedFormulaText,
    allowUnsupportedCellMetadata: allowCachedUnsupportedFormulaText,
    skipBroadBlankStyleCells: true,
    includeCellCoordinates: true,
    allowPreReleaseSheetFinalization: true,
    releaseArenaAfterMaterialization: true,
    releaseZipSource: true,
    maxMaterializedLazyPackageArtifactBytes: 8 * 1024 * 1024,
  }
  let largeSimpleImport = shouldTryLargeSimpleImport
    ? tryImportLargeSimpleXlsx({ byteLength: sourceByteLength }, fileName, workbookZip, largeSimpleImportOptions)
    : null
  if (!largeSimpleImport && shouldRetryDataOnlyLargeSimpleImport(inspection, sourceByteLength, allowCachedUnsupportedFormulaText)) {
    largeSimpleImport = tryImportLargeSimpleXlsx({ byteLength: sourceByteLength }, fileName, readValidXlsxZipContainer(data), {
      ...largeSimpleImportOptions,
      materializeMetadata: false,
    })
  }
  if (largeSimpleImport) {
    attachImportedXlsxSourceBytes(largeSimpleImport.snapshot, data)
    return largeSimpleImport
  }
  return importSheetJsWorkbook(data, fileName, XLSX_CONTENT_TYPE, readValidXlsxZipContainer(data), data)
}

export function importXlsm(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const workbookZip = readValidXlsxZipContainer(data)
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

function inspectLargeSimpleXlsxSource(
  data: Uint8Array,
  fileName: string,
  options: { readonly minByteLength?: number } = {},
): LargeSimpleXlsxHeadlessInspectResult | null {
  return tryInspectLargeSimpleXlsxHeadless({ byteLength: data.byteLength }, fileName, readValidXlsxZipContainer(data), {
    allowUnsupportedWorksheetFeaturesForMetrics: true,
    ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
    releaseZipSource: true,
  })
}

function readValidXlsxZipContainer(bytes: Uint8Array): XlsxZipEntries {
  try {
    const zip = readXlsxZipEntriesLazy(bytes)
    void zip['xl/workbook.xml']
    return zip
  } catch {
    throw new InvalidXlsxZipContainerError()
  }
}
