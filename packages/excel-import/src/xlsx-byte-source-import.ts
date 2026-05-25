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
import {
  hasFullImporterOnlyPackageMetadata,
  shouldBypassLargeSimpleByteThresholdForPackageArtifacts,
} from './xlsx-large-simple-package-artifact-threshold.js'
import type { tryInspectLargeSimpleXlsxHeadless } from './xlsx-large-simple-headless-inspect.js'
import type { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import { attachImportedXlsxSourceReader, detachImportedXlsxSourceBytes } from './xlsx-source-bytes.js'
import type { prepareSheetJsParserXlsxBytesFromZip } from './xlsx-style-only-blank-cells.js'
import { readXlsxZipEntriesLazy, readXlsxZipEntriesLazyFromByteSource, type XlsxZipByteSource, type XlsxZipEntries } from './xlsx-zip.js'

const largeCalcChainStreamingByteThreshold = 5_000_000
const sheetJsBlankStyleStripMinCellCount = 1_000
const requireModule = createRequire(import.meta.url)
const bundledLocalModules = readBundledLocalModules()

type ImportMetaGlob = (patterns: readonly string[], options: { readonly eager: true }) => Readonly<Record<string, unknown>>

declare global {
  interface ImportMeta {
    readonly glob: ImportMetaGlob
  }
}

interface SheetJsImporterModule {
  readonly importSheetJsWorkbook: (
    bytes: Uint8Array,
    fileName: string,
    contentType: typeof XLSX_CONTENT_TYPE,
    workbookZip: XlsxZipEntries | null,
    sourceBytesForUntouchedExport?: Uint8Array,
    options?: XlsxImportOptions,
  ) => ImportedWorkbook
  readonly importXlsxFromPreparedSheetJsParserData: (
    parserData: Uint8Array,
    fileName: string,
    contentType: typeof XLSX_CONTENT_TYPE,
    workbookZip: XlsxZipEntries,
    sourceByteLength: number,
  ) => ImportedWorkbook
}

type TryImportLargeSimpleXlsx = typeof tryImportLargeSimpleXlsx
type TryInspectLargeSimpleXlsxHeadless = typeof tryInspectLargeSimpleXlsxHeadless
type PrepareSheetJsParserXlsxBytesFromZip = typeof prepareSheetJsParserXlsxBytesFromZip

interface LargeSimpleImportModule {
  readonly tryImportLargeSimpleXlsx: TryImportLargeSimpleXlsx
}

interface LargeSimpleInspectModule {
  readonly tryInspectLargeSimpleXlsxHeadless: TryInspectLargeSimpleXlsxHeadless
}

interface StyleOnlyBlankCellsModule {
  readonly prepareSheetJsParserXlsxBytesFromZip: PrepareSheetJsParserXlsxBytesFromZip
}

let sheetJsImporterModule: SheetJsImporterModule | undefined
let largeSimpleImportModule: LargeSimpleImportModule | undefined
let largeSimpleInspectModule: LargeSimpleInspectModule | undefined
let styleOnlyBlankCellsModule: StyleOnlyBlankCellsModule | undefined

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
  const hasExternalWorkbookCompanions = (options.externalWorkbooks?.length ?? 0) > 0
  const shouldTryLargeSimpleImport =
    !hasExternalWorkbookCompanions &&
    (!hasCalcChain ||
      source.byteLength >= largeCalcChainStreamingByteThreshold ||
      hasLargeCalcChainFormulaSet ||
      bypassLargeSimpleByteThreshold)
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
    : loadLargeSimpleImportModule().tryImportLargeSimpleXlsx(
        { byteLength: source.byteLength },
        fileName,
        workbookZip,
        largeSimpleImportOptions,
      )
  if (!largeSimpleImport && shouldRetryDataOnlyLargeSimpleImport(inspection, source.byteLength, allowCachedUnsupportedFormulaText)) {
    const retryZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
    largeSimpleImport = retryZip
      ? loadLargeSimpleImportModule().tryImportLargeSimpleXlsx({ byteLength: source.byteLength }, fileName, retryZip, {
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
  if (options.attachSourceReaderForUntouchedExport === false && !hasExternalWorkbookCompanions) {
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
  const parserData = loadStyleOnlyBlankCellsModule().prepareSheetJsParserXlsxBytesFromZip(workbookZip, {
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
  const data = readAllSourceBytes(source)
  const imported = loadSheetJsImporterModule().importSheetJsWorkbook(
    data,
    fileName,
    XLSX_CONTENT_TYPE,
    readMaterializedWorkbookZip(data),
    undefined,
    options,
  )
  if (options.attachSourceReaderForUntouchedExport === false) {
    detachImportedXlsxSourceBytes(imported.snapshot)
  }
  return imported
}

function readMaterializedWorkbookZip(data: Uint8Array): XlsxZipEntries | null {
  try {
    return readXlsxZipEntriesLazy(data)
  } catch {
    return null
  }
}

function loadSheetJsImporterModule(): SheetJsImporterModule {
  sheetJsImporterModule ??= requireSheetJsImporterModule()
  return sheetJsImporterModule
}

function loadLargeSimpleImportModule(): LargeSimpleImportModule {
  largeSimpleImportModule ??= readLargeSimpleImportModule(
    readBundledLocalModule('./xlsx-large-simple-import.js') ?? requireLocalModule('./xlsx-large-simple-import.js'),
  )
  return largeSimpleImportModule
}

function loadLargeSimpleInspectModule(): LargeSimpleInspectModule {
  largeSimpleInspectModule ??= readLargeSimpleInspectModule(
    readBundledLocalModule('./xlsx-large-simple-headless-inspect.js') ?? requireLocalModule('./xlsx-large-simple-headless-inspect.js'),
  )
  return largeSimpleInspectModule
}

function loadStyleOnlyBlankCellsModule(): StyleOnlyBlankCellsModule {
  styleOnlyBlankCellsModule ??= readStyleOnlyBlankCellsModule(
    readBundledLocalModule('./xlsx-style-only-blank-cells.js') ?? requireLocalModule('./xlsx-style-only-blank-cells.js'),
  )
  return styleOnlyBlankCellsModule
}

function requireSheetJsImporterModule(): SheetJsImporterModule {
  const bundledModule = readBundledLocalModule('./xlsx-sheetjs-import.js')
  if (bundledModule) {
    return readSheetJsImporterModule(bundledModule)
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

function requireLocalModule(jsPath: string): unknown {
  try {
    return requireModule(jsPath)
  } catch (error) {
    if (!isModuleNotFoundError(error)) {
      throw error
    }
  }
  try {
    return requireModule(`../dist/${jsPath.slice(2)}`)
  } catch (error) {
    if (!isModuleNotFoundError(error)) {
      throw error
    }
    return requireModule(jsPath.replace(/\.js$/u, '.ts'))
  }
}

function readBundledLocalModules(): Readonly<Record<string, unknown>> {
  try {
    return import.meta.glob(
      [
        './xlsx-sheetjs-import.ts',
        './xlsx-large-simple-import.ts',
        './xlsx-large-simple-headless-inspect.ts',
        './xlsx-style-only-blank-cells.ts',
      ],
      { eager: true },
    )
  } catch {
    return {}
  }
}

function readBundledLocalModule(path: string): unknown {
  return bundledLocalModules[path] ?? bundledLocalModules[path.replace(/\.js$/u, '.ts')]
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

function readLargeSimpleImportModule(value: unknown): LargeSimpleImportModule {
  const loadedTryImportLargeSimpleXlsx = isRecord(value) ? value['tryImportLargeSimpleXlsx'] : undefined
  if (isLargeSimpleImportFunction(loadedTryImportLargeSimpleXlsx)) {
    return { tryImportLargeSimpleXlsx: loadedTryImportLargeSimpleXlsx }
  }
  throw new Error('Large-simple XLSX import module is missing required exports')
}

function readLargeSimpleInspectModule(value: unknown): LargeSimpleInspectModule {
  const loadedTryInspectLargeSimpleXlsxHeadless = isRecord(value) ? value['tryInspectLargeSimpleXlsxHeadless'] : undefined
  if (isLargeSimpleInspectFunction(loadedTryInspectLargeSimpleXlsxHeadless)) {
    return { tryInspectLargeSimpleXlsxHeadless: loadedTryInspectLargeSimpleXlsxHeadless }
  }
  throw new Error('Large-simple XLSX inspect module is missing required exports')
}

function readStyleOnlyBlankCellsModule(value: unknown): StyleOnlyBlankCellsModule {
  const loadedPrepareSheetJsParserXlsxBytesFromZip = isRecord(value) ? value['prepareSheetJsParserXlsxBytesFromZip'] : undefined
  if (isStyleOnlyBlankCellsFunction(loadedPrepareSheetJsParserXlsxBytesFromZip)) {
    return { prepareSheetJsParserXlsxBytesFromZip: loadedPrepareSheetJsParserXlsxBytesFromZip }
  }
  throw new Error('SheetJS parser-preparation module is missing required exports')
}

function isLargeSimpleImportFunction(value: unknown): value is TryImportLargeSimpleXlsx {
  return typeof value === 'function'
}

function isLargeSimpleInspectFunction(value: unknown): value is TryInspectLargeSimpleXlsxHeadless {
  return typeof value === 'function'
}

function isStyleOnlyBlankCellsFunction(value: unknown): value is PrepareSheetJsParserXlsxBytesFromZip {
  return typeof value === 'function'
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
): ReturnType<TryInspectLargeSimpleXlsxHeadless> {
  const inspectionZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
  return inspectionZip
    ? loadLargeSimpleInspectModule().tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, inspectionZip, {
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
    ...(source.inflateRawRange ? { inflateRawRange: (start: number, end: number) => source.inflateRawRange!(start, end) } : {}),
  }
}
