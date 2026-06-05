import type { LargeSimpleXlsxHeadlessInspectResult } from './xlsx-large-simple-headless-inspect.js'
import type { LargeSimpleXlsxImportOptions } from './xlsx-large-simple-import-types.js'
import {
  hasFullImporterOnlyPackageMetadata,
  shouldBypassLargeSimpleByteThresholdForPackageArtifacts,
} from './xlsx-large-simple-package-artifact-threshold.js'
import type { XlsxZipEntries } from './xlsx-zip.js'

export const denseSheetJsByteThreshold = 1_000_000
export const largeCalcChainStreamingFormulaThreshold = 50_000
export const largeCalcChainStreamingByteThreshold = 5_000_000
export const largeSimpleInMemoryUntouchedExportSourceLimit = 8 * 1024 * 1024
export const largeSimpleLazyPackageArtifactMaterializationLimit = 8 * 1024 * 1024

export interface XlsxImportLimits {
  maxMaterializedCells?: number
  maxMaterializedFormulaCells?: number
}

const defaultSheetJsFallbackImportLimits: Required<XlsxImportLimits> = {
  maxMaterializedCells: denseSheetJsByteThreshold,
  maxMaterializedFormulaCells: largeCalcChainStreamingFormulaThreshold,
}

export interface XlsxExternalWorkbookInput {
  readonly bytes: Uint8Array | ArrayBuffer
  readonly fileName?: string
  readonly workbookName?: string
  readonly target?: string
}

export type XlsxExternalLinkCacheArtifactMode = 'preserve-existing' | 'replace-refreshed'

export type XlsxExternalWorkbookHydrationStatus = 'refreshed' | 'skipped-no-match' | 'skipped-ambiguous-match' | 'skipped-empty-refresh'

export type XlsxExternalWorkbookHydrationMatchKind = 'exact-target' | 'unique-workbook-identity'

export interface XlsxExternalWorkbookHydrationReferenceDiagnostic {
  readonly bookIndex: number
  readonly workbookName?: string
  readonly target?: string
  readonly status: XlsxExternalWorkbookHydrationStatus
  readonly candidateCount: number
  readonly referenceCandidateCount?: number
  readonly matchKind?: XlsxExternalWorkbookHydrationMatchKind
  readonly matchedFileName?: string
  readonly matchedWorkbookName?: string
  readonly matchedTarget?: string
  readonly refreshedSheetCount?: number
  readonly refreshedCellCount?: number
}

export interface XlsxExternalWorkbookHydrationDiagnostics {
  readonly externalWorkbookCount: number
  readonly externalReferenceCount: number
  readonly refreshedBookIndices: readonly number[]
  readonly refreshedSheetCount: number
  readonly refreshedCellCount: number
  readonly skippedNoMatchCount: number
  readonly skippedAmbiguousMatchCount: number
  readonly skippedEmptyRefreshCount: number
  readonly references: readonly XlsxExternalWorkbookHydrationReferenceDiagnostic[]
}

export interface ImportedWorkbookDiagnostics {
  readonly externalWorkbookHydration?: XlsxExternalWorkbookHydrationDiagnostics
}

export interface XlsxImportOptions {
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly externalLinkCacheArtifactMode?: XlsxExternalLinkCacheArtifactMode
  readonly limits?: XlsxImportLimits | false
  readonly preferNativeSimpleImport?: boolean
  readonly nativeOnly?: boolean
}

export interface XlsxImportRoutePlan {
  readonly bypassLargeSimpleByteThreshold: boolean
  readonly hasExternalWorkbookCompanions: boolean
  readonly inspectionOptions: { readonly minByteLength?: number } | undefined
  readonly shouldInspectBeforeLargeSimpleRouting: boolean
  readonly shouldInspectBeforeSheetJsFallback: boolean
  readonly shouldTryLargeSimpleImport: boolean
  readonly shouldRetryDataOnlyLargeSimpleImport: boolean
  readonly createLargeSimpleImportOptions: (overrides?: LargeSimpleXlsxImportOptions) => LargeSimpleXlsxImportOptions
}

export class XlsxImportSizeLimitExceededError extends Error {
  readonly limits: Required<XlsxImportLimits>
  readonly stats: LargeSimpleXlsxHeadlessInspectResult['stats']
  readonly reason: 'cell-count' | 'formula-cell-count'

  constructor(args: {
    reason: 'cell-count' | 'formula-cell-count'
    limits: Required<XlsxImportLimits>
    stats: LargeSimpleXlsxHeadlessInspectResult['stats']
  }) {
    const observed = args.reason === 'cell-count' ? args.stats.cellCount : args.stats.formulaCellCount
    const limit = args.reason === 'cell-count' ? args.limits.maxMaterializedCells : args.limits.maxMaterializedFormulaCells
    super(
      `XLSX import exceeds the materialized ${args.reason === 'cell-count' ? 'cell' : 'formula cell'} limit ` +
        `(${observed.toLocaleString('en-US')} > ${limit.toLocaleString('en-US')}). ` +
        'Use inspectXlsx() for bounded metadata, raise importXlsx limits explicitly, or split the workbook before materializing it.',
    )
    this.name = 'XlsxImportSizeLimitExceededError'
    this.reason = args.reason
    this.limits = args.limits
    this.stats = args.stats
  }
}

export function resolveXlsxImportLimits(options: XlsxImportOptions): Required<XlsxImportLimits> | null {
  if (options.limits === false || options.limits === undefined) {
    return null
  }
  return {
    maxMaterializedCells: options.limits.maxMaterializedCells ?? Number.POSITIVE_INFINITY,
    maxMaterializedFormulaCells: options.limits.maxMaterializedFormulaCells ?? Number.POSITIVE_INFINITY,
  }
}

export function assertXlsxInspectionWithinMaterializationLimits(
  inspection: LargeSimpleXlsxHeadlessInspectResult | null,
  limits: Required<XlsxImportLimits> | null,
): void {
  if (!inspection || !limits) {
    return
  }
  if (inspection.stats.cellCount > limits.maxMaterializedCells) {
    throw new XlsxImportSizeLimitExceededError({ reason: 'cell-count', limits, stats: inspection.stats })
  }
  if (inspection.stats.formulaCellCount > limits.maxMaterializedFormulaCells) {
    throw new XlsxImportSizeLimitExceededError({ reason: 'formula-cell-count', limits, stats: inspection.stats })
  }
}

export function assertXlsxSheetJsFallbackWithinMaterializationLimits(
  inspection: LargeSimpleXlsxHeadlessInspectResult | null,
  options: XlsxImportOptions,
): void {
  assertXlsxInspectionWithinMaterializationLimits(inspection, resolveXlsxSheetJsFallbackLimits(options))
}

export function planXlsxImportRoute(args: {
  readonly workbookZip: XlsxZipEntries
  readonly sourceByteLength: number
  readonly options: XlsxImportOptions
  readonly inspection: LargeSimpleXlsxHeadlessInspectResult | null
}): XlsxImportRoutePlan {
  const hasCalcChain = Object.hasOwn(args.workbookZip, 'xl/calcChain.xml')
  const bypassLargeSimpleByteThreshold =
    shouldBypassLargeSimpleByteThresholdForPackageArtifacts(args.workbookZip) && !hasFullImporterOnlyPackageMetadata(args.workbookZip)
  const hasMaterializationLimits = args.options.limits !== undefined && args.options.limits !== false
  const nativeOnly = args.options.nativeOnly === true
  const needsCalcChainFormulaCountInspection =
    hasCalcChain && args.sourceByteLength >= denseSheetJsByteThreshold && args.sourceByteLength < largeCalcChainStreamingByteThreshold
  const hasLargeCalcChainFormulaSet =
    hasCalcChain && (args.inspection?.stats.formulaCellCount ?? 0) >= largeCalcChainStreamingFormulaThreshold
  const allowCachedUnsupportedFormulaText =
    hasCalcChain && (args.sourceByteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet)
  const hasExternalWorkbookCompanions = (args.options.externalWorkbooks?.length ?? 0) > 0
  const preferNativeSimpleImport = args.options.preferNativeSimpleImport === true
  const shouldTryLargeSimpleImport =
    !hasExternalWorkbookCompanions &&
    (nativeOnly ||
      !hasCalcChain ||
      args.sourceByteLength >= largeCalcChainStreamingByteThreshold ||
      hasLargeCalcChainFormulaSet ||
      bypassLargeSimpleByteThreshold)

  return {
    bypassLargeSimpleByteThreshold,
    hasExternalWorkbookCompanions,
    inspectionOptions:
      hasMaterializationLimits || bypassLargeSimpleByteThreshold || preferNativeSimpleImport || nativeOnly ? { minByteLength: 0 } : undefined,
    shouldInspectBeforeLargeSimpleRouting: hasMaterializationLimits || needsCalcChainFormulaCountInspection,
    shouldInspectBeforeSheetJsFallback: args.options.limits !== false && args.sourceByteLength >= denseSheetJsByteThreshold,
    shouldTryLargeSimpleImport,
    shouldRetryDataOnlyLargeSimpleImport: shouldRetryDataOnlyLargeSimpleImport(
      args.inspection,
      args.sourceByteLength,
      allowCachedUnsupportedFormulaText,
    ),
    createLargeSimpleImportOptions: (overrides = {}) => ({
      ...(hasMaterializationLimits || bypassLargeSimpleByteThreshold || preferNativeSimpleImport || nativeOnly ? { minByteLength: 0 } : {}),
      allowUnsupportedFormulaText: allowCachedUnsupportedFormulaText,
      allowUnsupportedCellMetadata: allowCachedUnsupportedFormulaText,
      releaseArenaAfterMaterialization: true,
      releaseZipSource: true,
      maxMaterializedLazyPackageArtifactBytes: largeSimpleLazyPackageArtifactMaterializationLimit,
      ...overrides,
    }),
  }
}

function resolveXlsxSheetJsFallbackLimits(options: XlsxImportOptions): Required<XlsxImportLimits> | null {
  if (options.limits === false) {
    return null
  }
  return resolveXlsxImportLimits(options) ?? defaultSheetJsFallbackImportLimits
}

export function shouldRetryDataOnlyLargeSimpleImport(
  inspection: LargeSimpleXlsxHeadlessInspectResult | null,
  sourceByteLength: number,
  allowCachedUnsupportedFormulaText: boolean,
): boolean {
  if (sourceByteLength >= denseSheetJsByteThreshold) {
    return true
  }
  const stats = inspection?.stats
  return (
    allowCachedUnsupportedFormulaText &&
    stats !== undefined &&
    (stats.cellCount >= denseSheetJsByteThreshold || stats.formulaCellCount >= largeCalcChainStreamingFormulaThreshold)
  )
}
