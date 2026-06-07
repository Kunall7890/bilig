export interface XlsxExternalWorkbookInput {
  readonly bytes: Uint8Array | ArrayBuffer
  readonly fileName?: string
  readonly workbookName?: string
  readonly target?: string
}

export const xlsxExternalWorkbookByteInputLimit = 1_000_000

export function assertXlsxExternalWorkbookByteInputWithinLimit(byteLength: number, fileName: string): void {
  if (byteLength <= xlsxExternalWorkbookByteInputLimit) {
    return
  }
  throw new Error(
    [
      `external workbook byte input is small-workbook only: ${fileName} is ${byteLength} bytes`,
      `limit is ${xlsxExternalWorkbookByteInputLimit} bytes`,
      'Large external workbook hydration must use a native file-backed path before it can be enabled for production large-XLSX jobs.',
    ].join('; '),
  )
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
  readonly preferNativeSimpleImport?: boolean
}
