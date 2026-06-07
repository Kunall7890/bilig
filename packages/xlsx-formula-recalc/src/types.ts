import type { CellValue, LiteralInput } from '@bilig/protocol'
import type { XlsxExternalWorkbookInput, XlsxExternalWorkbookHydrationDiagnostics, XlsxFormulaRecalcNativeDiagnostics } from '@bilig/xlsx'

export type { XlsxExternalWorkbookInput, XlsxExternalWorkbookHydrationDiagnostics, XlsxFormulaRecalcNativeDiagnostics } from '@bilig/xlsx'

export type XlsxFormulaRecalcEngineMode = 'streaming-native'
export type XlsxFormulaRecalcEngine = 'auto' | XlsxFormulaRecalcEngineMode
export type XlsxFormulaRecalcFallbackPolicy = 'error'

export type XlsxFormulaRecalcCellValue = CellValue
export type XlsxFormulaRecalcChange = unknown

export interface XlsxFormulaRecalcImportedDiagnostics {
  readonly externalWorkbookHydration?: XlsxExternalWorkbookHydrationDiagnostics
}

export interface XlsxFormulaRecalcEdit {
  readonly target: string
  readonly value: LiteralInput
}

export interface XlsxFormulaRecalcOptions {
  readonly fileName?: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly reads?: readonly string[]
  readonly engine?: XlsxFormulaRecalcEngine
  readonly maxRssBytes?: number
  readonly fallbackPolicy?: XlsxFormulaRecalcFallbackPolicy
}

export type XlsxFormulaRecalcDiagnostics = XlsxFormulaRecalcImportedDiagnostics & Partial<XlsxFormulaRecalcNativeDiagnostics>

export interface XlsxFormulaRecalcResult {
  readonly xlsx: Uint8Array
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly XlsxFormulaRecalcChange[]
  readonly diagnostics?: XlsxFormulaRecalcDiagnostics
}

export interface XlsxFormulaRecalcFileOptions extends XlsxFormulaRecalcOptions {
  readonly outputPath: string
}

export interface XlsxFormulaRecalcFileResult {
  readonly bytesWritten: number
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly XlsxFormulaRecalcChange[]
  readonly diagnostics?: XlsxFormulaRecalcDiagnostics
}
