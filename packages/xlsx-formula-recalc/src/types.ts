import type { RawCellContent, WorkPaperChange, WorkPaperConfig } from '@bilig/headless'
import type { ImportedWorkbookDiagnostics, XlsxExternalWorkbookInput } from '@bilig/headless/xlsx'
import type { CellValue } from '@bilig/protocol'
import type { XlsxFormulaRecalcEngine, XlsxFormulaRecalcFallbackPolicy, XlsxFormulaRecalcNativeDiagnostics } from '@bilig/xlsx'

export type { XlsxFormulaRecalcEngine, XlsxFormulaRecalcFallbackPolicy, XlsxFormulaRecalcNativeDiagnostics } from '@bilig/xlsx'

export type XlsxFormulaRecalcCellValue = CellValue

export interface XlsxFormulaRecalcEdit {
  readonly target: string
  readonly value: RawCellContent
}

export interface XlsxFormulaRecalcOptions {
  readonly fileName?: string
  readonly externalWorkbooks?: readonly XlsxExternalWorkbookInput[]
  readonly edits?: readonly XlsxFormulaRecalcEdit[]
  readonly reads?: readonly string[]
  readonly config?: WorkPaperConfig
  readonly engine?: XlsxFormulaRecalcEngine
  readonly maxRssBytes?: number
  readonly fallbackPolicy?: XlsxFormulaRecalcFallbackPolicy
}

export type XlsxFormulaRecalcDiagnostics = ImportedWorkbookDiagnostics & Partial<XlsxFormulaRecalcNativeDiagnostics>

export interface XlsxFormulaRecalcResult {
  readonly xlsx: Uint8Array
  readonly warnings: readonly string[]
  readonly sheetNames: readonly string[]
  readonly reads: Readonly<Record<string, XlsxFormulaRecalcCellValue>>
  readonly changes: readonly WorkPaperChange[]
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
  readonly changes: readonly WorkPaperChange[]
  readonly diagnostics?: XlsxFormulaRecalcDiagnostics
}
