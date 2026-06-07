export * from '@bilig/headless/xlsx'
export {
  WorkPaper,
  inspectXlsxCache,
  inspectXlsxCacheFileStreamingNative,
  parseQualifiedA1,
  parseQualifiedCellTarget,
  recalculateSheetjsWorkbook,
  recalculateXlsx,
  recalculateXlsxToFile,
  StreamingNativeXlsxRecalcError,
  xlsxCacheDoctorSchemaVersion,
} from './xlsx-recalc.js'
export type {
  StreamingNativeFormulaCounts,
  StreamingNativeXlsxCacheFormulaInspection,
  StreamingNativeXlsxCacheInspectionLimit,
  StreamingNativeXlsxCacheInspectionResult,
  StreamingNativeXlsxCacheLiteral,
  StreamingNativeXlsxCacheStatus,
  StreamingNativeXlsxCacheStatusSummary,
  XlsxCacheFormulaInspection,
  XlsxCacheInspectionLimit,
  XlsxCacheInspectionOptions,
  XlsxCacheInspectionResult,
  XlsxCacheStatus,
  XlsxCacheStatusSummary,
  XlsxFormulaRecalcCellValue,
  XlsxFormulaRecalcChange,
  XlsxFormulaRecalcDiagnostics,
  XlsxFormulaRecalcEdit,
  XlsxFormulaRecalcFileOptions,
  XlsxFormulaRecalcFileResult,
  XlsxFormulaRecalcNativeDiagnostics,
  XlsxFormulaRecalcOptions,
  XlsxFormulaRecalcPhaseRss,
  XlsxFormulaRecalcResult,
  XlsxFormulaRecalcWorkPaperConfig,
  XlsxFormulaRecalcWorkPaperEngine,
  XlsxFormulaRecalcWorkPaperFallbackPolicy,
} from './xlsx-recalc.js'

import { importXlsx as importHeadlessXlsx, type XlsxImportOptions } from '@bilig/headless/xlsx'

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}) {
  return importHeadlessXlsx(bytes, fileName, {
    preferNativeSimpleImport: true,
    ...options,
  })
}
