import { XLSM_CONTENT_TYPE, XLSX_CONTENT_TYPE, type ExcelWorkbookImportContentType } from './workbook-import-content-types.js'
import type { ImportedXlsxSourceReference } from './xlsx-source-bytes.js'

function isRangeReaderSource(source: ImportedXlsxSourceReference): boolean {
  return !(source instanceof Uint8Array) && typeof Reflect.get(source, 'readRange') === 'function'
}

export function canAttachSourceForUntouchedSheetJsExport(args: {
  readonly source: ImportedXlsxSourceReference
  readonly contentType: ExcelWorkbookImportContentType
  readonly shouldUseCachedFormulaOpenMode: boolean
  readonly hasExternalWorkbookCompanions: boolean
  readonly hasMaterializedExternalCacheSheets: boolean
  readonly hasTables: boolean
  readonly hasImportedFilters: boolean
  readonly hasSlicerConnectionArtifacts: boolean
  readonly hasConditionalFormats: boolean
  readonly hasHyperlinks: boolean
  readonly hasComments: boolean
  readonly hasMacroPayload: boolean
}): boolean {
  if (args.contentType !== XLSX_CONTENT_TYPE && args.contentType !== XLSM_CONTENT_TYPE) {
    return false
  }
  if (args.shouldUseCachedFormulaOpenMode) {
    return true
  }
  if (args.contentType !== XLSX_CONTENT_TYPE || !isRangeReaderSource(args.source)) {
    return false
  }
  return (
    !args.hasExternalWorkbookCompanions &&
    !args.hasMaterializedExternalCacheSheets &&
    !args.hasTables &&
    !args.hasImportedFilters &&
    !args.hasSlicerConnectionArtifacts &&
    !args.hasConditionalFormats &&
    !args.hasHyperlinks &&
    !args.hasComments &&
    !args.hasMacroPayload
  )
}
