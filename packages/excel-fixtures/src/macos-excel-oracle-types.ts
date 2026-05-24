import type { NormalizedFormulaValue } from './oracle-harness.js'

export interface MacosExcelOracleFormulaCell {
  readonly address: string
  readonly formula: string
}

export type MacosExcelLinkUpdateMode = 'all' | 'external' | 'never' | 'remote'
export type MacosExcelSortHeader = 'guess' | 'no' | 'yes'
export type MacosExcelSortOrder = 'ascending' | 'descending'
export type MacosExcelSortOrientation = 'columns' | 'rows'
export type MacosExcelAutoFilterOperator = 'autofilter and' | 'autofilter or' | 'filter by value'

export interface MacosExcelSortKey {
  readonly key: string
  readonly order?: MacosExcelSortOrder
}

export interface MacosExcelRecalculationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly formulaCells: readonly MacosExcelOracleFormulaCell[]
  readonly valueCells: readonly string[]
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
}

export interface MacosExcelRecalculationOracleResult {
  readonly excelVersion: string
  readonly rawValues: readonly string[]
  readonly values: readonly NormalizedFormulaValue[]
}

export interface MacosExcelInspectionOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly formulaCells: readonly MacosExcelOracleFormulaCell[]
  readonly inspectCells: readonly string[]
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
  readonly refreshWorkbook?: boolean
}

export interface MacosExcelPackageOpenSaveOracleRequest {
  readonly workbookPath: string
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
  readonly refreshWorkbook?: boolean
}

export type MacosExcelStructuralOperation =
  | { readonly kind: 'insertRows'; readonly range: string }
  | { readonly kind: 'insertColumns'; readonly range: string }
  | { readonly kind: 'deleteRows'; readonly range: string }
  | { readonly kind: 'deleteColumns'; readonly range: string }
  | { readonly kind: 'setCellValue'; readonly address: string; readonly value: string | number | boolean }
  | { readonly kind: 'clearCell'; readonly address: string }
  | { readonly kind: 'createSheet'; readonly name: string }
  | { readonly kind: 'renameSheet'; readonly newName: string }
  | { readonly kind: 'deleteSheet'; readonly name: string }
  | { readonly kind: 'moveSheet'; readonly name: string; readonly before?: string; readonly after?: string }
  | { readonly kind: 'moveRows'; readonly sourceRange: string; readonly destinationRange: string }
  | { readonly kind: 'moveColumns'; readonly sourceRange: string; readonly destinationRange: string }
  | { readonly kind: 'createDataTable'; readonly range: string; readonly rowInput?: string; readonly columnInput?: string }
  | {
      readonly kind: 'applySort'
      readonly range: string
      readonly keys: readonly MacosExcelSortKey[]
      readonly header?: MacosExcelSortHeader
      readonly orientation?: MacosExcelSortOrientation
    }
  | {
      readonly kind: 'applyTableSort'
      readonly tableName: string
      readonly keys: readonly MacosExcelSortKey[]
      readonly header?: MacosExcelSortHeader
      readonly orientation?: MacosExcelSortOrientation
    }
  | {
      readonly kind: 'applyTableAutoFilter'
      readonly tableName: string
      readonly field: number
      readonly criteria1?: string | number | boolean
      readonly operator?: MacosExcelAutoFilterOperator
      readonly criteria2?: string | number | boolean
      readonly visibleDropDown?: boolean
    }

export interface MacosExcelStructuralOperationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly operations: readonly MacosExcelStructuralOperation[]
  readonly inspectCells: readonly string[]
  readonly companionWorkbookPaths?: readonly string[]
  readonly formulaCells?: readonly MacosExcelOracleFormulaCell[]
  readonly appPath?: string
  readonly saveWorkbook?: boolean
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
}

export interface MacosExcelRejectedStructuralOperationOracleRequest {
  readonly workbookPath: string
  readonly worksheetName: string
  readonly operation: MacosExcelStructuralOperation
  readonly companionWorkbookPaths?: readonly string[]
  readonly appPath?: string
  readonly timeoutMs?: number
  readonly updateLinks?: MacosExcelLinkUpdateMode
}

export interface MacosExcelCellInspection {
  readonly address: string
  readonly formula?: string
  readonly rawValue: string
  readonly value: NormalizedFormulaValue
}

export interface MacosExcelInspectionOracleResult {
  readonly cells: readonly MacosExcelCellInspection[]
  readonly excelVersion: string
}

export interface MacosExcelPackageOpenSaveOracleResult {
  readonly excelVersion: string
}

export interface MacosExcelRejectedStructuralOperationOracleResult {
  readonly excelVersion: string
  readonly errorMessage: string
  readonly errorNumber: number
  readonly sheetNames: readonly string[]
}
