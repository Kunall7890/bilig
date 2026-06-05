export interface SheetJsCellAddress {
  r: number
  c: number
}

export interface SheetJsRange {
  s: SheetJsCellAddress
  e: SheetJsCellAddress
}

export type SheetJsExcelDataType = 'b' | 'n' | 'e' | 's' | 'd' | 'z'

export interface SheetJsColInfo {
  wpx?: number
  wch?: number
  hidden?: boolean
  [key: string]: unknown
}

export interface SheetJsRowInfo {
  hpx?: number
  hpt?: number
  hidden?: boolean
  [key: string]: unknown
}

export interface SheetJsComment {
  t: string
  a?: string
  [key: string]: unknown
}

export interface SheetJsCellObject {
  t: SheetJsExcelDataType
  v?: unknown
  f?: string
  F?: string
  w?: string
  z?: string | number
  l?: Record<string, unknown>
  c?: SheetJsComment[]
  s?: unknown
  [key: string]: unknown
}

export interface SheetJsWorkSheet {
  '!ref'?: string
  '!cols'?: SheetJsColInfo[]
  '!rows'?: SheetJsRowInfo[]
  '!merges'?: SheetJsRange[]
  '!autofilter'?: unknown
  '!data'?: unknown[]
  [cell: string]: unknown
}

export interface SheetJsDefinedName {
  Name?: string
  Ref?: string
  Sheet?: number
  Comment?: string
  [key: string]: unknown
}

export interface SheetJsWorkbookSheet {
  name?: string
  Hidden?: number
  CodeName?: string
  [key: string]: unknown
}

export interface SheetJsWorkbookMetadata {
  Names?: SheetJsDefinedName[]
  Sheets?: SheetJsWorkbookSheet[]
  Views?: unknown[]
  WBProps?: Record<string, unknown>
  [key: string]: unknown
}

export interface SheetJsWorkBook {
  SheetNames: string[]
  Sheets: Record<string, SheetJsWorkSheet | undefined>
  Workbook?: SheetJsWorkbookMetadata
  vbaraw?: unknown
  [key: string]: unknown
}

export interface SheetJsModule {
  read(data: unknown, options?: Record<string, unknown>): SheetJsWorkBook
  write(workbook: SheetJsWorkBook, options: Record<string, unknown>): unknown
  utils: {
    book_new(): SheetJsWorkBook
    book_append_sheet(workbook: SheetJsWorkBook, worksheet: SheetJsWorkSheet, name?: string, roll?: boolean): void
  }
}
