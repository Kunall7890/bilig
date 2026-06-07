import { createRequire } from 'node:module'
import type { SheetJsModule, SheetJsWorkBook, SheetJsWorkSheet } from '../xlsx-sheetjs-types.js'

interface LegacySheetJsModule extends SheetJsModule {
  readonly utils: SheetJsModule['utils'] & {
    readonly aoa_to_sheet: (data: readonly (readonly unknown[])[]) => SheetJsWorkSheet
  }
}

const requireModule = createRequire(import.meta.url)

let loadedLegacySheetJs: LegacySheetJsModule | undefined

export type SheetJsFallbackWorkbook = SheetJsWorkBook

export function readSheetJsFallbackWorkbook(bytes: Uint8Array): SheetJsFallbackWorkbook {
  return loadLegacySheetJs().read(bytes, {
    type: 'array',
    bookFiles: true,
    cellFormula: true,
    cellNF: true,
    cellStyles: false,
  })
}

export function buildBinaryWorkbook(): Uint8Array {
  const xlsx = loadLegacySheetJs()
  const workbook = xlsx.utils.book_new()
  const sheet = xlsx.utils.aoa_to_sheet([
    ['Name', 'Value'],
    ['alpha', 12],
  ])
  xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet1')
  xlsx.utils.book_append_sheet(workbook, xlsx.utils.aoa_to_sheet([['notes']]), 'Sheet2')
  return toUint8Array(xlsx.write(workbook, { bookType: 'xlsb', type: 'buffer' }))
}

export function buildLegacyWorkbook(): Uint8Array {
  const xlsx = loadLegacySheetJs()
  const workbook = xlsx.utils.book_new()
  const sheet = xlsx.utils.aoa_to_sheet([
    ['Department', 'Amount'],
    ['Operations', 1250],
    ['Finance', 1800],
  ])
  sheet.C2 = { t: 'n', f: 'B2+B3', v: 3050 }
  sheet['!ref'] = 'A1:C3'
  xlsx.utils.book_append_sheet(workbook, sheet, 'Salary')
  return toUint8Array(xlsx.write(workbook, { bookType: 'xls', type: 'buffer' }))
}

export function buildNamespacedFormulaWorkbook(): Uint8Array {
  const xlsx = loadLegacySheetJs()
  const workbook = xlsx.utils.book_new()
  const sheet = xlsx.utils.aoa_to_sheet([[1], [2]])
  sheet.A3 = { t: 'n', f: 'msoxl:=SUM(A1:A2)', v: 3 }
  sheet.B3 = { t: 'n', f: 'of:=SUM(A1:A2)', v: 3 }
  sheet['!ref'] = 'A1:B3'
  xlsx.utils.book_append_sheet(workbook, sheet, 'Expenses')
  return toUint8Array(xlsx.write(workbook, { bookType: 'ods', type: 'buffer' }))
}

function loadLegacySheetJs(): LegacySheetJsModule {
  if (loadedLegacySheetJs) {
    return loadedLegacySheetJs
  }
  const loaded: unknown = requireModule('xlsx')
  if (!isLegacySheetJsModule(loaded)) {
    throw new TypeError('SheetJS xlsx module is missing required legacy workbook fixture helpers.')
  }
  loadedLegacySheetJs = loaded
  return loadedLegacySheetJs
}

function isLegacySheetJsModule(value: unknown): value is LegacySheetJsModule {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const utils = objectField(value, 'utils')
  return (
    typeof objectField(value, 'read') === 'function' &&
    typeof objectField(value, 'write') === 'function' &&
    typeof utils === 'object' &&
    utils !== null &&
    typeof objectField(utils, 'book_new') === 'function' &&
    typeof objectField(utils, 'book_append_sheet') === 'function' &&
    typeof objectField(utils, 'aoa_to_sheet') === 'function'
  )
}

function objectField(value: object, key: string): unknown {
  return Reflect.get(value, key)
}

function toUint8Array(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) {
    return value
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value)
  }
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
  }
  throw new TypeError('SheetJS fixture writer did not return bytes.')
}
