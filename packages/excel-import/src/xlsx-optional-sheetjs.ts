import { createRequire } from 'node:module'
import type { SheetJsModule } from './xlsx-sheetjs-types.js'

const requireModule = createRequire(import.meta.url)

let loadedSheetJs: SheetJsModule | undefined

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isModuleNotFoundError(error: unknown): boolean {
  return isObjectRecord(error) && error['code'] === 'MODULE_NOT_FOUND'
}

function isOptionalSheetJsModule(value: unknown): value is SheetJsModule {
  if (!isObjectRecord(value) || typeof value['read'] !== 'function' || typeof value['write'] !== 'function') {
    return false
  }
  const utils = value['utils']
  return isObjectRecord(utils) && typeof utils['book_new'] === 'function' && typeof utils['book_append_sheet'] === 'function'
}

export function loadOptionalSheetJs(): SheetJsModule {
  if (loadedSheetJs) {
    return loadedSheetJs
  }
  try {
    const loadedModule: unknown = requireModule('xlsx')
    if (!isOptionalSheetJsModule(loadedModule)) {
      throw new TypeError('SheetJS xlsx module is missing required read, write, or workbook utility exports.')
    }
    loadedSheetJs = loadedModule
    return loadedSheetJs
  } catch (error) {
    if (!isModuleNotFoundError(error)) {
      throw error
    }
    throw new Error(
      'SheetJS xlsx is not a production dependency of @bilig/excel-import. Install xlsx only for legacy SheetJS fallback import/export paths, or use the native @bilig/xlsx source-preserving XLSX path.',
      { cause: error },
    )
  }
}
