import type { WorkbookSnapshot } from '@bilig/protocol'
import { createRequire } from 'node:module'

interface ImportedWorkbook {
  readonly snapshot: WorkbookSnapshot
}

interface HeadlessXlsxModule {
  readonly exportXlsx: (snapshot: WorkbookSnapshot) => Uint8Array
  readonly importXlsx: (bytes: Uint8Array | ArrayBuffer, fileName: string, options?: unknown) => ImportedWorkbook
}

type HeadlessExportXlsx = HeadlessXlsxModule['exportXlsx']
type HeadlessImportXlsx = HeadlessXlsxModule['importXlsx']

const requireModule = createRequire(import.meta.url)
let loadedModule: HeadlessXlsxModule | undefined

function loadHeadlessXlsxModule(): HeadlessXlsxModule {
  loadedModule ??= readHeadlessXlsxModule(
    tryRequire('@bilig/excel-import') ?? tryRequire('../../excel-import/src/index.ts') ?? tryRequire('../../excel-import/dist/index.js'),
  )
  return loadedModule
}

function tryRequire(path: string): unknown {
  try {
    return requireModule(path)
  } catch (error) {
    if (isRecord(error) && error['code'] === 'MODULE_NOT_FOUND') {
      return undefined
    }
    throw error
  }
}

function readHeadlessXlsxModule(value: unknown): HeadlessXlsxModule {
  const loadedExportXlsx = isRecord(value) ? value['exportXlsx'] : undefined
  const loadedImportXlsx = isRecord(value) ? value['importXlsx'] : undefined
  if (isHeadlessExportXlsx(loadedExportXlsx) && isHeadlessImportXlsx(loadedImportXlsx)) {
    return {
      exportXlsx: loadedExportXlsx,
      importXlsx: loadedImportXlsx,
    }
  }
  throw new Error('Bilig headless XLSX subpath is missing importXlsx/exportXlsx')
}

function isHeadlessExportXlsx(value: unknown): value is HeadlessExportXlsx {
  return typeof value === 'function'
}

function isHeadlessImportXlsx(value: unknown): value is HeadlessImportXlsx {
  return typeof value === 'function'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array {
  return loadHeadlessXlsxModule().exportXlsx(snapshot)
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options?: unknown): ImportedWorkbook {
  return loadHeadlessXlsxModule().importXlsx(bytes, fileName, options)
}
