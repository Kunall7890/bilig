#!/usr/bin/env bun

import { readFileSync, statSync } from 'node:fs'
import { createRequire } from 'node:module'
import { basename, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import type { ImportedWorkbook } from '../packages/excel-import/src/workbook-import-result.js'
import type { XlsxByteSourceImportOptions } from '../packages/excel-import/src/xlsx-byte-source-import.js'
import type {
  tryInspectLargeSimpleXlsxHeadless,
  tryInspectLargeSimpleXlsxHeadlessAsync,
} from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import type { XlsxZipEntries } from '../packages/excel-import/src/xlsx-zip.js'
import { readFlagArg, readStringArg } from './public-workbook-corpus-cli.ts'
import { FileBackedXlsxZipByteSource } from './public-workbook-corpus-xlsx-byte-source.ts'

interface PublicXlsxImportModule {
  readonly importXlsx: (bytes: Uint8Array, fileName: string) => ImportedWorkbook
}

interface XlsxByteSourceImportModule {
  readonly importXlsxFromZipByteSource: (
    source: FileBackedXlsxZipByteSource,
    fileName: string,
    options?: XlsxByteSourceImportOptions,
  ) => ImportedWorkbook
}

type TryInspectLargeSimpleXlsxHeadless = typeof tryInspectLargeSimpleXlsxHeadless
type TryInspectLargeSimpleXlsxHeadlessAsync = typeof tryInspectLargeSimpleXlsxHeadlessAsync

interface LargeSimpleInspectModule {
  readonly tryInspectLargeSimpleXlsxHeadless: TryInspectLargeSimpleXlsxHeadless
  readonly tryInspectLargeSimpleXlsxHeadlessAsync?: TryInspectLargeSimpleXlsxHeadlessAsync
}

interface XlsxZipModule {
  readonly readXlsxZipEntriesLazyFromByteSource: (source: FileBackedXlsxZipByteSource) => XlsxZipEntries | null
}

export interface ExternalXlsxStressWorkerSummary {
  readonly importMode: 'headless-inspect' | 'public-snapshot'
  readonly sheets: number
  readonly cells: number
  readonly formulas: number
  readonly warnings: number
  readonly workbookMetadataKeys: readonly string[]
  readonly sheetMetadataKeys: readonly string[]
}

const requireModule = createRequire(import.meta.url)
const headlessInspectCellThreshold = 1_000_000
const headlessInspectFormulaFreeCellThreshold = 10_000
const headlessInspectFormulaFreeByteThreshold = 100 * 1024 * 1024
const headlessInspectMetadataRichKeyThreshold = 4
const headlessInspectMetadataRichKeys = new Set([
  'cellMetadataRefs',
  'drawingArtifacts',
  'hyperlinks',
  'pivotArtifacts',
  'printerSettings',
  'richTextArtifacts',
  'styleRanges',
])
let xlsxByteSourceImportModule: XlsxByteSourceImportModule | undefined
let largeSimpleInspectModule: LargeSimpleInspectModule | undefined
let xlsxZipModule: XlsxZipModule | undefined

export async function runExternalXlsxStressWorker(): Promise<void> {
  const filePath = resolve(readStringArg('--file', ''))
  const fileName = readStringArg('--file-name', basename(filePath))
  const usePublicImport = readFlagArg('--public-import')
  const summary = usePublicImport
    ? summarizeExternalXlsxImportedWorkbook(importPublicXlsx(readFileSync(filePath), fileName))
    : await summarizeFileBackedXlsx(filePath, fileName)
  collectGarbage()
  process.stdout.write(`${JSON.stringify(summary)}\n`)
}

export function summarizeExternalXlsxImportedWorkbook(imported: ImportedWorkbook): ExternalXlsxStressWorkerSummary {
  let cells = imported.stats?.cellCount ?? 0
  let formulas = imported.stats?.formulaCellCount ?? 0
  const sheetMetadataKeys = new Set<string>()
  for (const sheet of imported.snapshot.sheets) {
    if (!imported.stats) {
      cells += sheet.cells.length
      for (const cell of sheet.cells) {
        if ('formula' in cell) {
          formulas += 1
        }
      }
    }
    for (const key of Object.keys(sheet.metadata ?? {})) {
      sheetMetadataKeys.add(key)
    }
  }
  return {
    importMode: 'public-snapshot',
    sheets: imported.snapshot.sheets.length,
    cells,
    formulas,
    warnings: imported.warnings.length,
    workbookMetadataKeys: Object.keys(imported.snapshot.workbook.metadata ?? {}).toSorted(),
    sheetMetadataKeys: [...sheetMetadataKeys].toSorted(),
  }
}

async function summarizeFileBackedXlsx(filePath: string, fileName: string): Promise<ExternalXlsxStressWorkerSummary> {
  const headless = await inspectFileBackedXlsxHeadless(filePath, fileName)
  if (headless && shouldSummarizeFileBackedHeadlessInspect(headless, statSync(filePath).size)) {
    return {
      importMode: 'headless-inspect',
      sheets: headless.sheetNames.length,
      cells: headless.stats.cellCount,
      formulas: headless.stats.formulaCellCount,
      warnings: headless.warnings.length,
      workbookMetadataKeys: headless.workbookMetadataKeys,
      sheetMetadataKeys: headless.sheetMetadataKeys,
    }
  }
  return summarizeExternalXlsxImportedWorkbook(importFileBackedXlsx(filePath, fileName))
}

export function shouldSummarizeFileBackedHeadlessInspect(
  inspected: ReturnType<TryInspectLargeSimpleXlsxHeadless>,
  sourceByteLength = 0,
): boolean {
  return (
    inspected !== null &&
    inspected.stats.cellCount > 0 &&
    (inspected.stats.cellCount >= headlessInspectCellThreshold ||
      (inspected.stats.formulaCellCount === 0 &&
        (inspected.stats.cellCount >= headlessInspectFormulaFreeCellThreshold ||
          sourceByteLength >= headlessInspectFormulaFreeByteThreshold ||
          hasMetadataRichHeadlessVisitorEvidence(inspected))))
  )
}

function hasMetadataRichHeadlessVisitorEvidence(inspected: NonNullable<ReturnType<TryInspectLargeSimpleXlsxHeadless>>): boolean {
  if (inspected.sheetMetadataKeys.includes('conditionalFormats')) {
    return false
  }
  let coveredKeyCount = 0
  for (const key of inspected.sheetMetadataKeys) {
    if (headlessInspectMetadataRichKeys.has(key)) {
      coveredKeyCount += 1
    }
  }
  return coveredKeyCount >= headlessInspectMetadataRichKeyThreshold
}

async function inspectFileBackedXlsxHeadless(filePath: string, fileName: string): Promise<ReturnType<TryInspectLargeSimpleXlsxHeadless>> {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    const zip = loadXlsxZipModule().readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      return null
    }
    const inspectModule = loadLargeSimpleInspectModule()
    const inspect = inspectModule.tryInspectLargeSimpleXlsxHeadlessAsync ?? inspectModule.tryInspectLargeSimpleXlsxHeadless
    return await inspect({ byteLength: source.byteLength }, fileName, zip, {
      allowUnsupportedWorksheetFeaturesForMetrics: true,
      afterWorksheetChunkBatch: collectGarbage,
      afterWorksheetScan: collectGarbage,
      minByteLength: 0,
      releaseZipSource: true,
    })
  } finally {
    source.release()
  }
}

function importPublicXlsx(bytes: Uint8Array, fileName: string): ImportedWorkbook {
  const importer = readPublicXlsxImportModule(requireModule('../packages/excel-import/src/index.js'))
  return importer.importXlsx(bytes, fileName)
}

function importFileBackedXlsx(filePath: string, fileName: string): ImportedWorkbook {
  const source = new FileBackedXlsxZipByteSource(filePath)
  try {
    return loadXlsxByteSourceImportModule().importXlsxFromZipByteSource(source, fileName, {
      attachSourceReaderForUntouchedExport: false,
    })
  } finally {
    source.release()
  }
}

function loadXlsxByteSourceImportModule(): XlsxByteSourceImportModule {
  xlsxByteSourceImportModule ??= readXlsxByteSourceImportModule(requireModule('../packages/excel-import/src/xlsx-byte-source-import.js'))
  return xlsxByteSourceImportModule
}

function loadLargeSimpleInspectModule(): LargeSimpleInspectModule {
  largeSimpleInspectModule ??= readLargeSimpleInspectModule(
    requireModule('../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'),
  )
  return largeSimpleInspectModule
}

function loadXlsxZipModule(): XlsxZipModule {
  xlsxZipModule ??= readXlsxZipModule(requireModule('../packages/excel-import/src/xlsx-zip.js'))
  return xlsxZipModule
}

function readPublicXlsxImportModule(value: unknown): PublicXlsxImportModule {
  if (isRecord(value) && typeof value['importXlsx'] === 'function') {
    return { importXlsx: value['importXlsx'] }
  }
  throw new Error('Public XLSX importer module is missing required exports')
}

function readXlsxByteSourceImportModule(value: unknown): XlsxByteSourceImportModule {
  if (isRecord(value) && typeof value['importXlsxFromZipByteSource'] === 'function') {
    return { importXlsxFromZipByteSource: value['importXlsxFromZipByteSource'] }
  }
  throw new Error('XLSX byte-source importer module is missing required exports')
}

function readLargeSimpleInspectModule(value: unknown): LargeSimpleInspectModule {
  if (isRecord(value) && typeof value['tryInspectLargeSimpleXlsxHeadless'] === 'function') {
    return {
      tryInspectLargeSimpleXlsxHeadless: value['tryInspectLargeSimpleXlsxHeadless'],
      ...(typeof value['tryInspectLargeSimpleXlsxHeadlessAsync'] === 'function'
        ? { tryInspectLargeSimpleXlsxHeadlessAsync: value['tryInspectLargeSimpleXlsxHeadlessAsync'] }
        : {}),
    }
  }
  throw new Error('Large-simple XLSX inspect module is missing required exports')
}

function readXlsxZipModule(value: unknown): XlsxZipModule {
  if (isRecord(value) && typeof value['readXlsxZipEntriesLazyFromByteSource'] === 'function') {
    return { readXlsxZipEntriesLazyFromByteSource: value['readXlsxZipEntriesLazyFromByteSource'] }
  }
  throw new Error('XLSX zip module is missing required exports')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function collectGarbage(): void {
  if (typeof Bun !== 'undefined' && typeof Bun.gc === 'function') {
    Bun.gc(true)
    return
  }
  const gc = Reflect.get(globalThis, 'gc')
  if (typeof gc === 'function') {
    gc()
  }
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    await runExternalXlsxStressWorker()
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exit(1)
  }
}
