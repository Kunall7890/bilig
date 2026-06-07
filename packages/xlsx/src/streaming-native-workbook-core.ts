import { createFileXlsxSourceReader, type FileXlsxSourceReader } from './file-source.js'
import { workbookSheetPathEntriesForSource, type WorkbookSheetPathEntry } from './workbook-sheet-paths.js'
import {
  readXlsxZipEntriesLazyFromByteSource,
  readXlsxZipEntryMetadata,
  type XlsxZipEntries,
  type XlsxZipEntryMetadata,
} from './zip-reader.js'

export type StreamingNativeWorkbookOpenPhase = 'open-source' | 'zip-central-directory' | 'workbook-metadata'
export type StreamingNativeWorkbookOpenErrorReason = 'invalid-or-zip64-xlsx'

export interface StreamingNativeWorkbookOpenPhaseInfo {
  readonly inputBytes: number
}

export interface StreamingNativeWorkbookCore {
  readonly inputPath: string
  readonly inputBytes: number
  readonly source: FileXlsxSourceReader
  readonly zip: XlsxZipEntries
  readonly sheetEntries: readonly WorkbookSheetPathEntry[]
  readonly sheetNames: readonly string[]
  readonly entryMetadata: readonly XlsxZipEntryMetadata[]
}

export class StreamingNativeWorkbookOpenError extends Error {
  readonly reason: StreamingNativeWorkbookOpenErrorReason
  readonly inputBytes: number

  constructor(message: string, reason: StreamingNativeWorkbookOpenErrorReason, inputBytes: number) {
    super(message)
    this.name = 'StreamingNativeWorkbookOpenError'
    this.reason = reason
    this.inputBytes = inputBytes
  }
}

export function openStreamingNativeWorkbookCore(
  inputPath: string,
  options: {
    readonly onPhase?: (phase: StreamingNativeWorkbookOpenPhase, info: StreamingNativeWorkbookOpenPhaseInfo) => void
  } = {},
): StreamingNativeWorkbookCore {
  const source = createFileXlsxSourceReader(inputPath)
  const inputBytes = source.byteLength
  try {
    const phaseInfo = { inputBytes }
    options.onPhase?.('open-source', phaseInfo)
    const entryMetadata = readXlsxZipEntryMetadata(source)
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      throw new StreamingNativeWorkbookOpenError(
        'XLSX native streaming core requires a ZIP central directory it can read lazily',
        'invalid-or-zip64-xlsx',
        inputBytes,
      )
    }
    options.onPhase?.('zip-central-directory', phaseInfo)
    const sheetEntries = workbookSheetPathEntriesForSource(zip)
    options.onPhase?.('workbook-metadata', phaseInfo)
    return {
      inputPath,
      inputBytes,
      source,
      zip,
      sheetEntries,
      sheetNames: sheetEntries.map((sheet) => sheet.name),
      entryMetadata: entryMetadata ?? [],
    }
  } catch (error) {
    source.release?.()
    throw error
  }
}

export function closeStreamingNativeWorkbookCore(core: StreamingNativeWorkbookCore): void {
  core.source.release?.()
}

export function withStreamingNativeWorkbookCore<T>(
  inputPath: string,
  options: {
    readonly onPhase?: (phase: StreamingNativeWorkbookOpenPhase, info: StreamingNativeWorkbookOpenPhaseInfo) => void
  },
  callback: (core: StreamingNativeWorkbookCore) => T,
): T {
  const core = openStreamingNativeWorkbookCore(inputPath, options)
  try {
    return callback(core)
  } finally {
    closeStreamingNativeWorkbookCore(core)
  }
}
