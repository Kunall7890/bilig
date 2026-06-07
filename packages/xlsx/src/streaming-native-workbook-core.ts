import { createFileXlsxSourceReader, type FileXlsxSourceReader } from './file-source.js'
import { workbookSheetPathEntriesForSource, type WorkbookSheetPathEntry } from './workbook-sheet-paths.js'
import {
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
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

export interface StreamingNativeWorkbookCellStats {
  readonly sheetNames: readonly string[]
  readonly nonEmptyCellCount: number
}

export interface StreamingNativeWorkbookPackagePartStats {
  readonly definedNameCount: number
  readonly tableCount: number
  readonly pivotTableCount: number
  readonly chartCount: number
  readonly macroModuleCount: number
  readonly macroByteLength: number
  readonly externalLinkCount: number
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

const textDecoder = new TextDecoder()
const worksheetCellStartTagPattern = /<(?:[A-Za-z_][\w.-]*:)?c\b/gu
const worksheetCellStartCarryLength = 128

export function scanStreamingNativeWorkbookCellStats(
  core: Pick<StreamingNativeWorkbookCore, 'zip' | 'sheetEntries'>,
): StreamingNativeWorkbookCellStats {
  let nonEmptyCellCount = 0
  for (const sheet of core.sheetEntries) {
    nonEmptyCellCount += countStreamingNativeWorksheetCellElements(core.zip, sheet.path)
  }
  return {
    sheetNames: core.sheetEntries.map((sheet) => sheet.name),
    nonEmptyCellCount,
  }
}

export function scanStreamingNativeWorkbookPackageParts(input: {
  readonly zip: XlsxZipEntries
  readonly entryMetadata?: readonly XlsxZipEntryMetadata[]
}): StreamingNativeWorkbookPackagePartStats {
  const entryMetadata = input.entryMetadata ?? []
  const paths = entryMetadata.length > 0 ? entryMetadata.map((entry) => entry.path) : Object.keys(input.zip)
  const metadataByPath = new Map(entryMetadata.map((entry) => [entry.path, entry]))
  const workbookXml = getZipText(input.zip, 'xl/workbook.xml') ?? ''
  const externalReferenceCount = workbookXml.match(/<(?:[A-Za-z_][\w.-]*:)?externalReference\b/gu)?.length ?? 0
  const externalLinkPartCount = paths.filter((path) => /^xl\/externalLinks\/externalLink[0-9]+\.xml$/u.test(path)).length
  const macroPaths = paths.filter((path) => /(?:^|\/)vbaProject\.bin$/iu.test(path))
  return {
    definedNameCount: workbookXml.match(/<(?:[A-Za-z_][\w.-]*:)?definedName\b/gu)?.length ?? 0,
    tableCount: paths.filter((path) => /^xl\/tables\/[^/]+\.xml$/u.test(path)).length,
    pivotTableCount: paths.filter((path) => /^xl\/pivotTables\/[^/]+\.xml$/u.test(path)).length,
    chartCount: paths.filter((path) => /^xl\/charts\/[^/]+\.xml$/u.test(path) || /^xl\/chartsheets\/[^/]+\.xml$/u.test(path)).length,
    macroModuleCount: macroPaths.length,
    macroByteLength: macroPaths.reduce(
      (sum, path) => sum + (metadataByPath.get(path)?.uncompressedSize ?? input.zip[path]?.byteLength ?? 0),
      0,
    ),
    externalLinkCount: Math.max(externalReferenceCount, externalLinkPartCount),
  }
}

function countStreamingNativeWorksheetCellElements(zip: XlsxZipEntries, sheetPath: string): number {
  let count = 0
  let buffer = ''
  const processBuffer = (final: boolean): void => {
    const safeEnd = final ? buffer.length : Math.max(0, buffer.length - worksheetCellStartCarryLength)
    if (safeEnd === 0 && !final) {
      return
    }
    for (const _match of buffer.slice(0, safeEnd).matchAll(worksheetCellStartTagPattern)) {
      count += 1
    }
    buffer = buffer.slice(safeEnd)
  }
  const streamed = forEachInflatedXlsxZipEntryChunk(
    zip,
    sheetPath,
    (chunk) => {
      buffer += textDecoder.decode(chunk, { stream: true })
      processBuffer(false)
    },
    { chunkSize: 64 * 1024, forceStreamingInflate: true },
  )
  if (!streamed) {
    return 0
  }
  buffer += textDecoder.decode()
  processBuffer(true)
  return count
}
