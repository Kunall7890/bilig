import { randomUUID } from 'node:crypto'
import { closeSync, openSync, renameSync, unlinkSync } from 'node:fs'
import { unzipSync } from 'fflate'
import { Deflate } from 'fflate-stream'

import {
  crc32Finalize,
  crc32Update,
  decodeCellAddress,
  decodeCellRange,
  encodeCellRange,
  exportXlsxSourceLiteralPatches as exportBiligXlsxSourceLiteralPatches,
  exportXlsxSourceLiteralPatchesToFile as exportBiligXlsxSourceLiteralPatchesToFile,
  exportXlsxSourceLiteralPatchesToFileAsync as exportBiligXlsxSourceLiteralPatchesToFileAsync,
  type FilePreparedZipEntry,
  type PreparedZipEntry,
  type PreparedZipEntrySizes,
  type XlsxCellRange,
  writeAllSync,
  zipSourcePreservingEntries,
  zipSourcePreservingEntriesToFile,
} from '@bilig/xlsx'
import type { LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { escapeXmlAttribute, getZipText, setXmlAttribute, setZipText } from './xlsx-export-xml.js'
import { readImportedXlsxSourceCellPatches, type ImportedXlsxSourceCellPatch, type ImportedXlsxSourceReader } from './xlsx-source-bytes.js'
import { readXmlAttribute, worksheetCellElementPattern, worksheetCellOpeningTagPattern } from './xlsx-style-xml.js'
import { workbookSheetPathEntriesFromSource } from './xlsx-workbook-sheet-paths.js'
import {
  forEachInflatedXlsxZipEntryChunk,
  forEachInflatedXlsxZipEntryChunkAsync,
  readXlsxZipEntriesLazy,
  readXlsxZipEntriesLazyFromByteSource,
  type XlsxZipByteSource,
} from './xlsx-zip.js'

interface WorksheetPatch {
  readonly literals: ReadonlyMap<string, LiteralInput>
}

type ImportedXlsxSourceReference = Uint8Array | ImportedXlsxSourceReader

type SourcePreservingZip = Record<string, Uint8Array>

export interface XlsxSourceLiteralPatch {
  readonly sheetName: string
  readonly address: string
  readonly value: LiteralInput
  readonly preserveFormula?: boolean
}

export interface XlsxSourceLiteralPatchExportInput {
  readonly source: ImportedXlsxSourceReference
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly sheetNames?: readonly string[]
  readonly workbookName?: string
}

export interface XlsxSourceLiteralPatchFileExportInput extends XlsxSourceLiteralPatchExportInput {
  readonly outputPath: string
}

export interface XlsxSourceLiteralPatchFileExportResult {
  readonly bytesWritten: number
}

function isXlsxZipByteSource(source: ImportedXlsxSourceReference): source is ImportedXlsxSourceReader & XlsxZipByteSource {
  return typeof (source as Partial<XlsxZipByteSource>).readRange === 'function'
}

function readSourcePreservingZip(source: ImportedXlsxSourceReference): SourcePreservingZip {
  if (source instanceof Uint8Array) {
    return readXlsxZipEntriesLazy(source)
  }
  if (isXlsxZipByteSource(source)) {
    const lazyZip = readXlsxZipEntriesLazyFromByteSource(source)
    if (lazyZip) {
      return lazyZip
    }
  }
  return unzipSync(source.readBytes())
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function removeXmlAttribute(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${name}=(["'])[\\s\\S]*?\\1`, 'u'), '')
}

function setCellType(openingTag: string, type: string | null): string {
  return type === null ? removeXmlAttribute(openingTag, 't') : setXmlAttribute(openingTag, 't', type)
}

function cellParts(cellXml: string): { readonly tagName: string; readonly openingTag: string; readonly body: string } | null {
  const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
  const tagName = openingTag ? /^<([^\s/>]+)/u.exec(openingTag)?.[1] : undefined
  if (!openingTag || !tagName) {
    return null
  }
  if (openingTag.endsWith('/>')) {
    return { tagName, openingTag, body: '' }
  }
  const closingTag = `</${tagName}>`
  return {
    tagName,
    openingTag,
    body: cellXml.slice(openingTag.length, cellXml.endsWith(closingTag) ? -closingTag.length : undefined),
  }
}

function inlineStringBody(value: string): string {
  const preserveSpace = /^\s|\s$/u.test(value) ? ' xml:space="preserve"' : ''
  return `<is><t${preserveSpace}>${escapeXmlText(value)}</t></is>`
}

function literalCellBody(value: LiteralInput): { readonly type: string | null; readonly body: string } | null {
  if (value === null) {
    return { type: null, body: '' }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { type: null, body: `<v>${String(value)}</v>` } : null
  }
  if (typeof value === 'boolean') {
    return { type: 'b', body: `<v>${value ? '1' : '0'}</v>` }
  }
  return { type: 'inlineStr', body: inlineStringBody(value) }
}

function rewriteCellXml(cellXml: string, type: string | null, body: string): string | null {
  const parts = cellParts(cellXml)
  if (!parts) {
    return null
  }
  const nextOpeningTag = setCellType(parts.openingTag, type)
  return `${nextOpeningTag.endsWith('/>') ? nextOpeningTag.slice(0, -2) + '>' : nextOpeningTag}${body}</${parts.tagName}>`
}

function patchLiteralCellXml(cellXml: string, value: LiteralInput): string | null {
  const valueBody = literalCellBody(value)
  return valueBody ? rewriteCellXml(cellXml, valueBody.type, valueBody.body) : null
}

function literalCellXml(address: string, value: LiteralInput): string | null {
  const valueBody = literalCellBody(value)
  return valueBody
    ? `<c r="${escapeXmlAttribute(address)}"${valueBody.type === null ? '' : ` t="${valueBody.type}"`}>${valueBody.body}</c>`
    : null
}

function rowNumberForAddress(address: string): number | null {
  try {
    const row = decodeCellAddress(address).r + 1
    return Number.isSafeInteger(row) && row > 0 ? row : null
  } catch {
    return null
  }
}

function insertCellIntoExistingRow(sheetXml: string, address: string, cellXml: string): string | null {
  const rowNumber = rowNumberForAddress(address)
  if (rowNumber === null) {
    return null
  }
  const rowPattern = new RegExp(
    `<((?:[A-Za-z_][\\w.-]*:)?row)\\b(?<attributes>[^>]*\\br=(["'])${rowNumber}\\3[^>]*)(?:\\/>|>(?<body>[\\s\\S]*?)<\\/\\1>)`,
    'u',
  )
  if (!rowPattern.test(sheetXml)) {
    return null
  }
  return sheetXml.replace(rowPattern, (_match: string, tagName: string, attributes: string, _quote: string, body = '') => {
    return `<${tagName}${attributes}>${body}${cellXml}</${tagName}>`
  })
}

function insertRowIntoSheetData(sheetXml: string, address: string, cellXml: string): string | null {
  const rowNumber = rowNumberForAddress(address)
  if (rowNumber === null) {
    return null
  }
  const rowXml = `<row r="${String(rowNumber)}">${cellXml}</row>`
  const nextRowPattern = /<((?:[A-Za-z_][\w.-]*:)?row)\b[^>]*\br=(["'])([0-9]+)\2[^>]*(?:\/>|>[\s\S]*?<\/\1>)/gu
  for (const match of sheetXml.matchAll(nextRowPattern)) {
    const existingRow = Number(match[3])
    if (Number.isSafeInteger(existingRow) && existingRow > rowNumber) {
      return `${sheetXml.slice(0, match.index)}${rowXml}${sheetXml.slice(match.index)}`
    }
  }
  return sheetXml.replace(/<\/((?:[A-Za-z_][\w.-]*:)?sheetData)>/u, `${rowXml}</$1>`)
}

function insertLiteralCell(sheetXml: string, address: string, value: LiteralInput): string | null {
  const cellXml = literalCellXml(address, value)
  return cellXml ? (insertCellIntoExistingRow(sheetXml, address, cellXml) ?? insertRowIntoSheetData(sheetXml, address, cellXml)) : null
}

function updateWorksheetDimension(sheetXml: string, addresses: Iterable<string>): string {
  const decoded = [...addresses].flatMap((address) => {
    try {
      return [decodeCellAddress(address)]
    } catch {
      return []
    }
  })
  if (decoded.length === 0 || !/<(?:[A-Za-z_][\w.-]*:)?dimension\b/u.test(sheetXml)) {
    return sheetXml
  }
  return sheetXml.replace(/<((?:[A-Za-z_][\w.-]*:)?dimension)\b([^>]*)\/>/u, (match, tagName: string, attributes: string) => {
    const ref = readXmlAttribute(match, 'ref')
    let range: XlsxCellRange | null = null
    try {
      range = ref ? decodeCellRange(ref) : null
    } catch {
      range = null
    }
    for (const cell of decoded) {
      range = range
        ? {
            s: { r: Math.min(range.s.r, cell.r), c: Math.min(range.s.c, cell.c) },
            e: { r: Math.max(range.e.r, cell.r), c: Math.max(range.e.c, cell.c) },
          }
        : { s: cell, e: cell }
    }
    return range ? `<${tagName}${setXmlAttribute(attributes, 'ref', encodeCellRange(range))}/>` : match
  })
}

function patchWorksheetXml(sheetXml: string, patch: WorksheetPatch): string | null {
  const patchedLiterals = new Set<string>()
  let failed = false
  let nextXml = sheetXml.replace(worksheetCellElementPattern, (cellXml: string) => {
    const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
    const address = openingTag ? readXmlAttribute(openingTag, 'r') : null
    if (!address) {
      return cellXml
    }
    if (patch.literals.has(address)) {
      const patched = patchLiteralCellXml(cellXml, patch.literals.get(address) ?? null)
      if (!patched) {
        failed = true
        return cellXml
      }
      patchedLiterals.add(address)
      return patched
    }
    return cellXml
  })
  if (failed) {
    return null
  }
  for (const [address, value] of patch.literals.entries()) {
    if (!patchedLiterals.has(address)) {
      const inserted = insertLiteralCell(nextXml, address, value)
      if (inserted === null) {
        return null
      }
      nextXml = inserted
    }
  }
  return updateWorksheetDimension(nextXml, patch.literals.keys())
}

function tryPrepareStreamingPatchedWorksheetEntry(
  zip: SourcePreservingZip,
  sheetPath: string,
  patch: WorksheetPatch,
): PreparedZipEntry | null {
  const compressedChunks: Uint8Array[] = []
  const sizes = tryWriteStreamingPatchedWorksheetEntry(zip, sheetPath, patch, (chunk) => {
    compressedChunks.push(chunk)
  })
  return sizes
    ? {
        compressedChunks,
        compressedSize: sizes.compressedSize,
        uncompressedSize: sizes.uncompressedSize,
        crc: sizes.crc,
      }
    : null
}

function tryPrepareStreamingPatchedWorksheetEntryFile(
  zip: SourcePreservingZip,
  sheetPath: string,
  patch: WorksheetPatch,
  compressedPath: string,
): FilePreparedZipEntry | null {
  const fd = openSync(compressedPath, 'w')
  let closed = false
  const close = (): void => {
    if (!closed) {
      closeSync(fd)
      closed = true
    }
  }
  try {
    const sizes = tryWriteStreamingPatchedWorksheetEntry(zip, sheetPath, patch, (chunk) => {
      writeAllSync(fd, chunk)
    })
    close()
    if (!sizes) {
      unlinkSync(compressedPath)
      return null
    }
    return { ...sizes, compressedPath }
  } catch {
    close()
    try {
      unlinkSync(compressedPath)
    } catch {
      // Best-effort cleanup only; the caller receives the export failure below.
    }
    return null
  } finally {
    close()
  }
}

async function tryPrepareStreamingPatchedWorksheetEntryFileAsync(
  zip: SourcePreservingZip,
  sheetPath: string,
  patch: WorksheetPatch,
  compressedPath: string,
): Promise<FilePreparedZipEntry | null> {
  const fd = openSync(compressedPath, 'w')
  let closed = false
  const close = (): void => {
    if (!closed) {
      closeSync(fd)
      closed = true
    }
  }
  try {
    const sizes = await tryWriteNativeStreamingPatchedWorksheetEntry(zip, sheetPath, patch, (chunk) => {
      writeAllSync(fd, chunk)
    })
    close()
    if (!sizes) {
      unlinkSync(compressedPath)
      return null
    }
    return { ...sizes, compressedPath }
  } catch {
    close()
    try {
      unlinkSync(compressedPath)
    } catch {
      // Best-effort cleanup only; the caller receives the export failure below.
    }
    return null
  } finally {
    close()
  }
}

function tryWriteStreamingPatchedWorksheetEntry(
  zip: SourcePreservingZip,
  sheetPath: string,
  patch: WorksheetPatch,
  writeCompressedChunk: (chunk: Uint8Array) => void,
): PreparedZipEntrySizes | null {
  if (patch.literals.size === 0) {
    return null
  }
  const textDecoder = new TextDecoder()
  const textEncoder = new TextEncoder()
  const patchedLiterals = new Set<string>()
  let compressedSize = 0
  let uncompressedSize = 0
  let crcState = 0xffffffff
  let failed = false
  let buffer = ''
  let pendingDeflateChunk: Uint8Array | null = null
  const deflate = new Deflate((chunk) => {
    if (chunk.byteLength === 0) {
      return
    }
    writeCompressedChunk(chunk)
    compressedSize += chunk.byteLength
  })
  const emitText = (text: string): void => {
    if (text.length === 0) {
      return
    }
    const bytes = textEncoder.encode(text)
    uncompressedSize += bytes.byteLength
    crcState = crc32Update(crcState, bytes)
    if (pendingDeflateChunk) {
      deflate.push(pendingDeflateChunk)
    }
    pendingDeflateChunk = bytes
  }
  const processBuffer = (final: boolean): void => {
    if (buffer.length === 0 || failed) {
      return
    }
    const safeEnd = final ? buffer.length : Math.max(0, buffer.lastIndexOf('<'))
    if (safeEnd === 0 && !final) {
      return
    }
    const cellPattern = new RegExp(worksheetCellElementPattern.source, 'gu')
    let emitOffset = 0
    for (const match of buffer.matchAll(cellPattern)) {
      const cellXml = match[0]
      const matchStart = match.index
      const matchEnd = matchStart + cellXml.length
      if (matchEnd > safeEnd) {
        break
      }
      emitText(buffer.slice(emitOffset, matchStart))
      const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
      const address = openingTag ? readXmlAttribute(openingTag, 'r') : null
      if (address && patch.literals.has(address)) {
        const patched = patchLiteralCellXml(cellXml, patch.literals.get(address) ?? null)
        if (!patched) {
          failed = true
          return
        }
        patchedLiterals.add(address)
        emitText(patched)
      } else {
        emitText(cellXml)
      }
      emitOffset = matchEnd
    }
    emitText(buffer.slice(emitOffset, safeEnd))
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
  if (!streamed || failed) {
    return null
  }
  buffer += textDecoder.decode()
  processBuffer(true)
  if (failed || patchedLiterals.size !== patch.literals.size) {
    return null
  }
  deflate.push(pendingDeflateChunk ?? new Uint8Array(), true)
  return {
    compressedSize,
    uncompressedSize,
    crc: crc32Finalize(crcState),
  }
}

async function tryWriteNativeStreamingPatchedWorksheetEntry(
  zip: SourcePreservingZip,
  sheetPath: string,
  patch: WorksheetPatch,
  writeCompressedChunk: (chunk: Uint8Array) => void,
): Promise<PreparedZipEntrySizes | null> {
  if (patch.literals.size === 0) {
    return null
  }
  const [{ once }, { createDeflateRaw }, { finished }] = await Promise.all([
    import('node:events'),
    import('node:zlib'),
    import('node:stream/promises'),
  ])
  const textDecoder = new TextDecoder()
  const textEncoder = new TextEncoder()
  const patchedLiterals = new Set<string>()
  let compressedSize = 0
  let uncompressedSize = 0
  let crcState = 0xffffffff
  let failed = false
  let buffer = ''
  const deflate = createDeflateRaw()
  deflate.on('data', (chunk: Uint8Array | Buffer) => {
    const bytes = chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk)
    if (bytes.byteLength === 0) {
      return
    }
    writeCompressedChunk(bytes)
    compressedSize += bytes.byteLength
  })
  const emitText = async (text: string): Promise<void> => {
    if (text.length === 0) {
      return
    }
    const bytes = textEncoder.encode(text)
    uncompressedSize += bytes.byteLength
    crcState = crc32Update(crcState, bytes)
    if (!deflate.write(bytes)) {
      await once(deflate, 'drain')
    }
  }
  const processBuffer = async (final: boolean): Promise<void> => {
    if (buffer.length === 0 || failed) {
      return
    }
    const safeEnd = final ? buffer.length : Math.max(0, buffer.lastIndexOf('<'))
    if (safeEnd === 0 && !final) {
      return
    }
    const cellPattern = new RegExp(worksheetCellElementPattern.source, 'gu')
    let emitOffset = 0
    for (const match of buffer.matchAll(cellPattern)) {
      const cellXml = match[0]
      const matchStart = match.index
      const matchEnd = matchStart + cellXml.length
      if (matchEnd > safeEnd) {
        break
      }
      // oxlint-disable-next-line eslint(no-await-in-loop) -- XML chunks must be emitted in worksheet order to preserve zlib backpressure.
      await emitText(buffer.slice(emitOffset, matchStart))
      const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
      const address = openingTag ? readXmlAttribute(openingTag, 'r') : null
      if (address && patch.literals.has(address)) {
        const patched = patchLiteralCellXml(cellXml, patch.literals.get(address) ?? null)
        if (!patched) {
          failed = true
          return
        }
        patchedLiterals.add(address)
        // oxlint-disable-next-line eslint(no-await-in-loop) -- Patched cell XML must remain ordered in the native deflate stream.
        await emitText(patched)
      } else {
        // oxlint-disable-next-line eslint(no-await-in-loop) -- Unpatched cell XML must remain ordered in the native deflate stream.
        await emitText(cellXml)
      }
      emitOffset = matchEnd
    }
    await emitText(buffer.slice(emitOffset, safeEnd))
    buffer = buffer.slice(safeEnd)
  }
  try {
    const streamed = await forEachInflatedXlsxZipEntryChunkAsync(
      zip,
      sheetPath,
      async (chunk) => {
        buffer += textDecoder.decode(chunk, { stream: true })
        await processBuffer(false)
      },
      { chunkSize: 64 * 1024, forceStreamingInflate: true },
    )
    if (!streamed || failed) {
      deflate.destroy()
      return null
    }
    buffer += textDecoder.decode()
    await processBuffer(true)
    if (failed || patchedLiterals.size !== patch.literals.size) {
      deflate.destroy()
      return null
    }
    deflate.end()
    await finished(deflate)
    return {
      compressedSize,
      uncompressedSize,
      crc: crc32Finalize(crcState),
    }
  } catch {
    deflate.destroy()
    return null
  }
}

function literalPatchesBySheet(patches: readonly ImportedXlsxSourceCellPatch[]): Map<string, Map<string, LiteralInput>> {
  const output = new Map<string, Map<string, LiteralInput>>()
  for (const patch of patches) {
    const sheetPatches = output.get(patch.sheetName) ?? new Map<string, LiteralInput>()
    sheetPatches.set(patch.address, patch.value)
    output.set(patch.sheetName, sheetPatches)
  }
  return output
}

function removeCalcChain(zip: Record<string, Uint8Array>): void {
  delete zip['xl/calcChain.xml']
  const contentTypesXml = getZipText(zip, '[Content_Types].xml')
  if (contentTypesXml) {
    setZipText(
      zip,
      '[Content_Types].xml',
      contentTypesXml.replace(/<Override\b(?=[^>]*\bPartName=(["'])\/xl\/calcChain\.xml\1)(?:[^>"']|"[^"]*"|'[^']*')*\/>/u, ''),
    )
  }
  const workbookRelationshipsXml = getZipText(zip, 'xl/_rels/workbook.xml.rels')
  if (workbookRelationshipsXml) {
    setZipText(
      zip,
      'xl/_rels/workbook.xml.rels',
      workbookRelationshipsXml.replace(
        /<Relationship\b(?=[^>]*(?:\bTarget=(["'])calcChain\.xml\1|\bType=(["'])http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain\2))(?:[^>"']|"[^"]*"|'[^']*')*\/>/u,
        '',
      ),
    )
  }
}

function ensureWorkbookRecalculation(zip: Record<string, Uint8Array>): boolean {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return false
  }
  const nextWorkbookXml = /<(?:[A-Za-z_][\w.-]*:)?calcPr\b/u.test(workbookXml)
    ? workbookXml.replace(
        /<((?:[A-Za-z_][\w.-]*:)?calcPr)\b([^>]*)(?:\/>|>[\s\S]*?<\/\1>)/u,
        (_match, tagName: string, attributes: string) => {
          const opening = `<${tagName}${attributes}/>`
          return setXmlAttribute(setXmlAttribute(setXmlAttribute(opening, 'calcMode', 'auto'), 'fullCalcOnLoad', '1'), 'forceFullCalc', '1')
        },
      )
    : workbookXml.replace(/<\/((?:[A-Za-z_][\w.-]*:)?workbook)>/u, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></$1>')
  setZipText(zip, 'xl/workbook.xml', nextWorkbookXml)
  return true
}

export function tryExportSourcePreservingXlsx(snapshot: WorkbookSnapshot, source: ImportedXlsxSourceReference): Uint8Array | null {
  const literalPatches = readImportedXlsxSourceCellPatches(snapshot)
  if (literalPatches.length === 0) {
    return null
  }
  const zip = readSourcePreservingZip(source)
  const preparedEntries = new Map<string, PreparedZipEntry>()
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  const sheetPathsByName = new Map(
    workbookSheetPathEntriesFromSource(
      zip,
      orderedSheets.map((sheet) => sheet.name),
    ).map((entry) => [entry.name, entry.path]),
  )
  const literalsBySheet = literalPatchesBySheet(literalPatches)
  for (const sheet of orderedSheets) {
    const literals = literalsBySheet.get(sheet.name) ?? new Map<string, LiteralInput>()
    if (literals.size === 0) {
      continue
    }
    const sheetPath = sheetPathsByName.get(sheet.name)
    if (!sheetPath) {
      return null
    }
    const patch = { literals }
    const streamingPatch = tryPrepareStreamingPatchedWorksheetEntry(zip, sheetPath, patch)
    if (streamingPatch) {
      preparedEntries.set(sheetPath, streamingPatch)
      continue
    }
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml) {
      return null
    }
    const patchedXml = patchWorksheetXml(sheetXml, patch)
    if (patchedXml === null) {
      return null
    }
    setZipText(zip, sheetPath, patchedXml)
  }
  removeCalcChain(zip)
  if (!ensureWorkbookRecalculation(zip)) {
    return null
  }
  return zipSourcePreservingEntries(zip, preparedEntries)
}

export function tryExportSourcePreservingXlsxToFile(
  snapshot: WorkbookSnapshot,
  source: ImportedXlsxSourceReference,
  outputPath: string,
): XlsxSourceLiteralPatchFileExportResult | null {
  const literalPatches = readImportedXlsxSourceCellPatches(snapshot)
  if (literalPatches.length === 0) {
    return null
  }
  const zip = readSourcePreservingZip(source)
  const preparedEntries = new Map<string, FilePreparedZipEntry>()
  const preparedEntryPaths: string[] = []
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  const sheetPathsByName = new Map(
    workbookSheetPathEntriesFromSource(
      zip,
      orderedSheets.map((sheet) => sheet.name),
    ).map((entry) => [entry.name, entry.path]),
  )
  const literalsBySheet = literalPatchesBySheet(literalPatches)
  for (const sheet of orderedSheets) {
    const literals = literalsBySheet.get(sheet.name) ?? new Map<string, LiteralInput>()
    if (literals.size === 0) {
      continue
    }
    const sheetPath = sheetPathsByName.get(sheet.name)
    if (!sheetPath) {
      cleanupTemporaryFiles(preparedEntryPaths)
      return null
    }
    const patch = { literals }
    const preparedEntryPath = `${outputPath}.${randomUUID()}.entry.tmp`
    const preparedEntry = tryPrepareStreamingPatchedWorksheetEntryFile(zip, sheetPath, patch, preparedEntryPath)
    if (!preparedEntry) {
      cleanupTemporaryFiles(preparedEntryPaths)
      return null
    }
    preparedEntryPaths.push(preparedEntryPath)
    preparedEntries.set(sheetPath, preparedEntry)
  }
  removeCalcChain(zip)
  if (!ensureWorkbookRecalculation(zip)) {
    cleanupTemporaryFiles(preparedEntryPaths)
    return null
  }
  const temporaryOutputPath = `${outputPath}.${randomUUID()}.tmp`
  try {
    const bytesWritten = zipSourcePreservingEntriesToFile(zip, preparedEntries, temporaryOutputPath)
    renameSync(temporaryOutputPath, outputPath)
    return { bytesWritten }
  } catch {
    try {
      unlinkSync(temporaryOutputPath)
    } catch {
      // Best-effort cleanup only; the caller receives the export failure below.
    }
    return null
  } finally {
    cleanupTemporaryFiles(preparedEntryPaths)
  }
}

export async function tryExportSourcePreservingXlsxToFileAsync(
  snapshot: WorkbookSnapshot,
  source: ImportedXlsxSourceReference,
  outputPath: string,
): Promise<XlsxSourceLiteralPatchFileExportResult | null> {
  const literalPatches = readImportedXlsxSourceCellPatches(snapshot)
  if (literalPatches.length === 0) {
    return null
  }
  const zip = readSourcePreservingZip(source)
  const preparedEntries = new Map<string, FilePreparedZipEntry>()
  const preparedEntryPaths: string[] = []
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  const sheetPathsByName = new Map(
    workbookSheetPathEntriesFromSource(
      zip,
      orderedSheets.map((sheet) => sheet.name),
    ).map((entry) => [entry.name, entry.path]),
  )
  const literalsBySheet = literalPatchesBySheet(literalPatches)
  for (const sheet of orderedSheets) {
    const literals = literalsBySheet.get(sheet.name) ?? new Map<string, LiteralInput>()
    if (literals.size === 0) {
      continue
    }
    const sheetPath = sheetPathsByName.get(sheet.name)
    if (!sheetPath) {
      cleanupTemporaryFiles(preparedEntryPaths)
      return null
    }
    const patch = { literals }
    const preparedEntryPath = `${outputPath}.${randomUUID()}.entry.tmp`
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Worksheets are prepared sequentially so each native stream and temp file is closed before the next.
    const preparedEntry = await tryPrepareStreamingPatchedWorksheetEntryFileAsync(zip, sheetPath, patch, preparedEntryPath)
    if (!preparedEntry) {
      cleanupTemporaryFiles(preparedEntryPaths)
      return null
    }
    preparedEntryPaths.push(preparedEntryPath)
    preparedEntries.set(sheetPath, preparedEntry)
  }
  removeCalcChain(zip)
  if (!ensureWorkbookRecalculation(zip)) {
    cleanupTemporaryFiles(preparedEntryPaths)
    return null
  }
  const temporaryOutputPath = `${outputPath}.${randomUUID()}.tmp`
  try {
    const bytesWritten = zipSourcePreservingEntriesToFile(zip, preparedEntries, temporaryOutputPath)
    renameSync(temporaryOutputPath, outputPath)
    return { bytesWritten }
  } catch {
    try {
      unlinkSync(temporaryOutputPath)
    } catch {
      // Best-effort cleanup only; the caller receives the export failure below.
    }
    return null
  } finally {
    cleanupTemporaryFiles(preparedEntryPaths)
  }
}

function cleanupTemporaryFiles(paths: readonly string[]): void {
  for (const path of paths) {
    try {
      unlinkSync(path)
    } catch {
      // Temporary file cleanup is best-effort only.
    }
  }
}

export function exportXlsxSourceLiteralPatches(input: XlsxSourceLiteralPatchExportInput): Uint8Array {
  return exportBiligXlsxSourceLiteralPatches(input)
}

export function exportXlsxSourceLiteralPatchesToFile(input: XlsxSourceLiteralPatchFileExportInput): XlsxSourceLiteralPatchFileExportResult {
  return exportBiligXlsxSourceLiteralPatchesToFile(input)
}

export async function exportXlsxSourceLiteralPatchesToFileAsync(
  input: XlsxSourceLiteralPatchFileExportInput,
): Promise<XlsxSourceLiteralPatchFileExportResult> {
  return exportBiligXlsxSourceLiteralPatchesToFileAsync(input)
}
