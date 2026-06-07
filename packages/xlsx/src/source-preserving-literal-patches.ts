import { randomUUID } from 'node:crypto'
import { closeSync, openSync, renameSync, unlinkSync } from 'node:fs'
import { deflateRawSync } from 'node:zlib'
import { Deflate } from 'fflate-stream'

import {
  crc32Finalize,
  crc32Update,
  type FilePreparedZipEntry,
  type PreparedZipEntry,
  type PreparedZipEntrySizes,
  writeAllSync,
  zipSourcePreservingEntries,
  zipSourcePreservingEntriesToFile,
} from './source-preserving-zip.js'
import {
  readXlsxZipEntries,
  readXlsxZipEntriesLazy,
  readXlsxZipEntriesLazyFromByteSource,
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  setZipText,
  type XlsxZipByteSource,
  type XlsxZipEntries,
} from './zip-reader.js'
import { workbookSheetPathEntriesFromSource } from './workbook-sheet-paths.js'
import { decodeCellAddress, decodeCellRange, encodeCellRange, type XlsxCellRange } from './address.js'
import {
  escapeXmlAttribute,
  escapeXmlText,
  readXmlAttribute,
  setXmlAttribute,
  worksheetCellElementPattern,
  worksheetCellOpeningTagPattern,
} from './xml.js'

export interface XlsxScalarPatchErrorValue {
  readonly kind: 'error'
  readonly value: string
}

export type XlsxScalarPatchValue = string | number | boolean | null | XlsxScalarPatchErrorValue

export interface XlsxSourceReader {
  readonly byteLength: number
  readBytes(): Uint8Array
  readRange?(start: number, end: number): Uint8Array
  readRangeInto?(start: number, end: number, target: Uint8Array): Uint8Array
  inflateRawRange?(start: number, end: number): Uint8Array
  inflateRawRangeChunksAsync?(
    start: number,
    end: number,
    onChunk: (chunk: Uint8Array) => boolean | void | Promise<boolean | void>,
    options: { readonly chunkSize: number },
  ): Promise<boolean>
  release?(): void
}

export interface XlsxSourceLiteralPatch {
  readonly sheetName: string
  readonly address: string
  readonly value: XlsxScalarPatchValue
  readonly preserveFormula?: boolean
}

export interface XlsxSourceTextPatch {
  readonly path: string
  patchText(text: string): string
}

export interface XlsxSourceLiteralPatchExportInput {
  readonly source: Uint8Array | XlsxSourceReader
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly textPatches?: readonly XlsxSourceTextPatch[]
  readonly sheetNames?: readonly string[]
  readonly workbookName?: string
}

export interface XlsxSourceLiteralPatchFileExportInput extends XlsxSourceLiteralPatchExportInput {
  readonly outputPath: string
}

export interface XlsxSourceLiteralPatchFileExportResult {
  readonly bytesWritten: number
}

interface WorksheetPatch {
  readonly literals: ReadonlyMap<string, XlsxScalarCellPatch>
}

type SourcePreservingZip = XlsxZipEntries

interface XlsxScalarCellPatch {
  readonly value: XlsxScalarPatchValue
  readonly preserveFormula: boolean
}

function isXlsxZipByteSource(source: Uint8Array | XlsxSourceReader): source is XlsxSourceReader & XlsxZipByteSource {
  return !(source instanceof Uint8Array) && typeof source.readRange === 'function'
}

function readSourcePreservingZip(source: Uint8Array | XlsxSourceReader): SourcePreservingZip {
  if (source instanceof Uint8Array) {
    return readXlsxZipEntriesLazy(source)
  }
  if (isXlsxZipByteSource(source)) {
    const lazyZip = readXlsxZipEntriesLazyFromByteSource(source)
    if (lazyZip) {
      return lazyZip
    }
  }
  return readXlsxZipEntries(source.readBytes())
}

function removeXmlAttribute(tag: string, name: string): string {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  return tag.replace(new RegExp(`\\s${escapedName}=(["'])[\\s\\S]*?\\1`, 'u'), '')
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

function isXlsxScalarPatchErrorValue(value: XlsxScalarPatchValue): value is XlsxScalarPatchErrorValue {
  return typeof value === 'object' && value !== null && value.kind === 'error'
}

function literalCellBody(value: XlsxScalarPatchValue): { readonly type: string | null; readonly body: string } | null {
  if (value === null) {
    return { type: null, body: '' }
  }
  if (isXlsxScalarPatchErrorValue(value)) {
    return { type: 'e', body: `<v>${escapeXmlText(value.value)}</v>` }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { type: null, body: `<v>${String(value)}</v>` } : null
  }
  if (typeof value === 'boolean') {
    return { type: 'b', body: `<v>${value ? '1' : '0'}</v>` }
  }
  return { type: 'inlineStr', body: inlineStringBody(value) }
}

function formulaCachedValue(value: XlsxScalarPatchValue): { readonly type: string | null; readonly valueXml: string } | null {
  if (value === null) {
    return { type: null, valueXml: '' }
  }
  if (isXlsxScalarPatchErrorValue(value)) {
    return { type: 'e', valueXml: `<v>${escapeXmlText(value.value)}</v>` }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { type: null, valueXml: `<v>${String(value)}</v>` } : null
  }
  if (typeof value === 'boolean') {
    return { type: 'b', valueXml: `<v>${value ? '1' : '0'}</v>` }
  }
  return { type: 'str', valueXml: `<v>${escapeXmlText(value)}</v>` }
}

function rewriteCellXml(cellXml: string, type: string | null, body: string): string | null {
  const parts = cellParts(cellXml)
  if (!parts) {
    return null
  }
  const nextOpeningTag = setCellType(parts.openingTag, type)
  return `${nextOpeningTag.endsWith('/>') ? `${nextOpeningTag.slice(0, -2)}>` : nextOpeningTag}${body}</${parts.tagName}>`
}

function patchLiteralCellXml(cellXml: string, value: XlsxScalarPatchValue): string | null {
  const valueBody = literalCellBody(value)
  return valueBody ? rewriteCellXml(cellXml, valueBody.type, valueBody.body) : null
}

const formulaElementPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b(?:[^/>]*\/>|[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?f>)/u
const cachedValueElementPattern = /<((?:[A-Za-z_][\w.-]*:)?v)\b[^>]*>[\s\S]*?<\/\1>/u

function patchFormulaCacheCellXml(cellXml: string, value: XlsxScalarPatchValue): string | null {
  const valueBody = formulaCachedValue(value)
  const parts = cellParts(cellXml)
  if (!valueBody || !parts) {
    return null
  }
  const formulaMatch = formulaElementPattern.exec(parts.body)
  if (!formulaMatch) {
    return null
  }
  const nextOpeningTag = setCellType(parts.openingTag, valueBody.type)
  const formulaEnd = formulaMatch.index + formulaMatch[0].length
  const nextBody =
    valueBody.valueXml.length === 0
      ? parts.body.replace(cachedValueElementPattern, '')
      : cachedValueElementPattern.test(parts.body)
        ? parts.body.replace(cachedValueElementPattern, valueBody.valueXml)
        : `${parts.body.slice(0, formulaEnd)}${valueBody.valueXml}${parts.body.slice(formulaEnd)}`
  return `${nextOpeningTag.endsWith('/>') ? `${nextOpeningTag.slice(0, -2)}>` : nextOpeningTag}${nextBody}</${parts.tagName}>`
}

function patchScalarCellXml(cellXml: string, patch: XlsxScalarCellPatch): string | null {
  return patch.preserveFormula ? patchFormulaCacheCellXml(cellXml, patch.value) : patchLiteralCellXml(cellXml, patch.value)
}

function literalCellXml(address: string, value: XlsxScalarPatchValue): string | null {
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

function insertLiteralCell(sheetXml: string, address: string, value: XlsxScalarPatchValue): string | null {
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
      const patched = patchScalarCellXml(cellXml, patch.literals.get(address) ?? { value: null, preserveFormula: false })
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
      const inserted = value.preserveFormula ? null : insertLiteralCell(nextXml, address, value.value)
      if (inserted === null) {
        return null
      }
      nextXml = inserted
    }
  }
  return updateWorksheetDimension(nextXml, patch.literals.keys())
}

function deflatedPreparedEntry(bytes: Uint8Array): PreparedZipEntry {
  const compressed = deflateRawSync(bytes)
  return {
    compressedChunks: [compressed],
    compressedSize: compressed.byteLength,
    uncompressedSize: bytes.byteLength,
    crc: crc32Finalize(crc32Update(0xffffffff, bytes)),
  }
}

function writeDeflatedPreparedEntryFile(bytes: Uint8Array, compressedPath: string): FilePreparedZipEntry {
  const prepared = deflatedPreparedEntry(bytes)
  const fd = openSync(compressedPath, 'w')
  try {
    for (const chunk of prepared.compressedChunks) {
      writeAllSync(fd, chunk)
    }
  } finally {
    closeSync(fd)
  }
  return {
    compressedPath,
    compressedSize: prepared.compressedSize,
    uncompressedSize: prepared.uncompressedSize,
    crc: prepared.crc,
  }
}

function applyTextPartPatches(
  zip: SourcePreservingZip,
  patches: readonly XlsxSourceTextPatch[] | undefined,
  prepare: (path: string, text: string) => void,
): void {
  for (const patch of patches ?? []) {
    const text = getZipText(zip, patch.path)
    if (text === null) {
      continue
    }
    const nextText = patch.patchText(text)
    if (nextText === text) {
      continue
    }
    prepare(patch.path, nextText)
  }
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
  const patchAddressNeedles = [...patch.literals.keys()].flatMap((address) => {
    const escaped = escapeXmlAttribute(address)
    return [`r="${escaped}"`, `r='${escaped}'`]
  })
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
    const safeXml = buffer.slice(0, safeEnd)
    if (!patchAddressNeedles.some((needle) => safeXml.includes(needle))) {
      emitText(safeXml)
      buffer = buffer.slice(safeEnd)
      return
    }
    const cellPattern = new RegExp(worksheetCellElementPattern.source, 'gu')
    let emitOffset = 0
    for (const match of safeXml.matchAll(cellPattern)) {
      const cellXml = match[0]
      const matchStart = match.index
      const matchEnd = matchStart + cellXml.length
      emitText(buffer.slice(emitOffset, matchStart))
      const openingTag = worksheetCellOpeningTagPattern.exec(cellXml)?.[0]
      const address = openingTag ? readXmlAttribute(openingTag, 'r') : null
      if (address && patch.literals.has(address)) {
        const patched = patchScalarCellXml(cellXml, patch.literals.get(address) ?? { value: null, preserveFormula: false })
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
    emitText(safeXml.slice(emitOffset))
    buffer = buffer.slice(safeEnd)
  }
  try {
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
  } catch {
    return null
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

function literalPatchesBySheet(patches: readonly XlsxSourceLiteralPatch[]): Map<string, Map<string, XlsxScalarCellPatch>> {
  const output = new Map<string, Map<string, XlsxScalarCellPatch>>()
  for (const patch of patches) {
    const sheetPatches = output.get(patch.sheetName) ?? new Map<string, XlsxScalarCellPatch>()
    sheetPatches.set(patch.address, { value: patch.value, preserveFormula: patch.preserveFormula === true })
    output.set(patch.sheetName, sheetPatches)
  }
  return output
}

function uniquePatchedSheetNames(patches: readonly XlsxSourceLiteralPatch[]): string[] {
  return [...new Set(patches.map((patch) => patch.sheetName))]
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

function normalizedPatchInput(input: XlsxSourceLiteralPatchExportInput): {
  readonly source: Uint8Array | XlsxSourceReader
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly textPatches: readonly XlsxSourceTextPatch[]
  readonly sheetNames: readonly string[]
} {
  return {
    source: input.source,
    patches: input.patches,
    textPatches: input.textPatches ?? [],
    sheetNames: input.sheetNames ?? uniquePatchedSheetNames(input.patches),
  }
}

export function exportXlsxSourceLiteralPatches(input: XlsxSourceLiteralPatchExportInput): Uint8Array {
  const { source, patches, textPatches, sheetNames } = normalizedPatchInput(input)
  const zip = readSourcePreservingZip(source)
  const preparedEntries = new Map<string, PreparedZipEntry>()
  const sheetPathsByName = new Map(workbookSheetPathEntriesFromSource(zip, sheetNames).map((entry) => [entry.name, entry.path]))
  const literalsBySheet = literalPatchesBySheet(patches)
  for (const sheetName of sheetNames) {
    const literals = literalsBySheet.get(sheetName) ?? new Map<string, XlsxScalarCellPatch>()
    if (literals.size === 0) {
      continue
    }
    const sheetPath = sheetPathsByName.get(sheetName)
    if (!sheetPath) {
      throw new Error(`Unable to resolve XLSX worksheet path for sheet: ${sheetName}`)
    }
    const streamingPatch = tryPrepareStreamingPatchedWorksheetEntry(zip, sheetPath, { literals })
    if (streamingPatch) {
      preparedEntries.set(sheetPath, streamingPatch)
      continue
    }
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml) {
      throw new Error(`Unable to read XLSX worksheet XML for sheet: ${sheetName}`)
    }
    const patchedXml = patchWorksheetXml(sheetXml, { literals })
    if (patchedXml === null) {
      throw new Error(`Unable to apply XLSX literal patches for sheet: ${sheetName}`)
    }
    preparedEntries.set(sheetPath, deflatedPreparedEntry(new TextEncoder().encode(patchedXml)))
  }
  applyTextPartPatches(zip, textPatches, (path, text) => {
    preparedEntries.set(path, deflatedPreparedEntry(new TextEncoder().encode(text)))
  })
  removeCalcChain(zip)
  if (!ensureWorkbookRecalculation(zip)) {
    throw new Error('Unable to update XLSX workbook recalculation settings')
  }
  return zipSourcePreservingEntries(zip, preparedEntries)
}

export function exportXlsxSourceLiteralPatchesToFile(input: XlsxSourceLiteralPatchFileExportInput): XlsxSourceLiteralPatchFileExportResult {
  const exported = exportXlsxSourceLiteralPatches(input)
  const temporaryOutputPath = `${input.outputPath}.${randomUUID()}.tmp`
  try {
    const fd = openSync(temporaryOutputPath, 'w')
    try {
      writeAllSync(fd, exported)
    } finally {
      closeSync(fd)
    }
    renameSync(temporaryOutputPath, input.outputPath)
    return { bytesWritten: exported.byteLength }
  } catch (error) {
    try {
      unlinkSync(temporaryOutputPath)
    } catch {
      // Best-effort cleanup only; the caller receives the export failure below.
    }
    throw error
  }
}

export async function exportXlsxSourceLiteralPatchesToFileAsync(
  input: XlsxSourceLiteralPatchFileExportInput,
): Promise<XlsxSourceLiteralPatchFileExportResult> {
  const { source, patches, textPatches, sheetNames } = normalizedPatchInput(input)
  const zip = readSourcePreservingZip(source)
  const preparedEntries = new Map<string, FilePreparedZipEntry>()
  const preparedEntryPaths: string[] = []
  const sheetPathsByName = new Map(workbookSheetPathEntriesFromSource(zip, sheetNames).map((entry) => [entry.name, entry.path]))
  const literalsBySheet = literalPatchesBySheet(patches)
  for (const sheetName of sheetNames) {
    const literals = literalsBySheet.get(sheetName) ?? new Map<string, XlsxScalarCellPatch>()
    if (literals.size === 0) {
      continue
    }
    const sheetPath = sheetPathsByName.get(sheetName)
    if (!sheetPath) {
      cleanupTemporaryFiles(preparedEntryPaths)
      throw new Error(`Unable to resolve XLSX worksheet path for sheet: ${sheetName}`)
    }
    const preparedEntryPath = `${input.outputPath}.${randomUUID()}.entry.tmp`
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Worksheets are prepared sequentially so each stream closes before the next.
    const preparedEntry = await tryPrepareStreamingPatchedWorksheetEntryFileAsync(zip, sheetPath, { literals }, preparedEntryPath)
    if (!preparedEntry) {
      cleanupTemporaryFiles(preparedEntryPaths)
      throw new Error(`Unable to apply XLSX literal patches for sheet: ${sheetName}`)
    }
    preparedEntryPaths.push(preparedEntryPath)
    preparedEntries.set(sheetPath, preparedEntry)
  }
  applyTextPartPatches(zip, textPatches, (path, text) => {
    const preparedEntryPath = `${input.outputPath}.${randomUUID()}.entry.tmp`
    const preparedEntry = writeDeflatedPreparedEntryFile(new TextEncoder().encode(text), preparedEntryPath)
    preparedEntryPaths.push(preparedEntryPath)
    preparedEntries.set(path, preparedEntry)
  })
  removeCalcChain(zip)
  if (!ensureWorkbookRecalculation(zip)) {
    cleanupTemporaryFiles(preparedEntryPaths)
    throw new Error('Unable to update XLSX workbook recalculation settings')
  }
  const temporaryOutputPath = `${input.outputPath}.${randomUUID()}.tmp`
  try {
    const bytesWritten = zipSourcePreservingEntriesToFile(zip, preparedEntries, temporaryOutputPath)
    renameSync(temporaryOutputPath, input.outputPath)
    return { bytesWritten }
  } catch (error) {
    try {
      unlinkSync(temporaryOutputPath)
    } catch {
      // Best-effort cleanup only; the caller receives the export failure below.
    }
    throw error
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
