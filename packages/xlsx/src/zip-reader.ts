import { inflateRawSync } from 'node:zlib'

export type XlsxZipEntries = Record<string, Uint8Array>
export type XlsxZipSource = Uint8Array | XlsxZipEntries

export interface XlsxZipByteSource {
  readonly byteLength: number
  readRange(start: number, end: number): Uint8Array
  readRangeInto?(start: number, end: number, target: Uint8Array): Uint8Array
  inflateRawRange?(start: number, end: number): Uint8Array
  inflateRawRangeChunksAsync?(
    start: number,
    end: number,
    onChunk: XlsxZipAsyncChunkConsumer,
    options: { readonly chunkSize: number },
  ): Promise<boolean>
  release?(): void
}

export interface XlsxZipEntryMetadata {
  readonly path: string
  readonly compressedSize: number
  readonly uncompressedSize: number
  readonly compressionMethod: number
}

export type XlsxZipChunkConsumer = (chunk: Uint8Array) => boolean | void
export type XlsxZipAsyncChunkConsumer = (chunk: Uint8Array) => boolean | void | Promise<boolean | void>

const localFileHeaderSignature = 0x04034b50
const centralDirectoryFileHeaderSignature = 0x02014b50
const endOfCentralDirectorySignature = 0x06054b50
const storedCompressionMethod = 0
const deflatedCompressionMethod = 8
const zip64Sentinel = 0xffffffff
const maxEndOfCentralDirectorySearch = 65_557
const defaultZipEntryChunkSize = 64 * 1024
const xlsxZipCentralDirectorySourceSymbol: unique symbol = Symbol('bilig.xlsxZipCentralDirectorySource')
const textDecoder = new TextDecoder()

interface CentralDirectoryEntry {
  readonly path: string
  readonly localHeaderOffset: number
  readonly compressedSize: number
  readonly uncompressedSize: number
  readonly compressionMethod: number
}

interface XlsxZipCentralDirectorySource {
  source: XlsxZipByteSource | null
  readonly entriesByPath: ReadonlyMap<string, CentralDirectoryEntry>
}

interface Uint8ArrayXlsxZipByteSource extends XlsxZipByteSource {
  readonly bytes: Uint8Array
}

type XlsxZipEntriesWithCentralDirectorySource = XlsxZipEntries & {
  readonly [xlsxZipCentralDirectorySourceSymbol]?: XlsxZipCentralDirectorySource
}

export function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/u, '')
}

export function readXlsxZipEntries(source: XlsxZipSource): XlsxZipEntries {
  if (!(source instanceof Uint8Array)) {
    return source
  }
  const entries = readCentralDirectoryEntries(byteSourceFromUint8Array(source))
  if (!entries) {
    throw new Error('Invalid XLSX ZIP central directory')
  }
  const output: XlsxZipEntries = {}
  const byteSource = byteSourceFromUint8Array(source)
  for (const entry of entries) {
    output[entry.path] = inflateCentralDirectoryEntry(byteSource, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod)
  }
  return output
}

export function readXlsxZipEntriesLazy(source: XlsxZipSource): XlsxZipEntries {
  if (!(source instanceof Uint8Array)) {
    return source
  }
  return readXlsxZipEntriesLazyFromByteSource(byteSourceFromUint8Array(source)) ?? readXlsxZipEntries(source)
}

export function readXlsxZipEntriesLazyFromByteSource(source: XlsxZipByteSource): XlsxZipEntries | null {
  const entries = readCentralDirectoryEntries(source)
  if (!entries) {
    return null
  }
  const output: XlsxZipEntries = {}
  const metadata: XlsxZipCentralDirectorySource = {
    source,
    entriesByPath: new Map(entries.map((entry) => [entry.path, entry])),
  }
  for (const entry of entries) {
    defineLazyZipEntry(output, metadata, entry)
  }
  defineLazyZipCentralDirectorySource(output, metadata)
  return output
}

export function readXlsxZipEntryMetadata(source: XlsxZipByteSource): readonly XlsxZipEntryMetadata[] | null {
  return (
    readCentralDirectoryEntries(source)?.map((entry) => ({
      path: entry.path,
      compressedSize: entry.compressedSize,
      uncompressedSize: entry.uncompressedSize,
      compressionMethod: entry.compressionMethod,
    })) ?? null
  )
}

export function getZipText(zip: XlsxZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? textDecoder.decode(file) : null
}

export function setZipText(zip: XlsxZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = new TextEncoder().encode(text)
}

export function forEachInflatedXlsxZipEntryChunk(
  zip: XlsxZipEntries,
  path: string,
  onChunk: XlsxZipChunkConsumer,
  options: { readonly chunkSize?: number } = {},
): boolean {
  const normalizedPath = normalizeZipPath(path)
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  const source = metadata?.source
  const entry = metadata?.entriesByPath.get(normalizedPath)
  if (metadata && (!source || !entry)) {
    return false
  }
  if (source && entry) {
    const inflated = inflateCentralDirectoryEntry(source, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod)
    emitInflatedChunks(inflated, options.chunkSize ?? defaultZipEntryChunkSize, onChunk)
    return true
  }
  const inflated = zip[normalizedPath]
  if (!inflated) {
    return false
  }
  emitInflatedChunks(inflated, options.chunkSize ?? defaultZipEntryChunkSize, onChunk)
  return true
}

export async function forEachInflatedXlsxZipEntryChunkAsync(
  zip: XlsxZipEntries,
  path: string,
  onChunk: XlsxZipAsyncChunkConsumer,
  options: { readonly chunkSize?: number } = {},
): Promise<boolean> {
  const normalizedPath = normalizeZipPath(path)
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  const source = metadata?.source
  const entry = metadata?.entriesByPath.get(normalizedPath)
  const chunkSize = options.chunkSize ?? defaultZipEntryChunkSize
  if (metadata && (!source || !entry)) {
    return false
  }
  if (!source || !entry) {
    const inflated = zip[normalizedPath]
    if (!inflated) {
      return false
    }
    await emitInflatedChunksAsync(inflated, chunkSize, onChunk)
    return true
  }
  await inflateCentralDirectoryEntryChunksAsync(source, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod, onChunk, {
    chunkSize,
  })
  return true
}

export function releaseLazyXlsxZipSource(zip: XlsxZipEntries): boolean {
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  if (!metadata?.source) {
    return false
  }
  metadata.source.release?.()
  metadata.source = null
  return true
}

export function readLazyXlsxZipSourceByteLength(zip: XlsxZipEntries): number | undefined {
  const metadata = (zip as XlsxZipEntriesWithCentralDirectorySource)[xlsxZipCentralDirectorySourceSymbol]
  return metadata ? (metadata.source?.byteLength ?? 0) : undefined
}

function byteSourceFromUint8Array(source: Uint8Array): Uint8ArrayXlsxZipByteSource {
  return {
    byteLength: source.byteLength,
    bytes: source,
    readRange(start, end) {
      return source.subarray(start, end)
    },
  }
}

function readCentralDirectoryEntries(source: XlsxZipByteSource): CentralDirectoryEntry[] | null {
  const endOfCentralDirectoryOffset = findEndOfCentralDirectoryOffset(source)
  if (endOfCentralDirectoryOffset === null) {
    return null
  }
  const endOfCentralDirectory = source.readRange(endOfCentralDirectoryOffset, endOfCentralDirectoryOffset + 22)
  const centralDirectoryOffset = readUint32(endOfCentralDirectory, 16)
  const centralDirectorySize = readUint32(endOfCentralDirectory, 12)
  if (
    centralDirectoryOffset === zip64Sentinel ||
    centralDirectorySize === zip64Sentinel ||
    centralDirectoryOffset + centralDirectorySize > source.byteLength
  ) {
    return null
  }
  const centralDirectory = source.readRange(centralDirectoryOffset, centralDirectoryOffset + centralDirectorySize)
  const entries: CentralDirectoryEntry[] = []
  let offset = 0
  const endOffset = centralDirectory.byteLength
  while (offset + 46 <= endOffset && readUint32(centralDirectory, offset) === centralDirectoryFileHeaderSignature) {
    const compressionMethod = readUint16(centralDirectory, offset + 10)
    const compressedSize = readUint32(centralDirectory, offset + 20)
    const uncompressedSize = readUint32(centralDirectory, offset + 24)
    const fileNameLength = readUint16(centralDirectory, offset + 28)
    const extraFieldLength = readUint16(centralDirectory, offset + 30)
    const fileCommentLength = readUint16(centralDirectory, offset + 32)
    const localHeaderOffset = readUint32(centralDirectory, offset + 42)
    const fileNameStart = offset + 46
    const fileNameEnd = fileNameStart + fileNameLength
    const nextOffset = fileNameEnd + extraFieldLength + fileCommentLength
    if (
      fileNameEnd > endOffset ||
      nextOffset > endOffset ||
      localHeaderOffset === zip64Sentinel ||
      localHeaderOffset + 30 > source.byteLength
    ) {
      return null
    }
    const path = normalizeZipPath(textDecoder.decode(centralDirectory.subarray(fileNameStart, fileNameEnd)))
    entries.push({ path, localHeaderOffset, compressedSize, uncompressedSize, compressionMethod })
    offset = nextOffset
  }
  return offset === endOffset ? entries : null
}

function defineLazyZipEntry(output: XlsxZipEntries, metadata: XlsxZipCentralDirectorySource, entry: CentralDirectoryEntry): void {
  Object.defineProperty(output, entry.path, {
    configurable: true,
    enumerable: true,
    get() {
      const source = metadata.source
      if (!source) {
        throw new Error('XLSX ZIP source has been released')
      }
      const bytes = inflateCentralDirectoryEntry(source, entry.localHeaderOffset, entry.compressedSize, entry.compressionMethod)
      Object.defineProperty(output, entry.path, {
        configurable: true,
        enumerable: true,
        value: bytes,
        writable: true,
      })
      return bytes
    },
  })
}

function defineLazyZipCentralDirectorySource(output: XlsxZipEntries, metadata: XlsxZipCentralDirectorySource): void {
  Object.defineProperty(output, xlsxZipCentralDirectorySourceSymbol, {
    configurable: false,
    enumerable: false,
    value: metadata,
  })
}

function inflateCentralDirectoryEntry(
  source: XlsxZipByteSource,
  localHeaderOffset: number,
  compressedSize: number,
  compressionMethod: number,
): Uint8Array {
  const { dataStart, dataEnd } = readEntryDataRange(source, localHeaderOffset, compressedSize)
  if (compressionMethod === storedCompressionMethod) {
    return new Uint8Array(source.readRange(dataStart, dataEnd))
  }
  if (compressionMethod === deflatedCompressionMethod) {
    return source.inflateRawRange ? source.inflateRawRange(dataStart, dataEnd) : inflateRawSync(source.readRange(dataStart, dataEnd))
  }
  throw new Error(`Unsupported XLSX compression method: ${String(compressionMethod)}`)
}

async function inflateCentralDirectoryEntryChunksAsync(
  source: XlsxZipByteSource,
  localHeaderOffset: number,
  compressedSize: number,
  compressionMethod: number,
  onChunk: XlsxZipAsyncChunkConsumer,
  options: { readonly chunkSize: number },
): Promise<void> {
  const { dataStart, dataEnd } = readEntryDataRange(source, localHeaderOffset, compressedSize)
  if (compressionMethod === storedCompressionMethod) {
    await forEachSourceChunkAsync(source, dataStart, dataEnd, options.chunkSize, onChunk)
    return
  }
  if (compressionMethod !== deflatedCompressionMethod) {
    throw new Error(`Unsupported XLSX compression method: ${String(compressionMethod)}`)
  }
  if (
    source.inflateRawRangeChunksAsync &&
    (await source.inflateRawRangeChunksAsync(dataStart, dataEnd, onChunk, { chunkSize: options.chunkSize }))
  ) {
    return
  }
  if (await inflateCentralDirectoryEntryChunksWithNodeZlibAsync(source, dataStart, dataEnd, onChunk, options.chunkSize)) {
    return
  }
  await emitInflatedChunksAsync(
    source.inflateRawRange ? source.inflateRawRange(dataStart, dataEnd) : inflateRawSync(source.readRange(dataStart, dataEnd)),
    options.chunkSize,
    onChunk,
  )
}

function readEntryDataRange(
  source: XlsxZipByteSource,
  localHeaderOffset: number,
  compressedSize: number,
): { readonly dataStart: number; readonly dataEnd: number } {
  const localHeader = readSourceRange(
    source,
    localHeaderOffset,
    localHeaderOffset + 30,
    source.readRangeInto ? new Uint8Array(30) : undefined,
  )
  if (readUint32(localHeader, 0) !== localFileHeaderSignature) {
    throw new Error('Invalid XLSX local file header')
  }
  const fileNameLength = readUint16(localHeader, 26)
  const extraFieldLength = readUint16(localHeader, 28)
  const dataStart = localHeaderOffset + 30 + fileNameLength + extraFieldLength
  const dataEnd = dataStart + compressedSize
  if (dataEnd > source.byteLength) {
    throw new Error('Invalid XLSX compressed data range')
  }
  return { dataStart, dataEnd }
}

async function inflateCentralDirectoryEntryChunksWithNodeZlibAsync(
  source: XlsxZipByteSource,
  dataStart: number,
  dataEnd: number,
  onChunk: XlsxZipAsyncChunkConsumer,
  chunkSize: number,
): Promise<boolean> {
  try {
    const [{ createInflateRaw }, { Readable }] = await Promise.all([import('node:zlib'), import('node:stream')])
    let offset = dataStart
    const compressedChunkSize = Math.max(1, Math.trunc(chunkSize))
    const compressedStream = new Readable({
      read() {
        if (offset >= dataEnd) {
          this.push(null)
          return
        }
        const end = Math.min(dataEnd, offset + compressedChunkSize)
        this.push(source.readRange(offset, end))
        offset = end
      },
    })
    const inflate = createInflateRaw()
    try {
      for await (const chunk of compressedStream.pipe(inflate)) {
        const bytes = chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk)
        if (!(await emitInflatedChunksAsync(bytes, chunkSize, onChunk))) {
          compressedStream.destroy()
          inflate.destroy()
          return true
        }
      }
    } finally {
      compressedStream.destroy()
      inflate.destroy()
    }
    return true
  } catch {
    return false
  }
}

function emitInflatedChunks(chunk: Uint8Array, chunkSize: number, onChunk: XlsxZipChunkConsumer): boolean {
  for (let offset = 0; offset < chunk.byteLength; offset += chunkSize) {
    if (onChunk(chunk.subarray(offset, Math.min(chunk.byteLength, offset + chunkSize))) === false) {
      return false
    }
  }
  return true
}

async function emitInflatedChunksAsync(chunk: Uint8Array, chunkSize: number, onChunk: XlsxZipAsyncChunkConsumer): Promise<boolean> {
  for (let offset = 0; offset < chunk.byteLength; offset += chunkSize) {
    // oxlint-disable-next-line eslint(no-await-in-loop) -- ZIP stream chunks must preserve order and backpressure.
    if ((await onChunk(chunk.subarray(offset, Math.min(chunk.byteLength, offset + chunkSize)))) === false) {
      return false
    }
  }
  return true
}

async function forEachSourceChunkAsync(
  source: XlsxZipByteSource,
  start: number,
  end: number,
  chunkSize: number,
  onChunk: XlsxZipAsyncChunkConsumer,
): Promise<void> {
  const normalizedChunkSize = Math.max(1, Math.trunc(chunkSize))
  const scratch = source.readRangeInto ? new Uint8Array(normalizedChunkSize) : undefined
  if (start === end) {
    await onChunk(readSourceRange(source, start, end, scratch))
    return
  }
  for (let offset = start; offset < end; offset += normalizedChunkSize) {
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Source chunks must preserve read order and consumer backpressure.
    if ((await onChunk(readSourceRange(source, offset, Math.min(end, offset + normalizedChunkSize), scratch))) === false) {
      return
    }
  }
}

function readSourceRange(source: XlsxZipByteSource, start: number, end: number, scratch: Uint8Array | undefined): Uint8Array {
  const length = end - start
  if (!scratch || !source.readRangeInto || length > scratch.byteLength) {
    return source.readRange(start, end)
  }
  const chunk = source.readRangeInto(start, end, scratch)
  if (chunk.byteLength !== length) {
    throw new Error('XLSX ZIP byte source returned an invalid chunk length')
  }
  return chunk
}

function findEndOfCentralDirectoryOffset(source: XlsxZipByteSource): number | null {
  const tailStart = Math.max(0, source.byteLength - maxEndOfCentralDirectorySearch)
  const tail = source.readRange(tailStart, source.byteLength)
  for (let offset = tail.byteLength - 22; offset >= 0; offset -= 1) {
    if (readUint32(tail, offset) === endOfCentralDirectorySignature) {
      return tailStart + offset
    }
  }
  return null
}

function readUint16(source: Uint8Array, offset: number): number {
  return source[offset]! | (source[offset + 1]! << 8)
}

function readUint32(source: Uint8Array, offset: number): number {
  return (source[offset]! | (source[offset + 1]! << 8) | (source[offset + 2]! << 16) | (source[offset + 3]! << 24)) >>> 0
}
