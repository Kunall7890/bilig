import { closeSync, openSync, readSync, writeSync } from 'node:fs'
import { deflateRawSync } from 'node:zlib'

import { readLazyXlsxZipEntryCompressedSource, type XlsxZipByteSource, type XlsxZipEntryCompressedSource } from './zip-reader.js'

type SourcePreservingZip = Record<string, Uint8Array>

interface SourcePreservingZipEntryRecord {
  readonly pathBytes: Uint8Array
  readonly utf8: boolean
  readonly flags: number
  readonly compressionMethod: number
  readonly crc: number
  readonly compressedSize: number
  readonly uncompressedSize: number
  readonly localHeaderOffset: number
}

export interface PreparedZipEntrySizes {
  readonly compressedSize: number
  readonly uncompressedSize: number
  readonly crc: number
}

export interface PreparedZipEntry extends PreparedZipEntrySizes {
  readonly compressedChunks: readonly Uint8Array[]
}

export interface FilePreparedZipEntry extends PreparedZipEntrySizes {
  readonly compressedPath: string
}

export interface ZipDosTimeParts {
  readonly time: number
  readonly date: number
}

export interface SourcePreservingZipOptions {
  readonly dosTime?: ZipDosTimeParts
}

const zipLocalFileHeaderSignature = 0x04034b50
const zipCentralDirectoryFileHeaderSignature = 0x02014b50
const zipEndOfCentralDirectorySignature = 0x06054b50
const zipDeflateCompressionMethod = 8
const zipVersionNeeded = 20
const zipSourceCopyChunkSize = 1024 * 1024
const zipTextEncoder = new TextEncoder()

let crc32Table: Uint32Array | undefined

function writeUint16(output: Uint8Array, offset: number, value: number): void {
  output[offset] = value & 0xff
  output[offset + 1] = (value >>> 8) & 0xff
}

function writeUint32(output: Uint8Array, offset: number, value: number): void {
  output[offset] = value & 0xff
  output[offset + 1] = (value >>> 8) & 0xff
  output[offset + 2] = (value >>> 16) & 0xff
  output[offset + 3] = (value >>> 24) & 0xff
}

function currentZipDosTimeParts(): ZipDosTimeParts {
  const now = new Date()
  const year = Math.max(1980, Math.min(2099, now.getFullYear()))
  return {
    time: (now.getHours() << 11) | (now.getMinutes() << 5) | Math.trunc(now.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate(),
  }
}

function getCrc32Table(): Uint32Array {
  if (crc32Table) {
    return crc32Table
  }
  const table = new Uint32Array(256)
  for (let index = 0; index < table.length; index += 1) {
    let value = index
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
    }
    table[index] = value >>> 0
  }
  crc32Table = table
  return table
}

function crc32(bytes: Uint8Array): number {
  return crc32Finalize(crc32Update(0xffffffff, bytes))
}

export function crc32Update(crc: number, bytes: Uint8Array): number {
  const table = getCrc32Table()
  let nextCrc = crc
  for (const byte of bytes) {
    nextCrc = table[(nextCrc ^ byte) & 0xff]! ^ (nextCrc >>> 8)
  }
  return nextCrc >>> 0
}

export function crc32Finalize(crc: number): number {
  return (crc ^ 0xffffffff) >>> 0
}

function deflateZipEntry(bytes: Uint8Array): Uint8Array {
  return deflateRawSync(bytes)
}

function encodeZipPath(path: string): { readonly bytes: Uint8Array; readonly utf8: boolean } {
  const bytes = zipTextEncoder.encode(path)
  if (bytes.byteLength > 0xffff) {
    throw new Error('XLSX ZIP entry path is too long')
  }
  return { bytes, utf8: bytes.byteLength !== path.length }
}

function assertZip32Size(value: number): void {
  if (!Number.isSafeInteger(value) || value < 0 || value > 0xffffffff) {
    throw new Error('XLSX ZIP entry is too large for source-preserving export')
  }
}

function totalByteLength(chunks: readonly Uint8Array[]): number {
  let total = 0
  for (const chunk of chunks) {
    total += chunk.byteLength
  }
  return total
}

function readUnmaterializedSourceEntry(zip: SourcePreservingZip, path: string): XlsxZipEntryCompressedSource | null {
  const descriptor = Object.getOwnPropertyDescriptor(zip, path)
  if (!descriptor || typeof descriptor.get !== 'function') {
    return null
  }
  return readLazyXlsxZipEntryCompressedSource(zip, path)
}

function compressedSourceEntryChunks(entry: XlsxZipEntryCompressedSource): Uint8Array[] {
  if (entry.dataEnd - entry.dataStart !== entry.compressedSize) {
    throw new Error('Invalid XLSX source ZIP compressed data length')
  }
  const chunks: Uint8Array[] = []
  for (let offset = entry.dataStart; offset < entry.dataEnd; offset += zipSourceCopyChunkSize) {
    chunks.push(entry.source.readRange(offset, Math.min(entry.dataEnd, offset + zipSourceCopyChunkSize)))
  }
  return chunks
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

function copyCompressedSourceEntryToZipOutput(entry: XlsxZipEntryCompressedSource, pushOutput: (chunk: Uint8Array) => void): void {
  if (entry.dataEnd - entry.dataStart !== entry.compressedSize) {
    throw new Error('Invalid XLSX source ZIP compressed data length')
  }
  const scratch = entry.source.readRangeInto ? new Uint8Array(zipSourceCopyChunkSize) : undefined
  for (let offset = entry.dataStart; offset < entry.dataEnd; offset += zipSourceCopyChunkSize) {
    pushOutput(readSourceRange(entry.source, offset, Math.min(entry.dataEnd, offset + zipSourceCopyChunkSize), scratch))
  }
}

function localFileHeader(record: SourcePreservingZipEntryRecord, time: number, date: number): Uint8Array {
  const output = new Uint8Array(30 + record.pathBytes.byteLength)
  writeUint32(output, 0, zipLocalFileHeaderSignature)
  writeUint16(output, 4, zipVersionNeeded)
  writeUint16(output, 6, record.flags | (record.utf8 ? 0x0800 : 0))
  writeUint16(output, 8, record.compressionMethod)
  writeUint16(output, 10, time)
  writeUint16(output, 12, date)
  writeUint32(output, 14, record.crc)
  writeUint32(output, 18, record.compressedSize)
  writeUint32(output, 22, record.uncompressedSize)
  writeUint16(output, 26, record.pathBytes.byteLength)
  writeUint16(output, 28, 0)
  output.set(record.pathBytes, 30)
  return output
}

function centralDirectoryHeader(record: SourcePreservingZipEntryRecord, time: number, date: number): Uint8Array {
  const output = new Uint8Array(46 + record.pathBytes.byteLength)
  writeUint32(output, 0, zipCentralDirectoryFileHeaderSignature)
  writeUint16(output, 4, zipVersionNeeded)
  writeUint16(output, 6, zipVersionNeeded)
  writeUint16(output, 8, record.flags | (record.utf8 ? 0x0800 : 0))
  writeUint16(output, 10, record.compressionMethod)
  writeUint16(output, 12, time)
  writeUint16(output, 14, date)
  writeUint32(output, 16, record.crc)
  writeUint32(output, 20, record.compressedSize)
  writeUint32(output, 24, record.uncompressedSize)
  writeUint16(output, 28, record.pathBytes.byteLength)
  writeUint16(output, 30, 0)
  writeUint16(output, 32, 0)
  writeUint16(output, 34, 0)
  writeUint16(output, 36, 0)
  writeUint32(output, 38, 0)
  writeUint32(output, 42, record.localHeaderOffset)
  output.set(record.pathBytes, 46)
  return output
}

function endOfCentralDirectory(recordCount: number, centralDirectorySize: number, centralDirectoryOffset: number): Uint8Array {
  if (recordCount > 0xffff) {
    throw new Error('XLSX ZIP has too many entries for source-preserving export')
  }
  assertZip32Size(centralDirectorySize)
  assertZip32Size(centralDirectoryOffset)
  const output = new Uint8Array(22)
  writeUint32(output, 0, zipEndOfCentralDirectorySignature)
  writeUint16(output, 4, 0)
  writeUint16(output, 6, 0)
  writeUint16(output, 8, recordCount)
  writeUint16(output, 10, recordCount)
  writeUint32(output, 12, centralDirectorySize)
  writeUint32(output, 16, centralDirectoryOffset)
  writeUint16(output, 20, 0)
  return output
}

export function zipSourcePreservingEntries(
  zip: SourcePreservingZip,
  preparedEntries: ReadonlyMap<string, PreparedZipEntry> = new Map(),
  options: SourcePreservingZipOptions = {},
): Uint8Array {
  const outputChunks: Uint8Array[] = []
  let outputByteLength = 0
  const records: SourcePreservingZipEntryRecord[] = []
  const { time, date } = options.dosTime ?? currentZipDosTimeParts()
  const pushOutput = (chunk: Uint8Array): void => {
    outputChunks.push(chunk)
    outputByteLength += chunk.byteLength
  }
  const paths = [...new Set([...Object.keys(zip), ...preparedEntries.keys()])]
  for (const path of paths) {
    const preparedEntry = preparedEntries.get(path)
    const sourceEntry = preparedEntry ? null : readUnmaterializedSourceEntry(zip, path)
    const bytes = preparedEntry || sourceEntry ? undefined : zip[path]
    if (!preparedEntry && !sourceEntry && !bytes) {
      continue
    }
    assertZip32Size(outputByteLength)
    const uncompressedSize = preparedEntry?.uncompressedSize ?? sourceEntry?.uncompressedSize ?? bytes!.byteLength
    const compressedChunks =
      preparedEntry?.compressedChunks ?? (sourceEntry ? compressedSourceEntryChunks(sourceEntry) : [deflateZipEntry(bytes!)])
    const compressedSize = preparedEntry?.compressedSize ?? sourceEntry?.compressedSize ?? totalByteLength(compressedChunks)
    const crc = preparedEntry?.crc ?? sourceEntry?.crc ?? crc32(bytes!)
    assertZip32Size(uncompressedSize)
    assertZip32Size(compressedSize)
    const { bytes: pathBytes, utf8 } = encodeZipPath(path)
    const record: SourcePreservingZipEntryRecord = {
      pathBytes,
      utf8,
      flags: 0,
      compressionMethod: sourceEntry?.compressionMethod ?? zipDeflateCompressionMethod,
      crc,
      compressedSize,
      uncompressedSize,
      localHeaderOffset: outputByteLength,
    }
    pushOutput(localFileHeader(record, time, date))
    for (const compressedChunk of compressedChunks) {
      pushOutput(compressedChunk)
    }
    records.push(record)
    delete zip[path]
  }

  const centralDirectoryOffset = outputByteLength
  for (const record of records) {
    pushOutput(centralDirectoryHeader(record, time, date))
  }
  const centralDirectorySize = outputByteLength - centralDirectoryOffset
  pushOutput(endOfCentralDirectory(records.length, centralDirectorySize, centralDirectoryOffset))

  const output = new Uint8Array(outputByteLength)
  let offset = 0
  for (const chunk of outputChunks) {
    output.set(chunk, offset)
    offset += chunk.byteLength
  }
  return output
}

export function writeAllSync(fd: number, chunk: Uint8Array): void {
  let offset = 0
  while (offset < chunk.byteLength) {
    offset += writeSync(fd, chunk, offset, chunk.byteLength - offset)
  }
}

export function zipSourcePreservingEntriesToFile(
  zip: SourcePreservingZip,
  preparedEntries: ReadonlyMap<string, FilePreparedZipEntry>,
  outputPath: string,
  options: SourcePreservingZipOptions = {},
): number {
  let outputByteLength = 0
  const records: SourcePreservingZipEntryRecord[] = []
  const { time, date } = options.dosTime ?? currentZipDosTimeParts()
  const paths = [...new Set([...Object.keys(zip), ...preparedEntries.keys()])]
  const fd = openSync(outputPath, 'w')
  const pushOutput = (chunk: Uint8Array): void => {
    writeAllSync(fd, chunk)
    outputByteLength += chunk.byteLength
  }
  try {
    for (const path of paths) {
      const preparedEntry = preparedEntries.get(path)
      const sourceEntry = preparedEntry ? null : readUnmaterializedSourceEntry(zip, path)
      const bytes = preparedEntry || sourceEntry ? undefined : zip[path]
      if (!preparedEntry && !sourceEntry && !bytes) {
        continue
      }
      assertZip32Size(outputByteLength)
      const { bytes: pathBytes, utf8 } = encodeZipPath(path)
      const localHeaderOffset = outputByteLength
      if (preparedEntry) {
        assertZip32Size(preparedEntry.uncompressedSize)
        assertZip32Size(preparedEntry.compressedSize)
        const record: SourcePreservingZipEntryRecord = {
          pathBytes,
          utf8,
          flags: 0,
          compressionMethod: zipDeflateCompressionMethod,
          crc: preparedEntry.crc,
          compressedSize: preparedEntry.compressedSize,
          uncompressedSize: preparedEntry.uncompressedSize,
          localHeaderOffset,
        }
        pushOutput(localFileHeader(record, time, date))
        copyFileToZipOutput(preparedEntry.compressedPath, pushOutput)
        records.push(record)
      } else if (sourceEntry) {
        assertZip32Size(sourceEntry.uncompressedSize)
        assertZip32Size(sourceEntry.compressedSize)
        const record: SourcePreservingZipEntryRecord = {
          pathBytes,
          utf8,
          flags: 0,
          compressionMethod: sourceEntry.compressionMethod,
          crc: sourceEntry.crc,
          compressedSize: sourceEntry.compressedSize,
          uncompressedSize: sourceEntry.uncompressedSize,
          localHeaderOffset,
        }
        pushOutput(localFileHeader(record, time, date))
        copyCompressedSourceEntryToZipOutput(sourceEntry, pushOutput)
        records.push(record)
      } else {
        const uncompressedSize = bytes!.byteLength
        const compressed = deflateZipEntry(bytes!)
        const compressedSize = compressed.byteLength
        const crc = crc32(bytes!)
        assertZip32Size(uncompressedSize)
        assertZip32Size(compressedSize)
        const record: SourcePreservingZipEntryRecord = {
          pathBytes,
          utf8,
          flags: 0,
          compressionMethod: zipDeflateCompressionMethod,
          crc,
          compressedSize,
          uncompressedSize,
          localHeaderOffset,
        }
        pushOutput(localFileHeader(record, time, date))
        pushOutput(compressed)
        records.push(record)
      }
      delete zip[path]
    }

    const centralDirectoryOffset = outputByteLength
    for (const record of records) {
      pushOutput(centralDirectoryHeader(record, time, date))
    }
    const centralDirectorySize = outputByteLength - centralDirectoryOffset
    pushOutput(endOfCentralDirectory(records.length, centralDirectorySize, centralDirectoryOffset))
  } finally {
    closeSync(fd)
  }
  return outputByteLength
}

function copyFileToZipOutput(path: string, pushOutput: (chunk: Uint8Array) => void): void {
  const fd = openSync(path, 'r')
  const scratch = new Uint8Array(64 * 1024)
  try {
    let bytesRead = 0
    do {
      bytesRead = readSync(fd, scratch, 0, scratch.byteLength, null)
      if (bytesRead > 0) {
        pushOutput(scratch.subarray(0, bytesRead))
      }
    } while (bytesRead > 0)
  } finally {
    closeSync(fd)
  }
}
