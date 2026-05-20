import { TextDecoder } from 'node:util'
import { inflateRawSync } from 'node:zlib'

export interface ZipEntryInfo {
  readonly path: string
  readonly compressionMethod: number
  readonly compressedSize: number
  readonly localHeaderOffset: number
}

export interface WorkbookPackageEntry {
  readonly path: string
  readonly compressionMethod: number
  readonly compressedSize: number
}

const decoder = new TextDecoder()
const eocdSignature = 0x06054b50
const centralDirectorySignature = 0x02014b50
const localFileHeaderSignature = 0x04034b50

export function isZipWorkbook(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04
}

export function readZipCentralDirectory(bytes: Uint8Array): ZipEntryInfo[] {
  const eocdOffset = findEndOfCentralDirectory(bytes)
  const centralDirectorySize = readUint32(bytes, eocdOffset + 12)
  const centralDirectoryOffset = readUint32(bytes, eocdOffset + 16)
  if (centralDirectorySize === 0xffffffff || centralDirectoryOffset === 0xffffffff) {
    throw new Error('Zip64 central directories are not supported by the low-memory XLSX footprint scanner')
  }
  const entries: ZipEntryInfo[] = []
  let offset = centralDirectoryOffset
  const endOffset = centralDirectoryOffset + centralDirectorySize
  while (offset < endOffset) {
    if (readUint32(bytes, offset) !== centralDirectorySignature) {
      throw new Error('Invalid XLSX central directory entry')
    }
    const compressionMethod = readUint16(bytes, offset + 10)
    const compressedSize = readUint32(bytes, offset + 20)
    const fileNameLength = readUint16(bytes, offset + 28)
    const extraFieldLength = readUint16(bytes, offset + 30)
    const fileCommentLength = readUint16(bytes, offset + 32)
    const localHeaderOffset = readUint32(bytes, offset + 42)
    const path = normalizeZipPath(decodeBytes(bytes.subarray(offset + 46, offset + 46 + fileNameLength)))
    entries.push({ path, compressionMethod, compressedSize, localHeaderOffset })
    offset += 46 + fileNameLength + extraFieldLength + fileCommentLength
  }
  return entries
}

export function readZipEntryBytes(bytes: Uint8Array, entry: ZipEntryInfo | undefined): Uint8Array | null {
  const payload = readZipEntryPayload(bytes, entry)
  if (!payload) {
    return null
  }
  if (payload.compressionMethod === 0) {
    return payload.compressed
  }
  if (payload.compressionMethod === 8) {
    return inflateRawSync(payload.compressed)
  }
  throw new Error(`Unsupported XLSX zip compression method ${String(payload.compressionMethod)} for ${payload.path}`)
}

export function readZipEntryPayload(
  bytes: Uint8Array,
  entry: ZipEntryInfo | undefined,
): { readonly compressed: Uint8Array; readonly compressionMethod: number; readonly path: string } | null {
  if (!entry) {
    return null
  }
  const offset = entry.localHeaderOffset
  if (readUint32(bytes, offset) !== localFileHeaderSignature) {
    throw new Error(`Invalid local file header for ${entry.path}`)
  }
  const fileNameLength = readUint16(bytes, offset + 26)
  const extraFieldLength = readUint16(bytes, offset + 28)
  const dataOffset = offset + 30 + fileNameLength + extraFieldLength
  const compressed = bytes.subarray(dataOffset, dataOffset + entry.compressedSize)
  return { compressed, compressionMethod: entry.compressionMethod, path: entry.path }
}

export function readZipEntryText(bytes: Uint8Array, entry: ZipEntryInfo | undefined): string | null {
  const entryBytes = readZipEntryBytes(bytes, entry)
  return entryBytes ? decodeBytes(entryBytes) : null
}

export function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

export function decodeBytes(bytes: Uint8Array): string {
  return decoder.decode(bytes)
}

function findEndOfCentralDirectory(bytes: Uint8Array): number {
  const minOffset = Math.max(0, bytes.length - 65_557)
  for (let offset = bytes.length - 22; offset >= minOffset; offset -= 1) {
    if (readUint32(bytes, offset) === eocdSignature) {
      return offset
    }
  }
  throw new Error('Invalid XLSX zip: end of central directory not found')
}

function readUint16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8)
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0
}
