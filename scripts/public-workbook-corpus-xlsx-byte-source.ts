import { createHash } from 'node:crypto'
import { closeSync, createReadStream, fstatSync, openSync, readSync } from 'node:fs'
import { createInflateRaw, inflateRawSync } from 'node:zlib'

import type { XlsxZipByteSource, XlsxZipChunkConsumer } from '../packages/excel-import/src/xlsx-zip.js'

export class FileBackedXlsxZipByteSource implements XlsxZipByteSource {
  readonly byteLength: number
  private fd: number | null

  constructor(private readonly path: string) {
    this.fd = openSync(path, 'r')
    this.byteLength = fstatSync(this.fd).size
  }

  readRange(start: number, end: number): Uint8Array {
    if (this.fd === null) {
      throw new Error('XLSX ZIP file source has been released')
    }
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > this.byteLength) {
      throw new Error('Invalid XLSX ZIP file byte range')
    }
    const output = Buffer.allocUnsafe(end - start)
    let offset = 0
    while (offset < output.byteLength) {
      const bytesRead = readSync(this.fd, output, offset, output.byteLength - offset, start + offset)
      if (bytesRead === 0) {
        throw new Error('Unexpected end of XLSX ZIP file source')
      }
      offset += bytesRead
    }
    return output
  }

  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array {
    if (this.fd === null) {
      throw new Error('XLSX ZIP file source has been released')
    }
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > this.byteLength) {
      throw new Error('Invalid XLSX ZIP file byte range')
    }
    const length = end - start
    if (length > target.byteLength) {
      throw new Error('XLSX ZIP read target is too small')
    }
    let offset = 0
    while (offset < length) {
      const bytesRead = readSync(this.fd, target, offset, length - offset, start + offset)
      if (bytesRead === 0) {
        throw new Error('Unexpected end of XLSX ZIP file source')
      }
      offset += bytesRead
    }
    return target.subarray(0, length)
  }

  inflateRawRange(start: number, end: number): Uint8Array {
    return inflateRawSync(this.readRange(start, end))
  }

  async inflateRawRangeChunksAsync(
    start: number,
    end: number,
    onChunk: XlsxZipChunkConsumer,
    options: { readonly chunkSize: number },
  ): Promise<boolean> {
    if (this.fd === null) {
      throw new Error('XLSX ZIP file source has been released')
    }
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > this.byteLength) {
      throw new Error('Invalid XLSX ZIP file byte range')
    }
    const stream =
      start === end
        ? null
        : createReadStream(this.path, {
            end: end - 1,
            highWaterMark: Math.max(1, Math.trunc(options.chunkSize)),
            start,
          })
    const inflate = createInflateRaw()
    try {
      if (!stream) {
        inflate.end()
      } else {
        stream.pipe(inflate)
      }
      for await (const chunk of inflate) {
        const bytes = chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk)
        if (!emitInflatedChunks(bytes, options.chunkSize, onChunk)) {
          stream?.destroy()
          inflate.destroy()
          return true
        }
      }
      return true
    } finally {
      stream?.destroy()
      inflate.destroy()
    }
  }

  release(): void {
    if (this.fd === null) {
      return
    }
    closeSync(this.fd)
    this.fd = null
  }
}

function emitInflatedChunks(chunk: Uint8Array, chunkSize: number, onChunk: XlsxZipChunkConsumer): boolean {
  const normalizedChunkSize = Math.max(1, Math.trunc(chunkSize))
  for (let offset = 0; offset < chunk.byteLength; offset += normalizedChunkSize) {
    if (onChunk(chunk.subarray(offset, Math.min(chunk.byteLength, offset + normalizedChunkSize))) === false) {
      return false
    }
  }
  return true
}

export function isZipWorkbookSource(source: XlsxZipByteSource): boolean {
  const magicByteLength = Math.min(4, source.byteLength)
  const scratch = source.readRangeInto ? new Uint8Array(4) : undefined
  const bytes = readXlsxZipByteSourceRange(source, 0, magicByteLength, scratch)
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04
}

export function sha256XlsxZipByteSourceHex(source: XlsxZipByteSource): string {
  const hash = createHash('sha256')
  const chunkSize = 64 * 1024
  const scratch = source.readRangeInto ? new Uint8Array(chunkSize) : undefined
  for (let offset = 0; offset < source.byteLength; offset += chunkSize) {
    const end = Math.min(source.byteLength, offset + chunkSize)
    hash.update(readXlsxZipByteSourceRange(source, offset, end, scratch))
  }
  return hash.digest('hex')
}

function readXlsxZipByteSourceRange(source: XlsxZipByteSource, start: number, end: number, scratch: Uint8Array | undefined): Uint8Array {
  const length = end - start
  if (!scratch || !source.readRangeInto || length > scratch.byteLength) {
    return source.readRange(start, end)
  }
  const chunk = source.readRangeInto(start, end, scratch)
  if (chunk.byteLength !== length) {
    throw new Error('XLSX ZIP byte source returned an invalid hash chunk length')
  }
  return chunk
}
