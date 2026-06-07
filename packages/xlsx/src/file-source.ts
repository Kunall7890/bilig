import { closeSync, openSync, readFileSync, readSync, statSync } from 'node:fs'

import type { XlsxSourceReader } from './source-preserving-literal-patches.js'
import type { XlsxZipByteSource } from './zip-reader.js'

export type FileXlsxSourceReader = XlsxSourceReader & XlsxZipByteSource & { readonly path: string }

export interface FileXlsxSourceReaderOptions {
  readonly maxReadBytes?: number | false
}

export const defaultFileXlsxSourceReadBytesLimit = 1_000_000

export function createFileXlsxSourceReader(path: string, options: FileXlsxSourceReaderOptions = {}): FileXlsxSourceReader {
  const fd = openSync(path, 'r')
  const byteLength = statSync(path).size
  const maxReadBytes = options.maxReadBytes ?? defaultFileXlsxSourceReadBytesLimit
  let closed = false

  const assertOpen = (): void => {
    if (closed) {
      throw new Error('XLSX file source has been released')
    }
  }
  const assertRange = (start: number, end: number): void => {
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > byteLength) {
      throw new Error(`Invalid XLSX file byte range: ${String(start)}..${String(end)}`)
    }
  }
  const assertReadBytesWithinLimit = (): void => {
    if (maxReadBytes === false || byteLength <= maxReadBytes) {
      return
    }
    throw new Error(
      [
        `XLSX file source readBytes is small-workbook only: ${path} is ${String(byteLength)} bytes`,
        `limit is ${String(maxReadBytes)} bytes`,
        'Use readRange/readRangeInto or a file-backed native XLSX API for large workbooks.',
      ].join('; '),
    )
  }

  return {
    byteLength,
    path,
    readBytes() {
      assertOpen()
      assertReadBytesWithinLimit()
      return readFileSync(path)
    },
    readRange(start, end) {
      assertOpen()
      assertRange(start, end)
      const length = end - start
      const buffer = new Uint8Array(length)
      let offset = 0
      while (offset < length) {
        const bytesRead = readSync(fd, buffer, offset, length - offset, start + offset)
        if (bytesRead === 0) {
          throw new Error('Unexpected EOF while reading XLSX file source')
        }
        offset += bytesRead
      }
      return buffer
    },
    readRangeInto(start, end, target) {
      assertOpen()
      assertRange(start, end)
      const length = end - start
      if (target.byteLength < length) {
        throw new Error('XLSX file source scratch buffer is too small')
      }
      let offset = 0
      while (offset < length) {
        const bytesRead = readSync(fd, target, offset, length - offset, start + offset)
        if (bytesRead === 0) {
          throw new Error('Unexpected EOF while reading XLSX file source')
        }
        offset += bytesRead
      }
      return target.subarray(0, length)
    },
    release() {
      if (!closed) {
        closeSync(fd)
        closed = true
      }
    },
  }
}
