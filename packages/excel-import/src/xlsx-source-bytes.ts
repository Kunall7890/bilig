import type { WorkbookSnapshot } from '@bilig/protocol'
import { randomUUID } from 'node:crypto'
import { closeSync, existsSync, mkdirSync, openSync, readFileSync, readSync, unlinkSync, writeSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { XlsxZipByteSource } from './xlsx-zip.js'

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')
const tempSourceDirectoryName = 'bilig-imported-xlsx-sources'
const tempSourceWriteChunkSize = 1024 * 1024

export interface ImportedXlsxSourceReader {
  readonly byteLength: number
  readBytes(): Uint8Array
  release?(): void
}

type ImportedXlsxSourceReference = Uint8Array | ImportedXlsxSourceReader

type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: ImportedXlsxSourceReference
}

type MutableSnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  [importedXlsxSourceBytes]?: ImportedXlsxSourceReference
}

export function attachImportedXlsxSourceBytes(snapshot: WorkbookSnapshot, bytes: Uint8Array): WorkbookSnapshot {
  return attachImportedXlsxSourceReference(snapshot, bytes)
}

export function attachImportedXlsxSourceReader(snapshot: WorkbookSnapshot, source: ImportedXlsxSourceReader): WorkbookSnapshot {
  return attachImportedXlsxSourceReference(snapshot, source)
}

function attachImportedXlsxSourceReference(snapshot: WorkbookSnapshot, source: ImportedXlsxSourceReference): WorkbookSnapshot {
  Object.defineProperty(snapshot, importedXlsxSourceBytes, {
    configurable: true,
    enumerable: false,
    value: source,
  })
  return snapshot
}

export function readImportedXlsxSourceBytes(snapshot: WorkbookSnapshot): Uint8Array | undefined {
  const source = (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
  return source instanceof Uint8Array ? source : source?.readBytes()
}

export function detachImportedXlsxSourceBytes(snapshot: WorkbookSnapshot): boolean {
  const source = (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
  if (source === undefined || !Object.prototype.hasOwnProperty.call(snapshot, importedXlsxSourceBytes)) {
    return false
  }
  if (!(source instanceof Uint8Array)) {
    source.release?.()
  }
  delete (snapshot as MutableSnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
  return true
}

export function createTempFileImportedXlsxSourceReader(bytes: Uint8Array): ImportedXlsxSourceReader & XlsxZipByteSource {
  const directory = join(tmpdir(), tempSourceDirectoryName)
  mkdirSync(directory, { recursive: true })
  const path = join(directory, `${randomUUID()}.xlsx`)
  writeTempSourceBytes(path, bytes)
  const reader = new TempFileImportedXlsxSourceReader(path, bytes.byteLength)
  tempFileSourceFinalizer?.register(reader, path, reader)
  return reader
}

function writeTempSourceBytes(path: string, bytes: Uint8Array): void {
  const fd = openSync(path, 'w')
  try {
    for (let offset = 0; offset < bytes.byteLength; offset += tempSourceWriteChunkSize) {
      const chunk = bytes.subarray(offset, Math.min(bytes.byteLength, offset + tempSourceWriteChunkSize))
      let chunkOffset = 0
      while (chunkOffset < chunk.byteLength) {
        chunkOffset += writeSync(fd, chunk, chunkOffset, chunk.byteLength - chunkOffset)
      }
    }
  } finally {
    closeSync(fd)
  }
}

class TempFileImportedXlsxSourceReader implements ImportedXlsxSourceReader, XlsxZipByteSource {
  private released = false

  constructor(
    private readonly path: string,
    readonly byteLength: number,
  ) {}

  readBytes(): Uint8Array {
    if (this.released || !existsSync(this.path)) {
      return new Uint8Array(0)
    }
    const bytes = readFileSync(this.path)
    return bytes.byteLength === this.byteLength ? bytes : new Uint8Array(bytes)
  }

  readRange(start: number, end: number): Uint8Array {
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > this.byteLength) {
      throw new Error('Invalid imported XLSX temp source byte range')
    }
    const output = Buffer.allocUnsafe(end - start)
    const fd = openSync(this.path, 'r')
    try {
      let offset = 0
      while (offset < output.byteLength) {
        const bytesRead = readSync(fd, output, offset, output.byteLength - offset, start + offset)
        if (bytesRead === 0) {
          throw new Error('Unexpected end of imported XLSX temp source')
        }
        offset += bytesRead
      }
      return output
    } finally {
      closeSync(fd)
    }
  }

  readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array {
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > this.byteLength) {
      throw new Error('Invalid imported XLSX temp source byte range')
    }
    const length = end - start
    if (length > target.byteLength) {
      throw new Error('Imported XLSX temp source read target is too small')
    }
    const fd = openSync(this.path, 'r')
    try {
      let offset = 0
      while (offset < length) {
        const bytesRead = readSync(fd, target, offset, length - offset, start + offset)
        if (bytesRead === 0) {
          throw new Error('Unexpected end of imported XLSX temp source')
        }
        offset += bytesRead
      }
    } finally {
      closeSync(fd)
    }
    return target.subarray(0, length)
  }

  release(): void {
    if (this.released) {
      return
    }
    this.released = true
    tempFileSourceFinalizer?.unregister(this)
    try {
      unlinkSync(this.path)
    } catch {
      // The temp source is best-effort cleanup. Export still works if the file was already removed.
    }
  }
}

const tempFileSourceFinalizer =
  typeof FinalizationRegistry === 'undefined'
    ? undefined
    : new FinalizationRegistry<string>((path) => {
        try {
          unlinkSync(path)
        } catch {
          // Temp-file cleanup is best effort.
        }
      })
