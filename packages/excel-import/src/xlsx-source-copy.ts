import { closeSync, openSync, writeFileSync, writeSync } from 'node:fs'

import type { ImportedXlsxSourceReference } from './xlsx-source-bytes.js'
import type { XlsxSourceLiteralPatchFileExportResult } from './xlsx-source-preserving-export.js'

const importedSourceCopyChunkSize = 1024 * 1024

type ImportedXlsxSourceRangeReader = Exclude<ImportedXlsxSourceReference, Uint8Array> & {
  readRange(start: number, end: number): Uint8Array
}

export function tryCopyImportedXlsxSourceToFile(
  source: ImportedXlsxSourceReference,
  outputPath: string,
): XlsxSourceLiteralPatchFileExportResult | null {
  if (source instanceof Uint8Array) {
    writeFileSync(outputPath, source)
    return { bytesWritten: source.byteLength }
  }
  if (!hasImportedXlsxSourceReadRange(source)) {
    return null
  }
  const fd = openSync(outputPath, 'w')
  let bytesWritten = 0
  try {
    for (let offset = 0; offset < source.byteLength; offset += importedSourceCopyChunkSize) {
      const end = Math.min(source.byteLength, offset + importedSourceCopyChunkSize)
      const chunk = source.readRange(offset, end)
      let chunkOffset = 0
      while (chunkOffset < chunk.byteLength) {
        chunkOffset += writeSync(fd, chunk, chunkOffset, chunk.byteLength - chunkOffset)
      }
      bytesWritten += chunk.byteLength
    }
  } finally {
    closeSync(fd)
  }
  return { bytesWritten }
}

function hasImportedXlsxSourceReadRange(source: Exclude<ImportedXlsxSourceReference, Uint8Array>): source is ImportedXlsxSourceRangeReader {
  return typeof Reflect.get(source, 'readRange') === 'function'
}
