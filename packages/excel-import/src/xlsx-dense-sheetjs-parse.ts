import type { Unzipped } from 'fflate'
import { decodeCellRange } from '@bilig/xlsx'

const textDecoder = new TextDecoder()

export function shouldUseDenseSheetJsParse(
  data: Uint8Array,
  workbookZip: Unzipped | null,
  options: {
    readonly minByteLength: number
    readonly maxColumnCount: number
  },
): boolean {
  if (!workbookZip || data.byteLength < options.minByteLength) {
    return false
  }
  let sawValidWorksheetDimension = false
  for (const path of Object.keys(workbookZip)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/u.test(path)) {
      continue
    }
    const bytes = workbookZip[path]
    if (!bytes) {
      continue
    }
    const dimensionRef = readWorksheetDimensionRef(bytes)
    if (!dimensionRef) {
      continue
    }
    const range = decodeWorksheetDimensionRef(dimensionRef)
    if (!range) {
      continue
    }
    sawValidWorksheetDimension = true
    if (range.e.c + 1 > options.maxColumnCount) {
      return false
    }
  }
  return sawValidWorksheetDimension
}

function readWorksheetDimensionRef(bytes: Uint8Array): string | null {
  const headerXml = textDecoder.decode(bytes.subarray(0, Math.min(bytes.byteLength, 4096)))
  return /<dimension\b[^>]*\bref="([^"]+)"/u.exec(headerXml)?.[1] ?? null
}

function decodeWorksheetDimensionRef(ref: string): ReturnType<typeof decodeCellRange> | null {
  try {
    return decodeCellRange(ref)
  } catch {
    return null
  }
}
