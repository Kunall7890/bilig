import { readKnownXmlLocalName } from './xlsx-large-simple-xml-name.js'

export const lessThan = 60
export const slash = 47
const greaterThan = 62
export const doubleQuote = 34
export const singleQuote = 39
export const packedAddressColumnFactor = 16_384
export const headlessTagUnknown = 0
export const headlessTagCell = 1
export const headlessTagRow = 2
export const headlessTagDimension = 3

export function sheetMetadataKeysForHeadlessElement(localName: string): readonly string[] {
  switch (localName) {
    case 'autoFilter':
      return ['filters']
    case 'colBreaks':
    case 'headerFooter':
    case 'pageMargins':
    case 'pageSetup':
    case 'printOptions':
    case 'rowBreaks':
      return ['printPageSetup']
    case 'cols':
      return ['columnMetadata', 'columns']
    case 'conditionalFormatting':
      return ['conditionalFormats']
    case 'controls':
    case 'legacyDrawing':
    case 'oleObjects':
      return ['controlArtifacts']
    case 'drawing':
      return ['drawingArtifacts']
    case 'hyperlinks':
      return ['hyperlinks']
    case 'mergeCells':
      return ['merges']
    case 'sheetFormatPr':
      return ['sheetFormatPr']
    default:
      return []
  }
}

export function headlessTagLocalName(code: number): string | null {
  switch (code) {
    case headlessTagCell:
      return 'c'
    case headlessTagRow:
      return 'row'
    case headlessTagDimension:
      return 'dimension'
    default:
      return null
  }
}

export function matchesDimensionTagName(bytes: Uint8Array, startIndex: number): boolean {
  return (
    bytes[startIndex + 1] === 105 &&
    bytes[startIndex + 2] === 109 &&
    bytes[startIndex + 3] === 101 &&
    bytes[startIndex + 4] === 110 &&
    bytes[startIndex + 5] === 115 &&
    bytes[startIndex + 6] === 105 &&
    bytes[startIndex + 7] === 111 &&
    bytes[startIndex + 8] === 110 &&
    !isXmlNameByte(bytes[startIndex + 9] ?? 0)
  )
}

export function metadataChildName(localName: string): string | null {
  switch (localName) {
    case 'conditionalFormatting':
      return 'cfRule'
    case 'mergeCells':
      return 'mergeCell'
    case 'tableParts':
      return 'tablePart'
    default:
      return null
  }
}

export function metadataCountMultiplier(localName: string, bytes: Uint8Array, nameEnd: number, tagEnd: number): number {
  return localName === 'conditionalFormatting' ? countSqrefRangesFromTag(bytes, nameEnd, tagEnd) : 1
}

export function cellContentHasRichTextRun(bytes: Uint8Array, startIndex: number, endIndex: number): boolean {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    if (bytes[index + 1] === slash) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === 'r') {
      return true
    }
    index = tag?.endIndex ?? index + 1
  }
  return false
}

export function findTagEnd(bytes: Uint8Array, startIndex: number, endIndex: number = bytes.byteLength): number | null {
  let quote: number | null = null
  for (let index = startIndex; index < endIndex; index += 1) {
    const byte = bytes[index] ?? 0
    if (quote !== null) {
      if (byte === quote) {
        quote = null
      }
      continue
    }
    if (byte === doubleQuote || byte === singleQuote) {
      quote = byte
      continue
    }
    if (byte === greaterThan) {
      return index
    }
  }
  return null
}

export function findClosingTag(
  bytes: Uint8Array,
  startIndex: number,
  localName: string,
  endIndex: number = bytes.byteLength,
): { readonly start: number; readonly end: number } | null {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan || bytes[index + 1] !== slash) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 2)
    if (tag?.localName === localName) {
      const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
      return tagEnd === null ? null : { start: index, end: tagEnd + 1 }
    }
    index += 1
  }
  return null
}

export function isSelfClosingTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slash
}

export function readDimensionAddressRange(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): { readonly startRow: number; readonly startColumn: number; readonly endRow: number; readonly endColumn: number } | null {
  const trimmed = trimAsciiWhitespace(bytes, startIndex, endIndex)
  if (trimmed.start === trimmed.end) {
    return null
  }
  const separator = findByte(bytes, trimmed.start, trimmed.end, 58)
  const firstEnd = separator ?? trimmed.end
  const start = decodePackedCellAddressBytes(bytes, trimmed.start, firstEnd)
  const end = separator === null ? start : decodePackedCellAddressBytes(bytes, separator + 1, trimmed.end)
  return start === null || end === null
    ? null
    : {
        startRow: packedAddressRow(start),
        startColumn: packedAddressColumn(start),
        endRow: packedAddressRow(end),
        endColumn: packedAddressColumn(end),
      }
}

export function countSqrefRangesFromTag(bytes: Uint8Array, nameEnd: number, tagEnd: number): number {
  const sqref = readXmlAttributeRangeFromTag(bytes, nameEnd, tagEnd, 'sqref')
  if (!sqref) {
    return 1
  }
  let count = 0
  let inToken = false
  for (let index = sqref.start; index < sqref.end; index += 1) {
    if (isAsciiWhitespace(bytes[index] ?? 0)) {
      inToken = false
      continue
    }
    if (!inToken) {
      count += 1
      inToken = true
    }
  }
  return Math.max(1, count)
}

export function readPositiveIntegerAttributeFromTag(
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  attributeName: string,
): number | null {
  const range = readXmlAttributeRangeFromTag(bytes, nameEnd, tagEnd, attributeName)
  if (!range || range.start === range.end) {
    return null
  }
  const value = readNonNegativeIntegerFromRange(bytes, range.start, range.end)
  return value !== null && value > 0 ? value : null
}

export function readNonNegativeIntegerFromRange(bytes: Uint8Array, startIndex: number, endIndex: number): number | null {
  if (startIndex === endIndex) {
    return null
  }
  let value = 0
  for (let index = startIndex; index < endIndex; index += 1) {
    const byte = bytes[index] ?? 0
    if (byte < 48 || byte > 57) {
      return null
    }
    value = value * 10 + byte - 48
  }
  return Number.isSafeInteger(value) ? value : null
}

export function readXmlTagName(bytes: Uint8Array, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = bytes[startIndex]
  if (first === undefined || first === 33 || first === slash || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < bytes.byteLength && isXmlNameByte(bytes[index] ?? 0)) {
    if (bytes[index] === 58) {
      localNameStart = index + 1
    }
    index += 1
  }
  return index === localNameStart
    ? null
    : { localName: readKnownXmlLocalName(bytes, localNameStart, index) ?? decodeAscii(bytes, localNameStart, index), endIndex: index }
}

export function readXmlAttributeRangeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
  attributeName: string,
): { readonly start: number; readonly end: number } | null {
  let index = startIndex
  while (index < tagEnd) {
    while (index < tagEnd && isAsciiWhitespace(bytes[index] ?? 0)) {
      index += 1
    }
    const nameStart = index
    while (index < tagEnd && isXmlNameByte(bytes[index] ?? 0)) {
      index += 1
    }
    const nameEnd = index
    index = skipAsciiWhitespace(bytes, index, tagEnd)
    if (bytes[index] !== 61) {
      index += 1
      continue
    }
    index = skipAsciiWhitespace(bytes, index + 1, tagEnd)
    const quote = bytes[index]
    if (quote !== doubleQuote && quote !== singleQuote) {
      index += 1
      continue
    }
    const valueStart = index + 1
    index = valueStart
    while (index < tagEnd && bytes[index] !== quote) {
      index += 1
    }
    const valueEnd = index
    if (attributeNameMatches(bytes, nameStart, nameEnd, attributeName)) {
      return { start: valueStart, end: valueEnd }
    }
    index += 1
  }
  return null
}

export function attributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, attributeName: string): boolean {
  return attributeValueMatches(bytes, startIndex, endIndex, attributeName)
}

export function attributeValueMatches(bytes: Uint8Array, startIndex: number, endIndex: number, value: string): boolean {
  if (endIndex - startIndex !== value.length) {
    return false
  }
  for (let index = 0; index < value.length; index += 1) {
    if (bytes[startIndex + index] !== value.charCodeAt(index)) {
      return false
    }
  }
  return true
}

export function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

export function trimAsciiWhitespace(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): { readonly start: number; readonly end: number } {
  let start = startIndex
  let end = endIndex
  while (start < end && isAsciiWhitespace(bytes[start] ?? 0)) {
    start += 1
  }
  while (end > start && isAsciiWhitespace(bytes[end - 1] ?? 0)) {
    end -= 1
  }
  return { start, end }
}

export function findByte(bytes: Uint8Array, startIndex: number, endIndex: number, target: number): number | null {
  for (let index = startIndex; index < endIndex; index += 1) {
    if (bytes[index] === target) {
      return index
    }
  }
  return null
}

export function decodePackedCellAddressBytes(bytes: Uint8Array, startIndex: number, endIndex: number): number | null {
  let column = 0
  let row = 0
  let letterCount = 0
  let digitCount = 0
  for (let index = startIndex; index < endIndex; index += 1) {
    const byte = bytes[index] ?? 0
    if (byte === 36) {
      continue
    }
    const upper = byte >= 97 && byte <= 122 ? byte - 32 : byte
    if (upper >= 65 && upper <= 90 && digitCount === 0) {
      column = column * 26 + upper - 64
      letterCount += 1
      continue
    }
    if (byte >= 48 && byte <= 57 && letterCount > 0) {
      row = row * 10 + byte - 48
      digitCount += 1
      continue
    }
    return null
  }
  return letterCount > 0 && letterCount <= 3 && digitCount > 0 && row > 0 && column > 0 ? packCellAddress(row - 1, column - 1) : null
}

export function packCellAddress(row: number, column: number): number {
  return row * packedAddressColumnFactor + column
}

export function packedAddressRow(value: number): number {
  return Math.floor(value / packedAddressColumnFactor)
}

export function packedAddressColumn(value: number): number {
  return value % packedAddressColumnFactor
}

export function decodeAscii(bytes: Uint8Array, startIndex: number, endIndex: number): string {
  let output = ''
  for (let index = startIndex; index < endIndex; index += 1) {
    output += String.fromCharCode(bytes[index] ?? 0)
  }
  return output
}

export function isAsciiWhitespace(byte: number): boolean {
  return byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32
}

export function isXmlNameByte(byte: number): boolean {
  return (
    (byte >= 65 && byte <= 90) ||
    (byte >= 97 && byte <= 122) ||
    (byte >= 48 && byte <= 57) ||
    byte === 45 ||
    byte === 46 ||
    byte === 58 ||
    byte === 95
  )
}
