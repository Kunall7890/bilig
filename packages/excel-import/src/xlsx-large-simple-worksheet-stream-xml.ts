import type { LargeSimpleXmlTextRange } from './xlsx-large-simple-cell-value-scan.js'
import {
  decodeAscii,
  decodeBytes,
  decodePackedCellAddressBytes,
  isAsciiWhitespace,
  isXmlNameByte,
} from './xlsx-large-simple-xml-byte-utils.js'
import { readKnownXmlLocalName } from './xlsx-large-simple-xml-name.js'

const lessThan = 60
const slash = 47
const greaterThan = 62
const doubleQuote = 34
const singleQuote = 39

export function findNextOpeningTag(
  bytes: Uint8Array,
  startIndex: number,
  localName: string,
  endIndex: number = bytes.byteLength,
): { readonly start: number; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === localName) {
      return { start: index, nameEnd: tag.endIndex }
    }
    index += 1
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

export function readElementTextRange(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
  elementName: string,
): LargeSimpleXmlTextRange | null {
  const tag = findNextOpeningTag(bytes, startIndex, elementName, endIndex)
  if (!tag) {
    return null
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, endIndex)
  if (tagEnd === null || isSelfClosingTag(bytes, tagEnd)) {
    return null
  }
  const closing = findClosingTag(bytes, tagEnd + 1, elementName, endIndex)
  return closing ? { start: tagEnd + 1, end: closing.start } : null
}

export function readElementXml(bytes: Uint8Array, startIndex: number, endIndex: number, elementName: string): string | null {
  const tag = findNextOpeningTag(bytes, startIndex, elementName, endIndex)
  if (!tag) {
    return null
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, endIndex)
  if (tagEnd === null) {
    return null
  }
  if (isSelfClosingTag(bytes, tagEnd)) {
    return decodeBytes(bytes, tag.start, tagEnd + 1)
  }
  const closing = findClosingTag(bytes, tagEnd + 1, elementName, endIndex)
  return closing ? decodeBytes(bytes, tag.start, closing.end) : null
}

export function hasElement(bytes: Uint8Array, startIndex: number, endIndex: number, elementName: string): boolean {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === elementName) {
      return true
    }
    index += 1
  }
  return false
}

export function countOpeningTags(bytes: Uint8Array, startIndex: number, endIndex: number, localName: string): number {
  let count = 0
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === localName) {
      count += 1
      index = tag.endIndex
      continue
    }
    index += 1
  }
  return count
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

export function readXmlAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeBytes(bytes, range.start, range.end) : null
}

export function readPackedCellAddressAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number): number | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, 'r')
  return range ? decodePackedCellAddressBytes(bytes, range.start, range.end) : null
}

export function readCellStyleIndexFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number): number | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, 's')
  if (!range) {
    return null
  }
  let value = 0
  if (range.start === range.end) {
    return null
  }
  for (let index = range.start; index < range.end; index += 1) {
    const byte = bytes[index] ?? 0
    if (byte < 48 || byte > 57) {
      return null
    }
    value = value * 10 + byte - 48
    if (!Number.isSafeInteger(value)) {
      return null
    }
  }
  return value
}

export function isSelfClosingTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slash
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

function attributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, attributeName: string): boolean {
  if (endIndex - startIndex !== attributeName.length) {
    return false
  }
  for (let index = 0; index < attributeName.length; index += 1) {
    if (bytes[startIndex + index] !== attributeName.charCodeAt(index)) {
      return false
    }
  }
  return true
}

function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}
