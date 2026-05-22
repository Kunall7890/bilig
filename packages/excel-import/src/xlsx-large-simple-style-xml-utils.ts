const slash = 47
const greaterThan = 62
const doubleQuote = 34
const singleQuote = 39
const equals = 61
const xfSupportUnsupported = 1 << 0
const xfSupportPotentialVisual = 1 << 1

export function closingStringTagRetainLength(elementName: string): number {
  return Math.max(256, elementName.length + 4)
}

export function findNextParentBoundaryOrChild(
  xml: string,
  startIndex: number,
  parentName: string,
  childName: string,
): { readonly kind: 'child' | 'parent-close'; readonly start: number; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < xml.length) {
    const tagStart = xml.indexOf('<', index)
    if (tagStart < 0) {
      return null
    }
    if (xml.charCodeAt(tagStart + 1) === 47) {
      const tag = readStringTagName(xml, tagStart + 2)
      if (tag?.localName === parentName) {
        return { kind: 'parent-close', start: tagStart, nameEnd: tag.endIndex }
      }
      index = tagStart + 1
      continue
    }
    const tag = readStringTagName(xml, tagStart + 1)
    if (tag?.localName === childName) {
      return { kind: 'child', start: tagStart, nameEnd: tag.endIndex }
    }
    index = tagStart + 1
  }
  return null
}

export function findNextOpeningTag(
  xml: string,
  startIndex: number,
  localName: string,
): { readonly start: number; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < xml.length) {
    const tagStart = xml.indexOf('<', index)
    if (tagStart < 0) {
      return null
    }
    const tag = readStringTagName(xml, tagStart + 1)
    if (tag?.localName === localName) {
      return { start: tagStart, nameEnd: tag.endIndex }
    }
    index = tagStart + 1
  }
  return null
}

export function findClosingStringElementEnd(xml: string, startIndex: number, localName: string): number | null {
  let index = startIndex
  while (index < xml.length) {
    const tagStart = xml.indexOf('</', index)
    if (tagStart < 0) {
      return null
    }
    const tag = readStringTagName(xml, tagStart + 2)
    if (tag?.localName === localName) {
      const tagEnd = findStringTagEnd(xml, tag.endIndex)
      return tagEnd === null ? null : tagEnd + 1
    }
    index = tagStart + 2
  }
  return null
}

export function findStringTagEnd(xml: string, startIndex: number): number | null {
  let quote: string | null = null
  for (let index = startIndex; index < xml.length; index += 1) {
    const char = xml[index]
    if (quote !== null) {
      if (char === quote) {
        quote = null
      }
      continue
    }
    if (char === '"' || char === "'") {
      quote = char
      continue
    }
    if (char === '>') {
      return index
    }
  }
  return null
}

export function isSelfClosingStringTag(xml: string, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && /\s/u.test(xml[index] ?? '')) {
    index -= 1
  }
  return xml[index] === '/'
}

export function readStringTagName(xml: string, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = xml.charCodeAt(startIndex)
  if (!Number.isFinite(first) || first === 33 || first === 47 || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < xml.length && isXmlNameChar(xml[index] ?? '')) {
    if (xml[index] === ':') {
      localNameStart = index + 1
    }
    index += 1
  }
  return index === localNameStart ? null : { localName: xml.slice(localNameStart, index), endIndex: index }
}

export function isXmlNameChar(char: string): boolean {
  return /[A-Za-z0-9_.:-]/u.test(char)
}

export function findByteTagEnd(bytes: Uint8Array, startIndex: number): number | null {
  let quote: number | null = null
  for (let index = startIndex; index < bytes.byteLength; index += 1) {
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

export function isSelfClosingByteTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespaceByte(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slash
}

export function readBooleanByteRange(bytes: Uint8Array, startIndex: number, endIndex: number): boolean | undefined {
  if (endIndex - startIndex === 1) {
    if (bytes[startIndex] === 49) {
      return true
    }
    if (bytes[startIndex] === 48) {
      return false
    }
  }
  if (asciiByteRangeEqualsIgnoreCase(bytes, startIndex, endIndex, 'true')) {
    return true
  }
  if (asciiByteRangeEqualsIgnoreCase(bytes, startIndex, endIndex, 'false')) {
    return false
  }
  return undefined
}

export function inspectRequiredXfOpeningTag(bytes: Uint8Array, startIndex: number, tagEnd: number): number {
  let index = startIndex
  let fillId: number | null = null
  let fontId: number | null = null
  let borderId: number | null = null
  let applyFill: boolean | undefined
  let applyFont: boolean | undefined
  let applyBorder: boolean | undefined
  let applyAlignment: boolean | undefined
  let applyProtection: boolean | undefined
  while (index < tagEnd) {
    while (index < tagEnd && isAsciiWhitespaceByte(bytes[index] ?? 0)) {
      index += 1
    }
    const nameStart = index
    while (index < tagEnd && isXmlNameByte(bytes[index] ?? 0)) {
      index += 1
    }
    const nameEnd = index
    index = skipAsciiWhitespaceBytes(bytes, index, tagEnd)
    if (bytes[index] !== equals) {
      index += 1
      continue
    }
    index = skipAsciiWhitespaceBytes(bytes, index + 1, tagEnd)
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
    switch (nameEnd - nameStart) {
      case 6:
        if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'fillId')) {
          fillId = readNonNegativeIntegerByteRange(bytes, valueStart, valueEnd)
        } else if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'fontId')) {
          fontId = readNonNegativeIntegerByteRange(bytes, valueStart, valueEnd)
        }
        break
      case 8:
        if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'borderId')) {
          borderId = readNonNegativeIntegerByteRange(bytes, valueStart, valueEnd)
        }
        break
      case 9:
        if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'applyFill')) {
          applyFill = readBooleanByteRange(bytes, valueStart, valueEnd)
        } else if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'applyFont')) {
          applyFont = readBooleanByteRange(bytes, valueStart, valueEnd)
        }
        break
      case 11:
        if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'applyBorder')) {
          applyBorder = readBooleanByteRange(bytes, valueStart, valueEnd)
        }
        break
      case 14:
        if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'applyAlignment')) {
          applyAlignment = readBooleanByteRange(bytes, valueStart, valueEnd)
        }
        break
      case 15:
        if (asciiByteRangeEquals(bytes, nameStart, nameEnd, 'applyProtection')) {
          applyProtection = readBooleanByteRange(bytes, valueStart, valueEnd)
        }
        break
      default:
        break
    }
    index += 1
  }
  let flags = 0
  if (applyBorder ?? (borderId !== null && borderId > 0)) {
    flags |= xfSupportUnsupported
  }
  if (
    (applyFill ?? (fillId !== null && fillId > 0)) ||
    (applyFont ?? (fontId !== null && fontId > 0)) ||
    applyAlignment === true ||
    applyProtection === true
  ) {
    flags |= xfSupportPotentialVisual
  }
  return flags
}

export function readNonNegativeIntegerByteRange(bytes: Uint8Array, startIndex: number, endIndex: number): number | null {
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

export function findByteInRange(bytes: Uint8Array, startIndex: number, endIndex: number, target: number): number | null {
  for (let index = startIndex; index < endIndex; index += 1) {
    if (bytes[index] === target) {
      return index
    }
  }
  return null
}

export function skipAsciiWhitespaceBytes(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespaceByte(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

export function asciiByteRangeEquals(bytes: Uint8Array, startIndex: number, endIndex: number, value: string): boolean {
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

export function asciiByteRangeEqualsIgnoreCase(bytes: Uint8Array, startIndex: number, endIndex: number, value: string): boolean {
  if (endIndex - startIndex !== value.length) {
    return false
  }
  for (let index = 0; index < value.length; index += 1) {
    const byte = bytes[startIndex + index] ?? 0
    const normalized = byte >= 65 && byte <= 90 ? byte + 32 : byte
    if (normalized !== value.charCodeAt(index)) {
      return false
    }
  }
  return true
}

export function isAsciiWhitespaceByte(byte: number): boolean {
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

export function isApplied(tag: string, attributeName: string, componentId: number | null): boolean {
  const value = readAttribute(tag, attributeName)
  if (value === '1' || value?.toLocaleLowerCase('en-US') === 'true') {
    return true
  }
  if (value === '0' || value?.toLocaleLowerCase('en-US') === 'false') {
    return false
  }
  return componentId !== null && componentId > 0
}

export function readNonNegativeIntegerAttribute(tag: string, attributeName: string): number | null {
  const value = readAttribute(tag, attributeName)
  if (!value || !/^[0-9]+$/u.test(value)) {
    return null
  }
  const number = Number(value)
  return Number.isSafeInteger(number) ? number : null
}

export function readNumberAttribute(tag: string, attributeName: string): number | null {
  const value = readAttribute(tag, attributeName)
  if (value === undefined || value.trim().length === 0) {
    return null
  }
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

export function readBooleanAttribute(value: string | undefined): boolean | undefined {
  if (value === '1' || value?.toLocaleLowerCase('en-US') === 'true') {
    return true
  }
  if (value === '0' || value?.toLocaleLowerCase('en-US') === 'false') {
    return false
  }
  return undefined
}

export function readAttribute(tag: string, attributeName: string): string | undefined {
  const pattern = new RegExp(`\\s${escapeRegExp(attributeName)}=(?:"([^"]*)"|'([^']*)')`, 'u')
  const match = pattern.exec(tag)
  return match?.[1] ?? match?.[2]
}

export function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
