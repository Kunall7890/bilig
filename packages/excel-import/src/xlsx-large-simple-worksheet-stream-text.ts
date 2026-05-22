import { decodeExcelEscapedText } from './xlsx-escaped-text.js'

export function decodeXmlText(value: string): string {
  if (!value.includes('&')) {
    return value
  }
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

export function stringItemText(xml: string): string {
  return normalizeWorksheetText(
    [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
      .map((match) => decodeXmlText(match[1] ?? ''))
      .join(''),
  )
}

export function normalizeWorksheetText(value: string): string {
  const hasCarriageReturn = value.includes('\r')
  const hasExcelEscape = maybeContainsExcelEscapedText(value)
  if (!hasCarriageReturn && !hasExcelEscape) {
    return value
  }
  const decoded = hasExcelEscape ? decodeExcelEscapedText(value) : value
  return hasCarriageReturn ? decoded.replace(/\r\n?/gu, '\n') : decoded
}

function maybeContainsExcelEscapedText(value: string): boolean {
  let index = value.indexOf('_x')
  while (index !== -1) {
    if (index + 6 < value.length && value[index + 6] === '_' && isFourHexDigits(value, index + 2)) {
      return true
    }
    index = value.indexOf('_x', index + 2)
  }
  index = value.indexOf('_X')
  while (index !== -1) {
    if (index + 6 < value.length && value[index + 6] === '_' && isFourHexDigits(value, index + 2)) {
      return true
    }
    index = value.indexOf('_X', index + 2)
  }
  return false
}

function isFourHexDigits(value: string, startIndex: number): boolean {
  return (
    isHexDigit(value.charCodeAt(startIndex)) &&
    isHexDigit(value.charCodeAt(startIndex + 1)) &&
    isHexDigit(value.charCodeAt(startIndex + 2)) &&
    isHexDigit(value.charCodeAt(startIndex + 3))
  )
}

function isHexDigit(code: number): boolean {
  return (code >= 48 && code <= 57) || (code >= 65 && code <= 70) || (code >= 97 && code <= 102)
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}
