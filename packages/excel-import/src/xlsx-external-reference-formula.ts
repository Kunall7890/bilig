import type { ExternalCachedValue } from './xlsx-external-cache.js'

export interface ParsedCellAddress {
  readonly row: number
  readonly col: number
}

export function skipDoubleQuotedString(source: string, startIndex: number): number {
  let index = startIndex + 1
  while (index < source.length) {
    if (source[index] === '"') {
      if (source[index + 1] === '"') {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return source.length
}

export function readSingleQuotedIdentifier(
  source: string,
  startIndex: number,
): { readonly value: string; readonly endIndex: number } | null {
  let index = startIndex + 1
  let value = ''
  while (index < source.length) {
    const current = source[index]
    if (current === "'") {
      if (source[index + 1] === "'") {
        value += "'"
        index += 2
        continue
      }
      return { value, endIndex: index + 1 }
    }
    value += current ?? ''
    index += 1
  }
  return null
}

export function readCellAddress(source: string, startIndex: number): { readonly address: string; readonly endIndex: number } | null {
  const match = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}/u.exec(source.slice(startIndex))
  if (!match) {
    return null
  }
  const address = match[0]
  const endIndex = startIndex + address.length
  if (/^[A-Za-z0-9_.$]/u.test(source[endIndex] ?? '')) {
    return null
  }
  return { address, endIndex }
}

export function parseCachedCellAddress(address: string): ParsedCellAddress | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/u.exec(address.replaceAll('$', '').toUpperCase())
  if (!match) {
    return null
  }
  let col = 0
  for (const character of match[1]!) {
    col = col * 26 + character.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  return col > 0 && Number.isSafeInteger(row) ? { row: row - 1, col: col - 1 } : null
}

export function readExternalSheetName(value: string): { readonly bookIndex: number; readonly sheetName: string } | null {
  const match = /^\[([1-9][0-9]*)\](.+)$/u.exec(value)
  if (!match) {
    return null
  }
  return { bookIndex: Number(match[1]), sheetName: match[2] ?? '' }
}

export function readUnquotedExternalSheetReference(
  source: string,
  startIndex: number,
): { readonly bookIndex: number; readonly sheetName: string; readonly endIndex: number } | null {
  const prefix = /^\[([1-9][0-9]*)\]/u.exec(source.slice(startIndex))
  if (!prefix) {
    return null
  }
  const bookIndex = Number(prefix[1])
  let index = startIndex + prefix[0].length
  const sheetStart = index
  while (index < source.length && /[A-Za-z0-9_.-]/u.test(source[index] ?? '')) {
    index += 1
  }
  if (index === sheetStart || source[index] !== '!') {
    return null
  }
  return { bookIndex, sheetName: source.slice(sheetStart, index), endIndex: index }
}

export function isRangeOrSpillReference(source: string, addressStartIndex: number, addressEndIndex: number): boolean {
  return source[addressStartIndex - 1] === ':' || source[addressEndIndex] === ':' || source[addressEndIndex] === '#'
}

export function formatA1Address(row: number, col: number): string {
  let current = col + 1
  let column = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    column = String.fromCharCode(65 + remainder) + column
    current = Math.floor((current - 1) / 26)
  }
  return `${column}${String(row + 1)}`
}

export function formatFormulaLiteral(value: ExternalCachedValue): string {
  switch (value.kind) {
    case 'number':
      return String(value.value)
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE'
    case 'error':
      return value.value
    case 'string':
      return `"${value.value.replaceAll('"', '""')}"`
  }
}

export function quoteFormulaSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

export function formatAbsoluteA1Address(row: number, col: number): string {
  const address = formatA1Address(row, col)
  const match = /^([A-Z]+)([1-9][0-9]*)$/u.exec(address)
  return match ? `$${match[1]}$${match[2]}` : address
}
