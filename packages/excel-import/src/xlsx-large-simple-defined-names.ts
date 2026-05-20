import type { WorkbookDefinedNameSnapshot, WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'

const definedNameElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?definedName>)/gu

export function readWorkbookDefinedNames(
  workbookXml: string,
  sheetNames: readonly string[],
): {
  readonly definedNames: WorkbookDefinedNameSnapshot[] | undefined
  readonly externalWorkbookReferenceSeen: boolean
  readonly ignoredCount: number
} {
  const definedNamesByKey = new Map<string, WorkbookDefinedNameSnapshot>()
  let externalWorkbookReferenceSeen = false
  let ignoredCount = 0
  for (const match of workbookXml.matchAll(definedNameElementPattern)) {
    const xml = match[0]
    const openingTag = /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xml)?.[0]
    const name = openingTag ? readXmlAttribute(openingTag, 'name')?.trim() : ''
    const localSheetId = openingTag ? readNonNegativeIntegerAttribute(openingTag, 'localSheetId') : null
    const scopeSheetName = localSheetId !== null ? sheetNames[localSheetId] : undefined
    const rawValue = openingTag?.endsWith('/>') ? '' : decodeXmlText(xml.replace(/^<[^>]*>/u, '').replace(/<\/[^>]*>$/u, '')).trim()
    if (!name || rawValue.length === 0 || (localSheetId !== null && scopeSheetName === undefined)) {
      ignoredCount += 1
      continue
    }
    if (definedNameReferencesExternalWorkbook(rawValue)) {
      externalWorkbookReferenceSeen = true
      continue
    }
    const value = isBuiltInPrintDefinedName(name)
      ? parsePrintDefinedNameValue(rawValue)
      : parseDefinedNameValue(rawValue, new Set(sheetNames))
    if (!value) {
      ignoredCount += 1
      continue
    }
    definedNamesByKey.set(definedNameKey(name, scopeSheetName), {
      name,
      ...(scopeSheetName !== undefined ? { scopeSheetName } : {}),
      value,
    })
  }
  const definedNames = [...definedNamesByKey.values()].toSorted(
    (left, right) => left.name.localeCompare(right.name) || (left.scopeSheetName ?? '').localeCompare(right.scopeSheetName ?? ''),
  )
  return {
    definedNames: definedNames.length > 0 ? definedNames : undefined,
    externalWorkbookReferenceSeen,
    ignoredCount,
  }
}

function definedNameReferencesExternalWorkbook(value: string): boolean {
  return /(?:^|[=,+(*/\s])'?\[[^\]]+\]/u.test(value)
}

function isBuiltInPrintDefinedName(name: string): boolean {
  const normalized = name.trim().toLocaleLowerCase('en-US')
  return normalized === '_xlnm.print_area' || normalized === '_xlnm.print_titles'
}

function parsePrintDefinedNameValue(value: string): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = value.trim()
  return trimmed.length > 0 ? { kind: 'formula', formula: trimmed.startsWith('=') ? trimmed : `=${trimmed}` } : null
}

function parseDefinedNameValue(value: string, sheetNames: ReadonlySet<string>): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = value.trim()
  if (trimmed.length === 0) {
    return null
  }
  const expression = trimmed.startsWith('=') ? trimmed.slice(1).trim() : trimmed
  const sheetReference = parseSheetReference(expression)
  if (sheetReference && sheetNames.has(sheetReference.sheetName)) {
    const parsedReference = parseDefinedNameReferenceValue(sheetReference.sheetName, sheetReference.reference)
    if (parsedReference) {
      return parsedReference
    }
  }
  const scalar = parseDefinedNameScalarValue(expression)
  if (scalar) {
    return scalar
  }
  return { kind: 'formula', formula: trimmed.startsWith('=') ? trimmed : `=${trimmed}` }
}

function parseDefinedNameReferenceValue(sheetName: string, reference: string): WorkbookDefinedNameValueSnapshot | null {
  const parts = reference.split(':')
  if (parts.length === 1) {
    const address = normalizeDefinedNameCellAddress(parts[0] ?? '')
    return address ? { kind: 'cell-ref', sheetName, address } : null
  }
  if (parts.length === 2) {
    const startAddress = normalizeDefinedNameCellAddress(parts[0] ?? '')
    const endAddress = normalizeDefinedNameCellAddress(parts[1] ?? '')
    return startAddress && endAddress ? { kind: 'range-ref', sheetName, startAddress, endAddress } : null
  }
  return null
}

function parseDefinedNameScalarValue(value: string): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = value.trim()
  if (/^[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?$/u.test(trimmed)) {
    const numberValue = Number(trimmed)
    return Number.isFinite(numberValue) ? { kind: 'scalar', value: numberValue } : null
  }
  if (/^TRUE$/iu.test(trimmed)) {
    return { kind: 'scalar', value: true }
  }
  if (/^FALSE$/iu.test(trimmed)) {
    return { kind: 'scalar', value: false }
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return { kind: 'scalar', value: trimmed.slice(1, -1).replace(/""/gu, '"') }
  }
  return null
}

function parseSheetReference(value: string): { readonly sheetName: string; readonly reference: string } | null {
  const quoted = parseQuotedSheetReference(value)
  if (quoted) {
    return quoted
  }
  const separatorIndex = value.indexOf('!')
  if (separatorIndex <= 0) {
    return null
  }
  const sheetName = value.slice(0, separatorIndex).trim()
  const reference = value.slice(separatorIndex + 1).trim()
  return sheetName.length > 0 && reference.length > 0 ? { sheetName, reference } : null
}

function parseQuotedSheetReference(value: string): { readonly sheetName: string; readonly reference: string } | null {
  if (!value.startsWith("'")) {
    return null
  }
  let sheetName = ''
  for (let index = 1; index < value.length; index += 1) {
    const character = value[index]
    if (character === "'" && value[index + 1] === "'") {
      sheetName += "'"
      index += 1
      continue
    }
    if (character === "'" && value[index + 1] === '!') {
      const reference = value.slice(index + 2).trim()
      return sheetName.trim().length > 0 && reference.length > 0 ? { sheetName, reference } : null
    }
    sheetName += character
  }
  return null
}

function normalizeDefinedNameCellAddress(value: string): string | null {
  const normalized = value.trim().replaceAll('$', '').toUpperCase()
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/u.exec(normalized)
  return match && decodeCellAddress(normalized) ? normalized : null
}

function definedNameKey(name: string, scopeSheetName: string | undefined): string {
  return `${scopeSheetName ?? '<workbook>'}\u0000${name.toUpperCase()}`
}

export function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function readNumberAttribute(xml: string, attributeName: string): number | null {
  const raw = readXmlAttribute(xml, attributeName)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const value = Number(raw)
  return Number.isFinite(value) ? value : null
}

function readNonNegativeIntegerAttribute(xml: string, attributeName: string): number | null {
  const value = readNumberAttribute(xml, attributeName)
  return Number.isInteger(value) && value !== null && value >= 0 ? value : null
}

export function decodeXmlText(value: string): string {
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

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function decodeCellAddress(address: string): { readonly row: number; readonly column: number } | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/iu.exec(address.replaceAll('$', ''))
  if (!match) {
    return null
  }
  let column = 0
  for (const letter of match[1]?.toUpperCase() ?? '') {
    column = column * 26 + letter.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  if (!Number.isSafeInteger(row) || row <= 0 || column <= 0) {
    return null
  }
  return { row: row - 1, column: column - 1 }
}

export function resolveTargetPath(basePath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1)
  }
  const parts = basePath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
}
