import type { StructuredReferenceSection, StructuredRefNode } from './ast.js'

const STRUCTURED_REFERENCE_ESCAPED_CHARACTERS = new Set(['[', ']', '#', "'", '@'])

export interface StructuredReferenceSpecifier {
  readonly columnName: string
  readonly endColumnName?: string
  readonly section?: StructuredReferenceSection
}

export function isStructuredReferenceEscapedCharacter(character: string | undefined): boolean {
  return character !== undefined && STRUCTURED_REFERENCE_ESCAPED_CHARACTERS.has(character)
}

export function scanStructuredReferenceBracket(
  source: string,
  startIndex: number,
): { readonly content: string; readonly endIndex: number } | undefined {
  if (source[startIndex] !== '[') {
    return undefined
  }

  let depth = 0
  for (let index = startIndex; index < source.length; index += 1) {
    const character = source[index]
    if (character === "'" && isStructuredReferenceEscapedCharacter(source[index + 1])) {
      index += 1
      continue
    }
    if (character === '[') {
      depth += 1
      continue
    }
    if (character === ']') {
      depth -= 1
      if (depth === 0) {
        return {
          content: source.slice(startIndex + 1, index),
          endIndex: index + 1,
        }
      }
      if (depth < 0) {
        return undefined
      }
    }
  }
  return undefined
}

export function parseStructuredReferenceColumnSpecifier(source: string): string | undefined {
  const trimmed = source.trim()
  if (trimmed.length === 0) {
    return undefined
  }

  const columnText = unwrapSingleStructuredReferenceColumnItem(trimmed)
  return columnText === undefined ? undefined : unescapeStructuredReferenceColumnSpecifier(columnText)
}

export function formatStructuredReferenceColumnSpecifier(columnName: string): string {
  let output = ''
  for (const character of columnName) {
    output += isStructuredReferenceEscapedCharacter(character) ? `'${character}` : character
  }
  return output
}

export function parseStructuredReferenceSpecifier(source: string): StructuredReferenceSpecifier | undefined {
  const trimmed = source.trim()
  if (trimmed.length === 0) {
    return undefined
  }
  if (!trimmed.startsWith('[') && !trimmed.startsWith('@')) {
    const columnName = parseStructuredReferenceColumnSpecifier(trimmed)
    return columnName === undefined ? undefined : { columnName }
  }

  const items = splitStructuredReferenceTopLevel(trimmed, ',')
  if (!items) {
    return undefined
  }
  let section: StructuredReferenceSection | undefined
  let columnName: string | undefined
  let endColumnName: string | undefined

  for (const item of items) {
    const span = parseStructuredReferenceColumnSpan(item)
    if (span) {
      section = section ?? span.section
      columnName = span.columnName
      endColumnName = span.endColumnName
      continue
    }

    const token = parseStructuredReferenceToken(item)
    if (!token) {
      return undefined
    }
    if (token.section) {
      section = section ?? token.section
    }
    if (token.columnName) {
      columnName = token.columnName
    }
  }

  if (!columnName) {
    return undefined
  }
  return {
    columnName,
    ...(endColumnName !== undefined ? { endColumnName } : {}),
    ...(section !== undefined ? { section } : {}),
  }
}

export function formatStructuredReferenceSpecifier(node: StructuredRefNode): string {
  const column = formatStructuredReferenceColumnSpecifier(node.columnName)
  const endColumn = node.endColumnName === undefined ? undefined : formatStructuredReferenceColumnSpecifier(node.endColumnName)
  if (node.section === 'this-row' && node.tableName.length === 0) {
    return endColumn === undefined ? `@${column}` : `@[${column}]:[${endColumn}]`
  }
  if (node.section === undefined) {
    return endColumn === undefined ? column : `[${column}]:[${endColumn}]`
  }
  const section = formatStructuredReferenceSection(node.section)
  return endColumn === undefined ? `[${section}],[${column}]` : `[${section}],[${column}]:[${endColumn}]`
}

function unwrapSingleStructuredReferenceColumnItem(source: string): string | undefined {
  if (!source.startsWith('[')) {
    return source
  }
  if (!source.endsWith(']')) {
    return undefined
  }
  const scanned = scanStructuredReferenceBracket(source, 0)
  if (scanned?.endIndex !== source.length) {
    return undefined
  }
  return scanned.content.trim()
}

function unescapeStructuredReferenceColumnSpecifier(source: string): string {
  let output = ''
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index]!
    const next = source[index + 1]
    if (character === "'" && isStructuredReferenceEscapedCharacter(next)) {
      output += next
      index += 1
      continue
    }
    output += character
  }
  return output
}

interface StructuredReferenceToken {
  readonly columnName?: string
  readonly section?: StructuredReferenceSection
}

function parseStructuredReferenceColumnSpan(source: string): StructuredReferenceSpecifier | undefined {
  const parts = splitStructuredReferenceTopLevel(source.trim(), ':')
  if (!parts || parts.length !== 2) {
    return undefined
  }
  const leftSource = parts[0]?.trim() ?? ''
  const rightSource = parts[1]?.trim() ?? ''
  if (!leftSource.includes('[') && !rightSource.includes('[')) {
    return undefined
  }
  const left = parseStructuredReferenceToken(leftSource)
  const right = parseStructuredReferenceToken(rightSource)
  if (!left?.columnName || !right?.columnName) {
    return undefined
  }
  return {
    columnName: left.columnName,
    endColumnName: right.columnName,
    ...((left.section ?? right.section) ? { section: left.section ?? right.section } : {}),
  }
}

function parseStructuredReferenceToken(source: string): StructuredReferenceToken | undefined {
  let trimmed = unwrapStructuredReferenceItem(source.trim())
  let section: StructuredReferenceSection | undefined
  if (trimmed.startsWith('@')) {
    section = 'this-row'
    trimmed = unwrapStructuredReferenceItem(trimmed.slice(1).trim())
  }
  const normalizedSection = normalizeStructuredReferenceSection(trimmed)
  if (normalizedSection) {
    return { section: normalizedSection }
  }
  const columnName = parseStructuredReferenceColumnSpecifier(trimmed)
  return columnName === undefined ? undefined : { columnName, ...(section !== undefined ? { section } : {}) }
}

function unwrapStructuredReferenceItem(source: string): string {
  return unwrapSingleStructuredReferenceColumnItem(source) ?? source
}

function normalizeStructuredReferenceSection(source: string): StructuredReferenceSection | undefined {
  const normalized = source.trim().replace(/^#/u, '').replace(/\s+/gu, ' ').toUpperCase()
  switch (normalized) {
    case 'ALL':
      return 'all'
    case 'DATA':
      return 'data'
    case 'HEADERS':
      return 'headers'
    case 'THIS ROW':
      return 'this-row'
    case 'TOTALS':
    case 'TOTAL ROW':
    case 'TOTALS ROW':
      return 'totals'
    default:
      return undefined
  }
}

function formatStructuredReferenceSection(section: StructuredReferenceSection): string {
  switch (section) {
    case 'all':
      return '#All'
    case 'data':
      return '#Data'
    case 'headers':
      return '#Headers'
    case 'this-row':
      return '#This Row'
    case 'totals':
      return '#Totals'
  }
}

function splitStructuredReferenceTopLevel(text: string, separator: ',' | ':'): string[] | undefined {
  const parts: string[] = []
  let depth = 0
  let start = 0
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === "'" && isStructuredReferenceEscapedCharacter(text[index + 1])) {
      index += 1
      continue
    }
    if (character === '[') {
      depth += 1
      continue
    }
    if (character === ']') {
      depth -= 1
      if (depth < 0) {
        return undefined
      }
      continue
    }
    if (character === separator && depth === 0) {
      parts.push(text.slice(start, index).trim())
      start = index + 1
    }
  }
  if (depth !== 0) {
    return undefined
  }
  parts.push(text.slice(start).trim())
  return parts
}
