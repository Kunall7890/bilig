import { formatAddress, parseCellAddress, parseStructuredReferenceColumnSpecifier, scanStructuredReferenceBracket } from '@bilig/formula'
import type { WorkbookTableRecord } from '../../workbook-store.js'

export interface InsertedTableCalculatedColumnFormulaWrite {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly source: string
}

export function collectInsertedTableCalculatedColumnFormulaWrites(
  tables: readonly WorkbookTableRecord[],
  sheetName: string,
  start: number,
  count: number,
): InsertedTableCalculatedColumnFormulaWrite[] {
  if (count <= 0) {
    return []
  }
  const end = start + count - 1
  const writes: InsertedTableCalculatedColumnFormulaWrite[] = []
  for (const table of tables) {
    if (table.sheetName !== sheetName || !table.columns?.some((column) => column.calculatedColumnFormula !== undefined)) {
      continue
    }
    const bounds = tableBounds(table)
    if (!bounds) {
      continue
    }
    const dataStartRow = bounds.startRow + (table.headerRow ? 1 : 0)
    const dataEndRow = bounds.endRow - (table.totalsRow ? 1 : 0)
    const fillStart = Math.max(start, dataStartRow)
    const fillEnd = Math.min(end, dataEndRow)
    if (fillEnd < fillStart) {
      continue
    }
    table.columns.forEach((column, columnIndex) => {
      if (column.calculatedColumnFormula === undefined) {
        return
      }
      const col = bounds.startCol + columnIndex
      for (let row = fillStart; row <= fillEnd; row += 1) {
        writes.push({
          sheetName,
          row,
          col,
          source: materializeCalculatedColumnFormulaForRow(table, column.calculatedColumnFormula, row),
        })
      }
    })
  }
  return writes
}

export function materializeCalculatedColumnFormulaForRow(table: WorkbookTableRecord, formula: string, row: number): string {
  const source = formula.trim().startsWith('=') ? formula.trim().slice(1) : formula.trim()
  return rewriteCurrentRowStructuredReferences(table, source, row)
}

function tableBounds(
  table: WorkbookTableRecord,
): { readonly startRow: number; readonly startCol: number; readonly endRow: number; readonly endCol: number } | undefined {
  try {
    const start = parseCellAddress(table.startAddress, table.sheetName)
    const end = parseCellAddress(table.endAddress, table.sheetName)
    return {
      startRow: Math.min(start.row, end.row),
      startCol: Math.min(start.col, end.col),
      endRow: Math.max(start.row, end.row),
      endCol: Math.max(start.col, end.col),
    }
  } catch {
    return undefined
  }
}

function rewriteCurrentRowStructuredReferences(table: WorkbookTableRecord, source: string, row: number): string {
  let output = ''
  let index = 0
  while (index < source.length) {
    const character = source[index]!
    if (character === '"') {
      const endIndex = skipDoubleQuotedString(source, index)
      output += source.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === "'") {
      const endIndex = skipSingleQuotedText(source, index)
      output += source.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === '[') {
      const structuredReference = scanStructuredReferenceBracket(source, index)
      const rewritten = structuredReference ? rewriteCurrentRowStructuredReference(table, structuredReference.content, row) : undefined
      if (structuredReference && rewritten) {
        output += rewritten
        index = structuredReference.endIndex
        continue
      }
      output += character
      index += 1
      continue
    }
    if (!isIdentifierStart(character)) {
      output += character
      index += 1
      continue
    }
    let identifierEnd = index + 1
    while (identifierEnd < source.length && isIdentifierPart(source[identifierEnd]!)) {
      identifierEnd += 1
    }
    const identifier = source.slice(index, identifierEnd)
    if (normalizeTableName(identifier) !== normalizeTableName(table.name) || source[identifierEnd] !== '[') {
      output += identifier
      index = identifierEnd
      continue
    }
    const structuredReference = scanStructuredReferenceBracket(source, identifierEnd)
    const rewritten = structuredReference ? rewriteCurrentRowStructuredReference(table, structuredReference.content, row) : undefined
    if (!structuredReference || !rewritten) {
      output += identifier
      index = identifierEnd
      continue
    }
    output += rewritten
    index = structuredReference.endIndex
  }
  return output
}

function rewriteCurrentRowStructuredReference(table: WorkbookTableRecord, text: string, row: number): string | undefined {
  const trimmed = text.trim()
  if (trimmed.startsWith('@')) {
    return currentRowColumnReference(table, trimmed.slice(1), row)
  }
  const parts = splitStructuredReferenceTopLevel(trimmed, ',')
  if (!parts || parts.length < 2 || parseStructuredReferenceSection(parts[0]!) !== '#this row') {
    return undefined
  }
  return currentRowColumnReference(table, parts.slice(1).join(','), row)
}

function currentRowColumnReference(table: WorkbookTableRecord, text: string, row: number): string | undefined {
  const span = splitStructuredReferenceTopLevel(text.trim(), ':')
  if (!span || span.length < 1 || span.length > 2) {
    return undefined
  }
  const startColumn = parseStructuredReferenceColumnSpecifier(span[0]!)
  const endColumn = parseStructuredReferenceColumnSpecifier(span[1] ?? span[0]!)
  if (!startColumn || !endColumn) {
    return undefined
  }
  const startIndex = findTableColumnIndex(table, startColumn)
  const endIndex = findTableColumnIndex(table, endColumn)
  const bounds = tableBounds(table)
  if (startIndex < 0 || endIndex < 0 || !bounds) {
    return undefined
  }
  const startCol = bounds.startCol + Math.min(startIndex, endIndex)
  const endCol = bounds.startCol + Math.max(startIndex, endIndex)
  const prefix = `${quoteSheetName(table.sheetName)}!`
  const startAddress = `${prefix}${formatAddress(row, startCol)}`
  const endAddress = `${prefix}${formatAddress(row, endCol)}`
  return startAddress === endAddress ? startAddress : `${startAddress}:${endAddress}`
}

function splitStructuredReferenceTopLevel(text: string, separator: ',' | ':'): string[] | undefined {
  const parts: string[] = []
  let current = ''
  let index = 0
  while (index < text.length) {
    const character = text[index]!
    if (character === '[') {
      const scanned = scanStructuredReferenceBracket(text, index)
      if (!scanned) {
        return undefined
      }
      current += text.slice(index, scanned.endIndex)
      index = scanned.endIndex
      continue
    }
    if (character === separator) {
      parts.push(current.trim())
      current = ''
      index += 1
      continue
    }
    current += character
    index += 1
  }
  parts.push(current.trim())
  return parts
}

function findTableColumnIndex(table: WorkbookTableRecord, columnName: string): number {
  const normalizedColumnName = normalizeColumnName(columnName)
  return table.columnNames.findIndex((candidate) => normalizeColumnName(candidate) === normalizedColumnName)
}

function normalizeColumnName(value: string): string {
  return value.replace(/\s+/gu, ' ').trim().toLocaleLowerCase('en-US')
}

function normalizeTableName(value: string): string {
  return value.trim().toLocaleLowerCase('en-US')
}

function parseStructuredReferenceSection(value: string): string {
  const trimmed = value.trim()
  const scanned = scanStructuredReferenceBracket(trimmed, 0)
  const unwrapped = scanned?.endIndex === trimmed.length ? scanned.content : trimmed
  return unwrapped.replace(/\s+/gu, ' ').trim().toLocaleLowerCase('en-US')
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function isIdentifierStart(character: string): boolean {
  return /[A-Za-z_]/u.test(character)
}

function isIdentifierPart(character: string): boolean {
  return /[A-Za-z0-9_.]/u.test(character)
}

function skipDoubleQuotedString(source: string, startIndex: number): number {
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

function skipSingleQuotedText(source: string, startIndex: number): number {
  let index = startIndex + 1
  while (index < source.length) {
    if (source[index] === "'") {
      if (source[index + 1] === "'") {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return source.length
}
