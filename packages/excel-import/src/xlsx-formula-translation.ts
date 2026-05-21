import type { WorkbookTableSnapshot } from '@bilig/protocol'
import { decodeA1CellRef, encodeA1CellRef, type A1CellRef } from './xlsx-a1-utils.js'

type StructuredReferenceSection = 'all' | 'data' | 'headers' | 'this-row' | 'totals'

interface StructuredReferenceParts {
  readonly section?: StructuredReferenceSection
  readonly startColumnName?: string
  readonly endColumnName?: string
}

interface StructuredReferenceRewriteContext {
  readonly formula: string
  readonly ownerSheetName: string
  readonly ownerAddress: string
  readonly tables: readonly WorkbookTableSnapshot[] | undefined
}

const namespacedSpreadsheetFormulaPattern = /^(?:msoxl|of):=/iu

const xlsxFutureFunctionNames: ReadonlySet<string> = new Set([
  'ARRAYTOTEXT',
  'BASE',
  'BITAND',
  'BITLSHIFT',
  'BITOR',
  'BITRSHIFT',
  'BITXOR',
  'CEILING.MATH',
  'CEILING.PRECISE',
  'CHOOSECOLS',
  'CHOOSEROWS',
  'COMBINA',
  'CONCAT',
  'DAYS',
  'DECIMAL',
  'DROP',
  'EXPAND',
  'FILTER',
  'FLOOR.MATH',
  'FLOOR.PRECISE',
  'FORMULATEXT',
  'GAUSS',
  'HSTACK',
  'IFNA',
  'IFS',
  'ISFORMULA',
  'ISOWEEKNUM',
  'ISOMITTED',
  'LAMBDA',
  'LET',
  'MAKEARRAY',
  'MAP',
  'MAXIFS',
  'MINIFS',
  'MUNIT',
  'NUMBERVALUE',
  'PDURATION',
  'PHI',
  'RRI',
  'REDUCE',
  'SCAN',
  'SEQUENCE',
  'SHEET',
  'SHEETS',
  'SORT',
  'SORTBY',
  'SWITCH',
  'TAKE',
  'TEXTAFTER',
  'TEXTBEFORE',
  'TEXTJOIN',
  'TEXTSPLIT',
  'TOCOL',
  'TOROW',
  'UNICHAR',
  'UNICODE',
  'UNIQUE',
  'VALUETOTEXT',
  'VSTACK',
  'WEBSERVICE',
  'WRAPCOLS',
  'WRAPROWS',
  'XLOOKUP',
  'XMATCH',
  'XOR',
] as const)

const xlsxWorksheetFutureFunctionNames: ReadonlySet<string> = new Set(['FILTER'] as const)

export function normalizeImportedFormulaSource(formula: string): string {
  const trimmed = formula.trim()
  const prefix = namespacedSpreadsheetFormulaPattern.exec(trimmed)
  return normalizeImportedLambdaParameterNames(
    transformFormulaFunctionNames(prefix ? trimmed.slice(prefix[0].length) : formula, normalizeImportedFunctionToken),
  )
}

export function encodeFormulaForXlsx(formula: string): string {
  return encodeLambdaParameterNamesForXlsx(transformFormulaFunctionNames(formula, encodeFunctionTokenForXlsx))
}

function isIdentifierStart(character: string): boolean {
  return /[A-Za-z_]/u.test(character)
}

function isIdentifierPart(character: string): boolean {
  return /[A-Za-z0-9_.]/u.test(character)
}

function skipWhitespace(source: string, startIndex: number): number {
  let index = startIndex
  while (index < source.length && /\s/u.test(source[index]!)) {
    index += 1
  }
  return index
}

function transformFormulaFunctionNames(formula: string, transform: (name: string) => string): string {
  let output = ''
  let index = 0
  while (index < formula.length) {
    const character = formula[index]!
    if (character === '"') {
      const endIndex = skipDoubleQuotedString(formula, index)
      output += formula.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === "'") {
      const endIndex = skipSingleQuotedSheetName(formula, index)
      output += formula.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (!isIdentifierStart(character)) {
      output += character
      index += 1
      continue
    }

    let endIndex = index + 1
    while (endIndex < formula.length && isIdentifierPart(formula[endIndex]!)) {
      endIndex += 1
    }
    const name = formula.slice(index, endIndex)
    const callStartIndex = skipWhitespace(formula, endIndex)
    output += formula[callStartIndex] === '(' ? transform(name) : name
    index = endIndex
  }
  return output
}

function normalizeImportedFunctionToken(name: string): string {
  const upper = name.toUpperCase()
  if (upper.startsWith('_XLFN._XLWS.')) {
    return name.slice('_xlfn._xlws.'.length)
  }
  if (upper.startsWith('_XLFN.')) {
    return name.slice('_xlfn.'.length)
  }
  if (upper.startsWith('_XLWS.')) {
    return name.slice('_xlws.'.length)
  }
  return name
}

function encodeFunctionTokenForXlsx(name: string): string {
  const upper = name.toUpperCase()
  if (upper.startsWith('_XLFN.')) {
    return name
  }
  if (upper.startsWith('_XLWS.')) {
    return `_xlfn.${name}`
  }
  if (!xlsxFutureFunctionNames.has(upper)) {
    return name
  }
  return xlsxWorksheetFutureFunctionNames.has(upper) ? `_xlfn._xlws.${upper}` : `_xlfn.${upper}`
}

const lambdaCallTokenPattern = /(?:^|[^A-Za-z0-9_.])((?:_xlfn\.)?LAMBDA)\s*\(/giu
const lambdaParameterNamePattern = /^(?:_xlpm\.)?[A-Za-z_][A-Za-z0-9_.]*$/iu

function normalizeImportedLambdaParameterNames(formula: string): string {
  return formula.replace(/\b_xlpm\.([A-Za-z_][A-Za-z0-9_.]*)/giu, '$1')
}

function encodeLambdaParameterNamesForXlsx(formula: string): string {
  return transformLambdaCalls(formula, encodeLambdaCallParameterNames)
}

function transformLambdaCalls(formula: string, transform: (call: string) => string): string {
  let output = ''
  let cursor = 0
  while (cursor < formula.length) {
    lambdaCallTokenPattern.lastIndex = cursor
    const match = lambdaCallTokenPattern.exec(formula)
    if (!match || match.index === undefined) {
      output += formula.slice(cursor)
      break
    }
    const tokenStart = match.index + match[0].indexOf(match[1]!)
    const openParenIndex = formula.indexOf('(', tokenStart + match[1]!.length)
    const closeParenIndex = openParenIndex >= 0 ? findMatchingParen(formula, openParenIndex) : -1
    if (openParenIndex < 0 || closeParenIndex < 0) {
      output += formula.slice(cursor)
      break
    }
    output += formula.slice(cursor, tokenStart)
    output += transform(formula.slice(tokenStart, closeParenIndex + 1))
    cursor = closeParenIndex + 1
  }
  return output
}

function encodeLambdaCallParameterNames(call: string): string {
  const openParenIndex = call.indexOf('(')
  if (openParenIndex < 0 || !call.endsWith(')')) {
    return call
  }
  const callee = call.slice(0, openParenIndex)
  const args = splitTopLevelFormulaArguments(call.slice(openParenIndex + 1, -1))
  if (args.length < 2) {
    return call
  }
  const parameterNames = args.slice(0, -1).map((arg) => normalizeLambdaParameterName(arg.trim()))
  if (parameterNames.some((name) => name === undefined)) {
    return call
  }
  const scopedParameters = new Set(parameterNames.filter((name): name is string => name !== undefined))
  const encodedParameters = parameterNames.map((name) => `_xlpm.${name}`)
  const encodedBody = prefixLambdaParameterReferences(
    transformLambdaCalls(args[args.length - 1]!, encodeLambdaCallParameterNames),
    scopedParameters,
  )
  return `${callee}(${[...encodedParameters, encodedBody].join(',')})`
}

function normalizeLambdaParameterName(parameter: string): string | undefined {
  if (!lambdaParameterNamePattern.test(parameter)) {
    return undefined
  }
  return parameter.replace(/^_xlpm\./iu, '')
}

function splitTopLevelFormulaArguments(source: string): string[] {
  const args: string[] = []
  let start = 0
  let depth = 0
  let index = 0
  while (index < source.length) {
    const character = source[index]!
    if (character === '"') {
      index = skipDoubleQuotedString(source, index)
      continue
    }
    if (character === "'") {
      index = skipSingleQuotedSheetName(source, index)
      continue
    }
    if (character === '(' || character === '[' || character === '{') {
      depth += 1
    } else if (character === ')' || character === ']' || character === '}') {
      depth = Math.max(0, depth - 1)
    } else if (character === ',' && depth === 0) {
      args.push(source.slice(start, index))
      start = index + 1
    }
    index += 1
  }
  args.push(source.slice(start))
  return args
}

function findMatchingParen(source: string, openParenIndex: number): number {
  let depth = 0
  let index = openParenIndex
  while (index < source.length) {
    const character = source[index]!
    if (character === '"') {
      index = skipDoubleQuotedString(source, index)
      continue
    }
    if (character === "'") {
      index = skipSingleQuotedSheetName(source, index)
      continue
    }
    if (character === '(') {
      depth += 1
    } else if (character === ')') {
      depth -= 1
      if (depth === 0) {
        return index
      }
    }
    index += 1
  }
  return -1
}

function prefixLambdaParameterReferences(source: string, parameterNames: ReadonlySet<string>): string {
  if (parameterNames.size === 0) {
    return source
  }
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
      const endIndex = skipSingleQuotedSheetName(source, index)
      output += source.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (!isIdentifierStart(character)) {
      output += character
      index += 1
      continue
    }
    let endIndex = index + 1
    while (endIndex < source.length && isIdentifierPart(source[endIndex]!)) {
      endIndex += 1
    }
    const name = source.slice(index, endIndex)
    const callStartIndex = skipWhitespace(source, endIndex)
    output += parameterNames.has(name) && source[callStartIndex] !== '(' ? `_xlpm.${name}` : name
    index = endIndex
  }
  return output
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

function skipSingleQuotedSheetName(source: string, startIndex: number): number {
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

function readBalancedStructuredReference(
  source: string,
  startIndex: number,
): { readonly text: string; readonly endIndex: number } | undefined {
  let depth = 0
  for (let index = startIndex; index < source.length; index += 1) {
    const character = source[index]
    if (character === '[') {
      depth += 1
    } else if (character === ']') {
      depth -= 1
      if (depth === 0) {
        return {
          text: source.slice(startIndex + 1, index),
          endIndex: index + 1,
        }
      }
    }
  }
  return undefined
}

function normalizeStructuredReferenceSection(item: string): StructuredReferenceSection | undefined {
  const normalized = item.replace(/\s+/gu, ' ').trim().toUpperCase()
  switch (normalized) {
    case '#ALL':
      return 'all'
    case '#DATA':
      return 'data'
    case '#HEADERS':
      return 'headers'
    case '#THIS ROW':
    case '#THISROW':
    case '@':
      return 'this-row'
    case '#TOTALS':
    case '#TOTAL ROW':
    case '#TOTALS ROW':
      return 'totals'
    default:
      return undefined
  }
}

function unescapeStructuredColumnName(item: string): string {
  return item.replace(/^@/u, '').replace(/''/gu, "'").trim()
}

function hasBalancedOuterBrackets(item: string): boolean {
  if (!item.startsWith('[') || !item.endsWith(']')) {
    return false
  }
  let depth = 0
  for (let index = 0; index < item.length; index += 1) {
    const character = item[index]
    if (character === '[') {
      depth += 1
    } else if (character === ']') {
      depth -= 1
      if (depth === 0) {
        return index === item.length - 1
      }
      if (depth < 0) {
        return false
      }
    }
  }
  return false
}

function unwrapStructuredReferenceItem(item: string): string {
  const trimmed = item.trim()
  return hasBalancedOuterBrackets(trimmed) ? trimmed.slice(1, -1).trim() : trimmed
}

function splitStructuredReferenceTopLevel(text: string, separator: ',' | ':'): string[] | undefined {
  const parts: string[] = []
  let depth = 0
  let startIndex = 0
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === '[') {
      depth += 1
    } else if (character === ']') {
      depth -= 1
      if (depth < 0) {
        return undefined
      }
    } else if (character === separator && depth === 0) {
      parts.push(text.slice(startIndex, index).trim())
      startIndex = index + 1
    }
  }
  if (depth !== 0) {
    return undefined
  }
  parts.push(text.slice(startIndex).trim())
  return parts
}

function parseStructuredReferenceToken(
  item: string,
): { readonly section?: StructuredReferenceSection; readonly columnName?: string } | undefined {
  let trimmed = unwrapStructuredReferenceItem(item)
  if (trimmed.length === 0) {
    return undefined
  }

  const section = normalizeStructuredReferenceSection(trimmed)
  if (section) {
    return { section }
  }

  let tokenSection: StructuredReferenceSection | undefined
  if (trimmed.startsWith('@')) {
    tokenSection = 'this-row'
    trimmed = unwrapStructuredReferenceItem(trimmed.slice(1).trim())
  }
  if (trimmed.length === 0) {
    return tokenSection ? { section: tokenSection } : undefined
  }

  return {
    ...(tokenSection ? { section: tokenSection } : {}),
    columnName: unescapeStructuredColumnName(trimmed),
  }
}

function parseStructuredReferenceParts(text: string): StructuredReferenceParts | undefined {
  if (text.trim().length === 0) {
    return {}
  }

  const items = splitStructuredReferenceTopLevel(text.trim(), ',')
  if (!items) {
    return undefined
  }

  let section: StructuredReferenceSection | undefined
  let startColumnName: string | undefined
  let endColumnName: string | undefined
  for (const item of items) {
    if (item.length === 0) {
      continue
    }

    const spanItems = splitStructuredReferenceTopLevel(item, ':')
    if (!spanItems || spanItems.length === 0 || spanItems.length > 2) {
      return undefined
    }

    if (spanItems.length === 2) {
      const spanStart = parseStructuredReferenceToken(spanItems[0] ?? '')
      const spanEnd = parseStructuredReferenceToken(spanItems[1] ?? '')
      if (!spanStart?.columnName || !spanEnd?.columnName || startColumnName !== undefined) {
        return undefined
      }
      if (spanStart.section) {
        section = spanStart.section
      }
      if (spanEnd.section && spanEnd.section !== section) {
        return undefined
      }
      startColumnName = spanStart.columnName
      endColumnName = spanEnd.columnName
      continue
    }

    const parsedItem = parseStructuredReferenceToken(item)
    if (!parsedItem) {
      continue
    }
    if (parsedItem.section) {
      section = parsedItem.section
    }
    if (parsedItem.columnName) {
      if (startColumnName !== undefined) {
        return undefined
      }
      startColumnName = parsedItem.columnName
      endColumnName = parsedItem.columnName
    }
  }

  return section || startColumnName
    ? {
        ...(section ? { section } : {}),
        ...(startColumnName ? { startColumnName } : {}),
        ...(endColumnName ? { endColumnName } : {}),
      }
    : undefined
}

function decodeAddress(address: string): A1CellRef | undefined {
  try {
    return decodeA1CellRef(address)
  } catch {
    return undefined
  }
}

function encodeAddress(row: number, col: number): string {
  return encodeA1CellRef({ r: row, c: col })
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function formatFormulaReference(sheetName: string, startRow: number, startCol: number, endRow: number, endCol: number): string {
  const startAddress = encodeAddress(startRow, startCol)
  const endAddress = encodeAddress(endRow, endCol)
  const prefix = `${quoteSheetName(sheetName)}!`
  return startAddress === endAddress ? `${prefix}${startAddress}` : `${prefix}${startAddress}:${endAddress}`
}

function normalizeStructuredColumnLookupName(columnName: string): string {
  return columnName.replace(/\s+/gu, ' ').trim().toLocaleLowerCase('en-US')
}

function findTableColumnIndex(table: WorkbookTableSnapshot, columnName: string): number {
  const normalizedColumnName = normalizeStructuredColumnLookupName(columnName)
  return table.columnNames.findIndex((candidate) => normalizeStructuredColumnLookupName(candidate) === normalizedColumnName)
}

function rewriteStructuredReference(
  table: WorkbookTableSnapshot,
  parts: StructuredReferenceParts,
  _ownerSheetName: string,
  ownerAddress: string,
): string | undefined {
  const tableStart = decodeAddress(table.startAddress)
  const tableEnd = decodeAddress(table.endAddress)
  const owner = decodeAddress(ownerAddress)
  if (!tableStart || !tableEnd || !owner) {
    return undefined
  }

  const section = parts.section ?? 'data'
  let startRow = tableStart.r + (table.headerRow && section === 'data' ? 1 : 0)
  let endRow = tableEnd.r - (table.totalsRow && section === 'data' ? 1 : 0)

  if (section === 'all') {
    startRow = tableStart.r
    endRow = tableEnd.r
  } else if (section === 'headers') {
    if (!table.headerRow) {
      return '#REF!'
    }
    startRow = tableStart.r
    endRow = tableStart.r
  } else if (section === 'totals') {
    if (!table.totalsRow) {
      return '#REF!'
    }
    startRow = tableEnd.r
    endRow = tableEnd.r
  } else if (section === 'this-row') {
    if (owner.r < tableStart.r || owner.r > tableEnd.r) {
      return '#REF!'
    }
    startRow = owner.r
    endRow = owner.r
  }

  let startCol = tableStart.c
  let endCol = tableEnd.c
  if (parts.startColumnName) {
    const startColumnIndex = findTableColumnIndex(table, parts.startColumnName)
    const endColumnIndex = findTableColumnIndex(table, parts.endColumnName ?? parts.startColumnName)
    if (startColumnIndex < 0 || endColumnIndex < 0) {
      return '#REF!'
    }
    startCol = tableStart.c + Math.min(startColumnIndex, endColumnIndex)
    endCol = tableStart.c + Math.max(startColumnIndex, endColumnIndex)
  }

  if (endRow < startRow || endCol < startCol) {
    return '#REF!'
  }
  return formatFormulaReference(table.sheetName, startRow, startCol, endRow, endCol)
}

function ownerTableForAddress(
  tables: readonly WorkbookTableSnapshot[],
  ownerSheetName: string,
  ownerAddress: string,
): WorkbookTableSnapshot | undefined {
  const owner = decodeAddress(ownerAddress)
  if (!owner) {
    return undefined
  }
  return tables.find((table) => {
    if (table.sheetName !== ownerSheetName) {
      return false
    }
    const tableStart = decodeAddress(table.startAddress)
    const tableEnd = decodeAddress(table.endAddress)
    return (
      tableStart !== undefined &&
      tableEnd !== undefined &&
      owner.r >= tableStart.r &&
      owner.r <= tableEnd.r &&
      owner.c >= tableStart.c &&
      owner.c <= tableEnd.c
    )
  })
}

function isExternalWorkbookReferencePrefix(source: string, startIndex: number): boolean {
  const match = /^\[([1-9][0-9]*)\]/u.exec(source.slice(startIndex))
  if (!match) {
    return false
  }
  let index = startIndex + match[0].length
  const sheetStart = index
  while (index < source.length && /[A-Za-z0-9_.-]/u.test(source[index] ?? '')) {
    index += 1
  }
  return index > sheetStart && source[index] === '!'
}

export function translateImportedFormulaStructuredReferences({
  formula,
  ownerSheetName,
  ownerAddress,
  tables,
}: StructuredReferenceRewriteContext): string {
  if (!tables || tables.length === 0 || !formula.includes('[')) {
    return formula
  }

  const tablesByName = new Map(tables.map((table) => [table.name.toLocaleLowerCase('en-US'), table]))
  const ownerTable = ownerTableForAddress(tables, ownerSheetName, ownerAddress)
  let output = ''
  let index = 0
  while (index < formula.length) {
    const character = formula[index]!
    if (character === '"') {
      const endIndex = skipDoubleQuotedString(formula, index)
      output += formula.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === "'") {
      const endIndex = skipSingleQuotedSheetName(formula, index)
      output += formula.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === '[') {
      const structuredReference = !isExternalWorkbookReferencePrefix(formula, index)
        ? readBalancedStructuredReference(formula, index)
        : undefined
      const parts = structuredReference ? parseStructuredReferenceParts(structuredReference.text) : undefined
      const rewritten = ownerTable
        ? parts
          ? rewriteStructuredReference(ownerTable, parts, ownerSheetName, ownerAddress)
          : undefined
        : undefined
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
    while (identifierEnd < formula.length && isIdentifierPart(formula[identifierEnd]!)) {
      identifierEnd += 1
    }
    const tableName = formula.slice(index, identifierEnd)
    const table = tablesByName.get(tableName.toLocaleLowerCase('en-US'))
    if (!table || formula[identifierEnd] !== '[') {
      output += tableName
      index = identifierEnd
      continue
    }

    const structuredReference = readBalancedStructuredReference(formula, identifierEnd)
    const parts = structuredReference ? parseStructuredReferenceParts(structuredReference.text) : undefined
    const rewritten = parts ? rewriteStructuredReference(table, parts, ownerSheetName, ownerAddress) : undefined
    if (!structuredReference || !rewritten) {
      output += tableName
      index = identifierEnd
      continue
    }
    output += rewritten
    index = structuredReference.endIndex
  }
  return output
}
