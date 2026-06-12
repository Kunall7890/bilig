import { parseFormula, translateFormulaReferences, type FormulaNode } from '@bilig/formula'
import { ErrorCode, formatErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

import { decodeCellAddress } from './address.js'
import { isStreamingNativeExternalReferenceAlias, normalizeExternalWorkbookReferences } from './streaming-native-external-cache.js'
import { evaluateStreamingNativeWasmFormulas } from './streaming-native-row-chain-wasm.js'
import { normalizeStructuredReferenceColumnName } from './streaming-native-text.js'
import type {
  StreamingNativeMutablePendingCellRow,
  StreamingNativePendingCellRows,
  StreamingNativeSharedStringReference,
} from './streaming-native-cell-arena.js'
import type { XlsxSourceLiteralPatch } from './source-preserving-literal-patches.js'
import type {
  NativeFormulaCell,
  NativeTable,
  PendingCellValue,
  SheetScanState,
  StreamingNativeFormulaCounts,
} from './streaming-native-recalc.js'

interface EvaluationContext {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly rowValues: StreamingNativeMutablePendingCellRow
  readonly sheetRows: StreamingNativePendingCellRows
  readonly sheetRowsByName: ReadonlyMap<string, StreamingNativePendingCellRows>
  readonly externalCachedRowsByAlias: ReadonlyMap<string, StreamingNativePendingCellRows>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
}

interface StreamingNativeFormulaReadTarget {
  readonly source: string
  readonly sheetName: string
  readonly row: number
  readonly col: number
}

export class UnsupportedStreamingNativeFormulaError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'UnsupportedStreamingNativeFormulaError'
  }
}

const emptyCellValue: CellValue = Object.freeze({ tag: ValueTag.Empty })

function stringCellValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

export function evaluateFormulaCells(
  sheetScans: ReadonlyMap<string, SheetScanState>,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
  externalCachedRowsByAlias: ReadonlyMap<string, StreamingNativePendingCellRows>,
  patches: XlsxSourceLiteralPatch[],
): StreamingNativeFormulaCounts {
  const nativeFormulaCells = evaluateStreamingNativeWasmFormulas({
    sheetScans,
    tablesBySheet,
    resolveFormulaSource: (scan, cell) => normalizeExternalWorkbookReferences(resolveFormulaSource(scan, cell)),
  })
  let evaluatedFormulaCellCount = nativeFormulaCells.evaluatedFormulaCellCount
  let patchedFormulaCacheCount = nativeFormulaCells.patches.length
  let unsupportedFormulaCellCount = 0
  patches.push(...nativeFormulaCells.patches)
  const sheetRowsByName = new Map([...sheetScans].map(([sheetName, scan]) => [sheetName, scan.rows] as const))
  for (const scan of sheetScans.values()) {
    const formulaCells = scan.formulaCells.toSorted((left, right) => left.row - right.row || left.col - right.col)
    const formulaPatches = new Map<string, XlsxSourceLiteralPatch>()
    const maxPasses = Math.max(1, formulaCells.length + 1)
    let converged = formulaCells.length === 0
    for (let pass = 0; pass < maxPasses; pass += 1) {
      let changed = false
      for (const cell of formulaCells) {
        if (nativeFormulaCells.processedCells.has(`${cell.sheetName}!${cell.address}`)) {
          continue
        }
        const rowValues = scan.rows.get(cell.row)
        if (!rowValues) {
          continue
        }
        try {
          const formula = resolveFormulaSource(scan, cell)
          const ast = parseStreamingNativeFormula(formula)
          const value = normalizeFormulaResultValue(
            evaluateFormulaAst(ast, {
              sheetName: cell.sheetName,
              row: cell.row,
              col: cell.col,
              rowValues,
              sheetRows: scan.rows,
              sheetRowsByName,
              externalCachedRowsByAlias,
              tablesBySheet,
            }),
          )
          changed = changed || !cellValuesEqual(resolvedCellValue(rowValues.get(cell.col)), value)
          rowValues.set(cell.col, value)
          const patchValue = literalInputForFormulaCache(value)
          if (patchValue === undefined) {
            const resultDetail = value.tag === ValueTag.Error ? ` error ${String(value.code)}` : ''
            throw new UnsupportedStreamingNativeFormulaError(
              `unsupported formula result${resultDetail} at ${cell.sheetName}!${cell.address}`,
            )
          }
          formulaPatches.set(`${cell.sheetName}!${cell.address}`, {
            sheetName: cell.sheetName,
            address: cell.address,
            value: patchValue,
            preserveFormula: true,
          })
        } catch (error) {
          unsupportedFormulaCellCount += 1
          throw error
        }
      }
      if (!changed) {
        converged = true
        break
      }
    }
    if (!converged) {
      throw new UnsupportedStreamingNativeFormulaError(`row-local formulas did not converge for sheet ${scan.sheetName}`)
    }
    evaluatedFormulaCellCount += formulaPatches.size
    patchedFormulaCacheCount += formulaPatches.size
    patches.push(...formulaPatches.values())
  }
  return {
    scannedFormulaCellCount: [...sheetScans.values()].reduce((sum, scan) => sum + scan.scannedFormulaCellCount, 0),
    targetedFormulaCellCount: [...sheetScans.values()].reduce((sum, scan) => sum + scan.formulaCells.length, 0),
    evaluatedFormulaCellCount,
    patchedFormulaCacheCount,
    unsupportedFormulaCellCount,
    nativeKernelFormulaCellCount: nativeFormulaCells.evaluatedFormulaCellCount,
    nativeKernelBatchCount: nativeFormulaCells.batchCount,
  }
}

function cellValuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true
    case ValueTag.Number:
      return right.tag === ValueTag.Number && Object.is(left.value, right.value)
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code
    default:
      return false
  }
}

export function resolveFormulaSource(scan: SheetScanState, cell: NativeFormulaCell): string {
  if (cell.formula) {
    return cell.formula
  }
  if (!cell.sharedFormulaIndex) {
    throw new UnsupportedStreamingNativeFormulaError(`missing formula text at ${cell.sheetName}!${cell.address}`)
  }
  const master = scan.sharedFormulaMasters.get(cell.sharedFormulaIndex)
  if (!master) {
    throw new UnsupportedStreamingNativeFormulaError(
      `missing shared formula master ${cell.sharedFormulaIndex} at ${cell.sheetName}!${cell.address}`,
    )
  }
  return translateFormulaReferences(master.formula, cell.row - master.row, cell.col - master.col)
}

function parseStreamingNativeFormula(formula: string): FormulaNode {
  return parseFormula(normalizeExternalWorkbookReferences(formula))
}

function evaluateFormulaAst(node: FormulaNode, context: EvaluationContext): CellValue {
  switch (node.kind) {
    case 'NumberLiteral':
      return { tag: ValueTag.Number, value: node.value }
    case 'BooleanLiteral':
      return { tag: ValueTag.Boolean, value: node.value }
    case 'StringLiteral':
      return stringCellValue(node.value)
    case 'CellRef':
      return readCellReference(node, context)
    case 'StructuredRef':
      return readStructuredReference(node, context)
    case 'UnaryExpr': {
      const value = evaluateFormulaAst(node.argument, context)
      if (node.operator === '+' && value.tag === ValueTag.String) {
        return value
      }
      const number = coerceNumber(value)
      return { tag: ValueTag.Number, value: node.operator === '-' ? -number : number }
    }
    case 'BinaryExpr':
      return evaluateBinaryExpr(node.operator, evaluateFormulaAst(node.left, context), evaluateFormulaAst(node.right, context))
    case 'CallExpr':
      return evaluateCallExpr(node, context)
    case 'ErrorLiteral':
      return { tag: ValueTag.Error, code: node.code }
    case 'OmittedArgument':
    case 'ArrayConstant':
    case 'NameRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
    case 'InvokeExpr':
      throw new UnsupportedStreamingNativeFormulaError(`unsupported formula node: ${node.kind}`)
  }
}

function readCellReference(node: Extract<FormulaNode, { readonly kind: 'CellRef' }>, context: EvaluationContext): CellValue {
  const sheetName = node.sheetName ?? context.sheetName
  const address = decodeCellAddress(node.ref.replaceAll('$', ''))
  return readScannedCell(sheetName, address.r, address.c, context)
}

function readScannedCell(sheetName: string, row: number, col: number, context: EvaluationContext): CellValue {
  const externalRows = context.externalCachedRowsByAlias.get(sheetName)
  if (externalRows) {
    const rowValues = externalRows.get(row)
    if (!rowValues) {
      throw new UnsupportedStreamingNativeFormulaError(`external workbook cache row is missing for ${sheetName}!${String(row + 1)}`)
    }
    return resolvedCellValue(rowValues.get(col))
  }
  if (isStreamingNativeExternalReferenceAlias(sheetName)) {
    throw new UnsupportedStreamingNativeFormulaError(`external workbook cache sheet is missing for ${sheetName}`)
  }
  const sheetRows = sheetName === context.sheetName ? context.sheetRows : context.sheetRowsByName.get(sheetName)
  if (!sheetRows) {
    throw new UnsupportedStreamingNativeFormulaError(`sheet was not scanned for direct reference: ${sheetName}`)
  }
  const rowValues = sheetName === context.sheetName && row === context.row ? context.rowValues : sheetRows.get(row)
  return resolvedCellValue(rowValues?.get(col))
}

function readStructuredReference(node: Extract<FormulaNode, { readonly kind: 'StructuredRef' }>, context: EvaluationContext): CellValue {
  if (node.endColumnName) {
    throw new UnsupportedStreamingNativeFormulaError(
      `multi-column structured reference is not scalar: ${node.columnName}:${node.endColumnName}`,
    )
  }
  if (node.section !== undefined && node.section !== 'this-row') {
    throw new UnsupportedStreamingNativeFormulaError(`unsupported structured reference section: ${node.section}`)
  }
  const table = findCurrentRowTable(node.tableName, context)
  const columnName = normalizeStructuredReferenceColumnName(node.columnName)
  const columnIndex = table.columns.findIndex((column) => column === columnName)
  const normalizedColumnIndex =
    columnIndex >= 0
      ? columnIndex
      : table.columns.findIndex((column) => column.toLocaleLowerCase('en-US') === columnName.toLocaleLowerCase('en-US'))
  if (normalizedColumnIndex < 0) {
    throw new UnsupportedStreamingNativeFormulaError(`unknown structured reference column: ${node.columnName}`)
  }
  return resolvedCellValue(context.rowValues.get(table.range.s.c + normalizedColumnIndex))
}

function findCurrentRowTable(tableName: string, context: EvaluationContext): NativeTable {
  const tables = context.tablesBySheet.get(context.sheetName) ?? []
  const matching = tables.filter((table) => {
    const nameMatches =
      tableName.length === 0 ||
      table.name.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US') ||
      table.displayName.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US')
    return nameMatches && rowIsInTableDataBody(table, context.row)
  })
  const containingFormulaCell = matching.find((table) => context.col >= table.range.s.c && context.col <= table.range.e.c)
  const table = containingFormulaCell ?? matching[0]
  if (!table) {
    throw new UnsupportedStreamingNativeFormulaError(
      tableName.length === 0 ? 'unable to resolve current-row table' : `unable to resolve current-row table: ${tableName}`,
    )
  }
  return table
}

function rowIsInTableDataBody(table: NativeTable, row: number): boolean {
  const start = table.range.s.r + table.headerRowCount
  const end = table.range.e.r - table.totalsRowCount
  return row >= start && row <= end
}

function evaluateBinaryExpr(operator: string, left: CellValue, right: CellValue): CellValue {
  const error = binaryErrorValue(left, right)
  if (error) {
    return error
  }
  switch (operator) {
    case '+':
      return { tag: ValueTag.Number, value: coerceNumber(left) + coerceNumber(right) }
    case '-':
      return { tag: ValueTag.Number, value: coerceNumber(left) - coerceNumber(right) }
    case '*':
      return { tag: ValueTag.Number, value: coerceNumber(left) * coerceNumber(right) }
    case '/': {
      const divisor = coerceNumber(right)
      return divisor === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : { tag: ValueTag.Number, value: coerceNumber(left) / divisor }
    }
    case '&':
      return stringCellValue(cellValueText(left) + cellValueText(right))
    case '^': {
      const value = Math.pow(coerceNumber(left), coerceNumber(right))
      return Number.isFinite(value) ? { tag: ValueTag.Number, value } : { tag: ValueTag.Error, code: ErrorCode.Num }
    }
    case '=':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) === 0 }
    case '<>':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) !== 0 }
    case '>':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) > 0 }
    case '>=':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) >= 0 }
    case '<':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) < 0 }
    case '<=':
      return { tag: ValueTag.Boolean, value: compareCellValues(left, right) <= 0 }
    case ':':
      throw new UnsupportedStreamingNativeFormulaError(`unsupported binary operator: ${operator}`)
    default:
      throw new UnsupportedStreamingNativeFormulaError(`unknown binary operator: ${operator}`)
  }
}

function binaryErrorValue(left: CellValue, right: CellValue): CellValue | null {
  if (left.tag === ValueTag.Error) {
    return left
  }
  if (right.tag === ValueTag.Error) {
    return right
  }
  return null
}

function evaluateCallExpr(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  const callee = normalizedFormulaFunctionName(node.callee)
  if (callee === 'IF') {
    if (node.args.length < 2 || node.args.length > 3) {
      throw new UnsupportedStreamingNativeFormulaError('IF requires 2 or 3 arguments')
    }
    const condition = evaluateFormulaAst(node.args[0]!, context)
    if (condition.tag === ValueTag.Error) {
      return condition
    }
    return coerceBoolean(condition)
      ? evaluateFormulaAst(node.args[1]!, context)
      : node.args[2]
        ? evaluateFormulaAst(node.args[2], context)
        : { tag: ValueTag.Boolean, value: false }
  }
  if (callee === 'IFS') {
    if (node.args.length < 2 || node.args.length % 2 !== 0) {
      throw new UnsupportedStreamingNativeFormulaError('IFS requires condition/value pairs')
    }
    for (let index = 0; index < node.args.length; index += 2) {
      const condition = evaluateFormulaAst(node.args[index]!, context)
      if (condition.tag === ValueTag.Error) {
        return condition
      }
      if (coerceBoolean(condition)) {
        return evaluateFormulaAst(node.args[index + 1]!, context)
      }
    }
    return { tag: ValueTag.Error, code: ErrorCode.NA }
  }
  if (callee === 'VLOOKUP') {
    return evaluateVlookup(node, context)
  }
  if (callee === 'SUM') {
    return evaluateSum(node, context)
  }
  if (callee === 'COUNTA') {
    return evaluateCounta(node, context)
  }
  if (callee === 'AVERAGE') {
    return evaluateAverage(node, context)
  }
  if (callee === 'ROUND') {
    return evaluateRound(node, context)
  }
  if (callee === 'REPT') {
    return evaluateRept(node, context)
  }
  if (callee === 'ISERROR') {
    return evaluateIsError(node, context)
  }
  if (callee === 'INDEX') {
    return evaluateIndex(node, context)
  }
  if (callee === 'MATCH') {
    return evaluateMatch(node, context)
  }
  throw new UnsupportedStreamingNativeFormulaError(`unsupported function: ${node.callee}`)
}

function evaluateSum(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length < 1) {
    throw new UnsupportedStreamingNativeFormulaError('SUM requires at least 1 argument')
  }
  let total = 0
  for (const argument of node.args) {
    if (argument.kind === 'OmittedArgument') {
      continue
    }
    if (argument.kind === 'RangeRef') {
      for (const value of readScannedCellRange(argument, context)) {
        if (value.tag === ValueTag.Error) {
          return value
        }
        if (value.tag === ValueTag.Number) {
          total += value.value
        }
      }
      continue
    }
    const value = evaluateFormulaAst(argument, context)
    if (value.tag === ValueTag.Error) {
      return value
    }
    total += coerceNumber(value)
  }
  return { tag: ValueTag.Number, value: total }
}

function evaluateCounta(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length < 1) {
    throw new UnsupportedStreamingNativeFormulaError('COUNTA requires at least 1 argument')
  }
  let count = 0
  for (const argument of node.args) {
    const values = argument.kind === 'RangeRef' ? readScannedCellRange(argument, context) : [evaluateFormulaAst(argument, context)]
    for (const value of values) {
      if (value.tag !== ValueTag.Empty) {
        count += 1
      }
    }
  }
  return { tag: ValueTag.Number, value: count }
}

function evaluateAverage(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length < 1) {
    throw new UnsupportedStreamingNativeFormulaError('AVERAGE requires at least 1 argument')
  }
  let total = 0
  let count = 0
  for (const argument of node.args) {
    if (argument.kind === 'RangeRef') {
      for (const value of readScannedCellRange(argument, context)) {
        if (value.tag === ValueTag.Error) {
          return value
        }
        if (value.tag === ValueTag.Number) {
          total += value.value
          count += 1
        }
      }
      continue
    }
    const value = evaluateFormulaAst(argument, context)
    if (value.tag === ValueTag.Error) {
      return value
    }
    total += coerceNumber(value)
    count += 1
  }
  return count === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : { tag: ValueTag.Number, value: total / count }
}

function evaluateRound(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length !== 2) {
    throw new UnsupportedStreamingNativeFormulaError('ROUND requires 2 arguments')
  }
  const valueInput = evaluateFormulaAst(node.args[0]!, context)
  if (valueInput.tag === ValueTag.Error) {
    return valueInput
  }
  const digitsInput = evaluateFormulaAst(node.args[1]!, context)
  if (digitsInput.tag === ValueTag.Error) {
    return digitsInput
  }
  const value = coerceNumber(valueInput)
  const digits = Math.trunc(coerceNumber(digitsInput))
  const rounded = roundHalfAwayFromZero(value, digits)
  return Number.isFinite(rounded) ? { tag: ValueTag.Number, value: rounded } : { tag: ValueTag.Error, code: ErrorCode.Num }
}

function evaluateRept(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length !== 2) {
    throw new UnsupportedStreamingNativeFormulaError('REPT requires 2 arguments')
  }
  const text = cellValueText(evaluateFormulaAst(node.args[0]!, context))
  const count = repeatCountForRept(evaluateFormulaAst(node.args[1]!, context))
  if (count === null || count < 0) {
    return { tag: ValueTag.Error, code: ErrorCode.Value }
  }
  if (text.length * count > 32767) {
    return { tag: ValueTag.Error, code: ErrorCode.Value }
  }
  return stringCellValue(text.repeat(count))
}

function repeatCountForRept(value: CellValue): number | null {
  switch (value.tag) {
    case ValueTag.Number:
      return Number.isFinite(value.value) ? Math.trunc(value.value) : null
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String: {
      const numeric = Number(value.value)
      return Number.isFinite(numeric) ? Math.trunc(numeric) : null
    }
    case ValueTag.Error:
      return null
    default:
      return null
  }
}

function evaluateIsError(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length !== 1) {
    throw new UnsupportedStreamingNativeFormulaError('ISERROR requires 1 argument')
  }
  return { tag: ValueTag.Boolean, value: evaluateFormulaAst(node.args[0]!, context).tag === ValueTag.Error }
}

function evaluateIndex(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length < 2 || node.args.length > 3) {
    throw new UnsupportedStreamingNativeFormulaError('INDEX requires 2 or 3 arguments')
  }
  const rangeArgument = node.args[0]!
  if (rangeArgument.kind !== 'RangeRef') {
    throw new UnsupportedStreamingNativeFormulaError('INDEX array must be a cell range')
  }
  const range = decodeFormulaCellRange(rangeArgument, context.sheetName)
  const rowIndex = integerFromFormulaValue(evaluateFormulaAst(node.args[1]!, context))
  const colIndex = node.args[2] ? integerFromFormulaValue(evaluateFormulaAst(node.args[2], context)) : 1
  if (rowIndex < 1 || colIndex < 1) {
    return { tag: ValueTag.Error, code: ErrorCode.Value }
  }
  if (rowIndex > range.endRow - range.startRow + 1 || colIndex > range.width) {
    return { tag: ValueTag.Error, code: ErrorCode.Ref }
  }
  return readScannedCell(range.sheetName, range.startRow + rowIndex - 1, range.startCol + colIndex - 1, context)
}

function evaluateMatch(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length < 2 || node.args.length > 3) {
    throw new UnsupportedStreamingNativeFormulaError('MATCH requires 2 or 3 arguments')
  }
  const matchType = node.args[2] ? coerceNumber(evaluateFormulaAst(node.args[2], context)) : 1
  if (matchType !== 0) {
    throw new UnsupportedStreamingNativeFormulaError('MATCH supports exact match only')
  }
  const rangeArgument = node.args[1]!
  if (rangeArgument.kind !== 'RangeRef') {
    throw new UnsupportedStreamingNativeFormulaError('MATCH lookup array must be a cell range')
  }
  const range = decodeFormulaCellRange(rangeArgument, context.sheetName)
  if (range.width !== 1 && range.endRow !== range.startRow) {
    throw new UnsupportedStreamingNativeFormulaError('MATCH exact lookup array must be one-dimensional')
  }
  const lookupValue = evaluateFormulaAst(node.args[0]!, context)
  const values = readScannedCellRange(rangeArgument, context)
  for (let index = 0; index < values.length; index += 1) {
    if (compareCellValues(values[index]!, lookupValue) === 0) {
      return { tag: ValueTag.Number, value: index + 1 }
    }
  }
  return { tag: ValueTag.Error, code: ErrorCode.NA }
}

function roundHalfAwayFromZero(value: number, digits: number): number {
  if (!Number.isFinite(value) || !Number.isFinite(digits)) {
    return Number.NaN
  }
  if (digits >= 0) {
    const factor = 10 ** digits
    return (Math.sign(value) * Math.round(Math.abs(value) * factor)) / factor
  }
  const factor = 10 ** -digits
  return Math.sign(value) * Math.round(Math.abs(value) / factor) * factor
}

function readScannedCellRange(
  rangeNode: Extract<FormulaNode, { readonly kind: 'RangeRef' }>,
  context: EvaluationContext,
): readonly CellValue[] {
  const range = decodeFormulaCellRange(rangeNode, context.sheetName)
  const values: CellValue[] = []
  const endCol = range.startCol + range.width - 1
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= endCol; col += 1) {
      values.push(readScannedCell(range.sheetName, row, col, context))
    }
  }
  return values
}

function evaluateVlookup(node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>, context: EvaluationContext): CellValue {
  if (node.args.length < 3 || node.args.length > 4) {
    throw new UnsupportedStreamingNativeFormulaError('VLOOKUP requires 3 or 4 arguments')
  }
  const lookupValue = evaluateFormulaAst(node.args[0]!, context)
  const columnIndex = integerFromFormulaValue(evaluateFormulaAst(node.args[2]!, context))
  if (columnIndex < 1) {
    return { tag: ValueTag.Error, code: ErrorCode.Value }
  }
  const matchMode = vlookupMatchMode(node.args[3], context)
  const range = resolveVlookupTableRange(node.args[1], context)
  if (columnIndex > range.width) {
    return { tag: ValueTag.Error, code: ErrorCode.Ref }
  }
  const resultCol = range.startCol + columnIndex - 1
  let approximateMatchRow: number | null = null
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const candidate = readScannedCell(range.sheetName, row, range.startCol, context)
    const comparison = compareVlookupValues(candidate, lookupValue)
    if (comparison === 0) {
      return readScannedCell(range.sheetName, row, resultCol, context)
    }
    if (matchMode === 'approximate' && comparison < 0) {
      approximateMatchRow = row
      continue
    }
    if (matchMode === 'approximate' && comparison > 0) {
      break
    }
  }
  if (approximateMatchRow !== null) {
    return readScannedCell(range.sheetName, approximateMatchRow, resultCol, context)
  }
  return { tag: ValueTag.Error, code: ErrorCode.NA }
}

function resolveVlookupTableRange(
  node: FormulaNode | undefined,
  context: EvaluationContext,
): {
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly width: number
} {
  if (!node) {
    throw new UnsupportedStreamingNativeFormulaError('VLOOKUP table array must be a scalar cell range')
  }
  if (node.kind === 'RangeRef') {
    return decodeFormulaCellRange(node, context.sheetName)
  }
  if (node.kind === 'CallExpr' && normalizedFormulaFunctionName(node.callee) === 'INDIRECT') {
    return decodeFormulaCellRange(resolveIndirectCellRange(node, context), context.sheetName)
  }
  throw new UnsupportedStreamingNativeFormulaError('VLOOKUP table array must be a scalar cell range')
}

function resolveIndirectCellRange(
  node: Extract<FormulaNode, { readonly kind: 'CallExpr' }>,
  context: EvaluationContext,
): Extract<FormulaNode, { readonly kind: 'RangeRef' }> {
  if (node.args.length < 1 || node.args.length > 2) {
    throw new UnsupportedStreamingNativeFormulaError('INDIRECT requires 1 or 2 arguments')
  }
  if (node.args[1]) {
    const a1Style = evaluateFormulaAst(node.args[1], context)
    if (!coerceBoolean(a1Style)) {
      throw new UnsupportedStreamingNativeFormulaError('INDIRECT R1C1 references are not supported')
    }
  }
  const referenceText = cellValueText(evaluateFormulaAst(node.args[0]!, context))
  const range = parseFormula(referenceText)
  if (range.kind !== 'RangeRef' || range.refKind !== 'cells' || range.sheetEndName !== undefined) {
    throw new UnsupportedStreamingNativeFormulaError('INDIRECT must resolve to a scalar cell range')
  }
  return range
}

function normalizedFormulaFunctionName(name: string): string {
  return name.toUpperCase().replace(/^_XLFN\./u, '')
}

function vlookupMatchMode(node: FormulaNode | undefined, context: EvaluationContext): 'exact' | 'approximate' {
  if (!node) {
    return 'approximate'
  }
  const value = evaluateFormulaAst(node, context)
  if ((value.tag === ValueTag.Boolean && !value.value) || (value.tag === ValueTag.Number && value.value === 0)) {
    return 'exact'
  }
  if ((value.tag === ValueTag.Boolean && value.value) || (value.tag === ValueTag.Number && value.value === 1)) {
    return 'approximate'
  }
  throw new UnsupportedStreamingNativeFormulaError('VLOOKUP range lookup must be TRUE/FALSE or 1/0')
}

function integerFromFormulaValue(value: CellValue): number {
  const number = coerceNumber(value)
  if (!Number.isInteger(number)) {
    throw new UnsupportedStreamingNativeFormulaError(`expected integer formula value, received ${String(number)}`)
  }
  return number
}

function decodeFormulaCellRange(
  range: Extract<FormulaNode, { readonly kind: 'RangeRef' }>,
  currentSheetName: string,
): {
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly width: number
} {
  if (range.refKind !== 'cells' || range.sheetEndName !== undefined) {
    throw new UnsupportedStreamingNativeFormulaError('expected scalar cell range')
  }
  const start = decodeCellAddress(range.start.replaceAll('$', ''))
  const end = decodeCellAddress(range.end.replaceAll('$', ''))
  const startRow = Math.min(start.r, end.r)
  const endRow = Math.max(start.r, end.r)
  const startCol = Math.min(start.c, end.c)
  const endCol = Math.max(start.c, end.c)
  return {
    sheetName: range.sheetName ?? currentSheetName,
    startRow,
    endRow,
    startCol,
    width: endCol - startCol + 1,
  }
}

function compareVlookupValues(left: CellValue, right: CellValue): number {
  try {
    return compareCellValues(left, right)
  } catch (error) {
    throw new UnsupportedStreamingNativeFormulaError(error instanceof Error ? error.message : String(error))
  }
}

function coerceNumber(value: CellValue): number {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String: {
      const numeric = Number(value.value)
      if (Number.isFinite(numeric)) {
        return numeric
      }
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce string to number: ${value.value}`)
    }
    case ValueTag.Error:
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce error to number: ${String(value.code)}`)
    default:
      throw new UnsupportedStreamingNativeFormulaError('cannot coerce unknown value to number')
  }
}

function coerceBoolean(value: CellValue): boolean {
  switch (value.tag) {
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Number:
      return value.value !== 0
    case ValueTag.Empty:
      return false
    case ValueTag.String:
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce string to boolean: ${value.value}`)
    case ValueTag.Error:
      throw new UnsupportedStreamingNativeFormulaError(`cannot coerce error to boolean: ${String(value.code)}`)
    default:
      throw new UnsupportedStreamingNativeFormulaError('cannot coerce unknown value to boolean')
  }
}

function compareCellValues(left: CellValue, right: CellValue): number {
  if (isNumericComparable(left) && isNumericComparable(right)) {
    return coerceNumber(left) - coerceNumber(right)
  }
  const leftText = cellValueText(left).toLocaleLowerCase('en-US')
  const rightText = cellValueText(right).toLocaleLowerCase('en-US')
  return leftText < rightText ? -1 : leftText > rightText ? 1 : 0
}

function isNumericComparable(value: CellValue): boolean {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean || value.tag === ValueTag.Empty
}

function cellValueText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      throw new UnsupportedStreamingNativeFormulaError(`cannot use error as text: ${String(value.code)}`)
    default:
      throw new UnsupportedStreamingNativeFormulaError('cannot use unknown value as text')
  }
}

export function readTargets(
  reads: readonly StreamingNativeFormulaReadTarget[],
  sheetScans: ReadonlyMap<string, SheetScanState>,
): Readonly<Record<string, CellValue>> {
  const output: Record<string, CellValue> = {}
  for (const read of reads) {
    const row = sheetScans.get(read.sheetName)?.rows.get(read.row)
    output[read.source] = resolvedCellValue(row?.get(read.col))
  }
  return output
}

function resolvedCellValue(value: PendingCellValue | undefined): CellValue {
  if (value === undefined || isSharedStringReference(value)) {
    return emptyCellValue
  }
  return value
}

function normalizeFormulaResultValue(value: CellValue): CellValue {
  return value.tag === ValueTag.Empty ? { tag: ValueTag.Number, value: 0 } : value
}

function literalInputForFormulaCache(value: CellValue): XlsxSourceLiteralPatch['value'] | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return Number.isFinite(value.value) ? value.value : undefined
    case ValueTag.Boolean:
      return value.value
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return { kind: 'error', value: value.code === ErrorCode.Field ? '#VALUE!' : formatErrorCode(value.code) }
  }
}

function isSharedStringReference(value: PendingCellValue): value is StreamingNativeSharedStringReference {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}
