import { parseFormula, type FormulaNode } from '@bilig/formula'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { createKernelSync } from '@bilig/wasm-kernel'
import { decodeCellAddress, type XlsxSourceLiteralPatch } from '@bilig/xlsx'

import type { NativeFormulaCell, NativeTable, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'

export interface StreamingNativeWasmRowChainResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

interface RowChainCandidate {
  readonly firstCell: NativeFormulaCell
  readonly secondCell: NativeFormulaCell
  readonly firstOperatorCode: number
  readonly secondScale: number
  readonly secondOffset: number
  readonly leftValue: number
  readonly rightValue: number
  readonly rowValues: Map<number, PendingCellValue>
}

interface NumericSource {
  readonly constant?: number
  readonly column?: number
}

interface FirstFormulaPlan {
  readonly operatorCode: number
  readonly left: NumericSource
  readonly right: NumericSource
}

interface SecondFormulaPlan {
  readonly firstFormulaColumn: number
  readonly scale: number
  readonly offset: number
}

export function evaluateStreamingNativeWasmRowChains(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
}): StreamingNativeWasmRowChainResult {
  const candidates = collectRowChainCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const kernel = createKernelSync()
  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  let batchCount = 0
  for (const group of groupCandidates(candidates).values()) {
    if (group.length === 0) {
      continue
    }
    batchCount += 1
    kernel.init(group.length * 2, 1, 1, 1, 1)
    kernel.evalDenseDirectScalarRowChainStoreTargetBatch(
      Float64Array.from(group.map((candidate) => candidate.leftValue)),
      Float64Array.from(group.map((candidate) => candidate.rightValue)),
      Uint32Array.from(group.map((_candidate, index) => index * 2)),
      Uint32Array.from(group.map((_candidate, index) => index * 2 + 1)),
      group.length,
      group[0]!.firstOperatorCode,
      group[0]!.secondScale,
      group[0]!.secondOffset,
    )
    const numbers = kernel.readNumbers()
    for (let index = 0; index < group.length; index += 1) {
      const candidate = group[index]!
      const firstValue = numbers[index * 2]!
      const secondValue = numbers[index * 2 + 1]!
      if (!Number.isFinite(firstValue) || !Number.isFinite(secondValue)) {
        continue
      }
      candidate.rowValues.set(candidate.firstCell.col, { tag: ValueTag.Number, value: firstValue })
      candidate.rowValues.set(candidate.secondCell.col, { tag: ValueTag.Number, value: secondValue })
      patches.push(formulaPatch(candidate.firstCell, firstValue), formulaPatch(candidate.secondCell, secondValue))
      processedCells.add(cellKey(candidate.firstCell))
      processedCells.add(cellKey(candidate.secondCell))
    }
  }
  return {
    batchCount,
    evaluatedFormulaCellCount: processedCells.size,
    patches,
    processedCells,
  }
}

function collectRowChainCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
}): RowChainCandidate[] {
  const candidates: RowChainCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    const formulaCellsByRow = formulaCellsByRowNumber(scan.formulaCells)
    for (const [row, rowFormulaCells] of formulaCellsByRow.entries()) {
      const rowValues = scan.rows.get(row)
      if (!rowValues) {
        continue
      }
      const formulaCellsByColumn = new Map(rowFormulaCells.map((cell) => [cell.col, cell]))
      for (const secondCell of rowFormulaCells) {
        const secondPlan = tryCompileSecondFormula(scan, secondCell, args)
        if (!secondPlan) {
          continue
        }
        const firstCell = formulaCellsByColumn.get(secondPlan.firstFormulaColumn)
        if (!firstCell) {
          continue
        }
        const firstPlan = tryCompileFirstFormula(scan, firstCell, args)
        if (!firstPlan) {
          continue
        }
        const leftValue = readNumericSource(rowValues, firstPlan.left)
        const rightValue = readNumericSource(rowValues, firstPlan.right)
        if (leftValue === null || rightValue === null || (firstPlan.operatorCode === 5 && rightValue === 0)) {
          continue
        }
        candidates.push({
          firstCell,
          secondCell,
          firstOperatorCode: firstPlan.operatorCode,
          secondScale: secondPlan.scale,
          secondOffset: secondPlan.offset,
          leftValue,
          rightValue,
          rowValues,
        })
      }
    }
  }
  return candidates
}

function tryCompileFirstFormula(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): FirstFormulaPlan | null {
  try {
    return compileFirstFormula(parseFormula(args.resolveFormulaSource(scan, cell)), scan.sheetName, cell.row, cell.col, args.tablesBySheet)
  } catch {
    return null
  }
}

function tryCompileSecondFormula(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): SecondFormulaPlan | null {
  try {
    return compileSecondFormula(parseFormula(args.resolveFormulaSource(scan, cell)), scan.sheetName, cell.row, cell.col, args.tablesBySheet)
  } catch {
    return null
  }
}

function formulaCellsByRowNumber(formulaCells: readonly NativeFormulaCell[]): Map<number, NativeFormulaCell[]> {
  const output = new Map<number, NativeFormulaCell[]>()
  for (const cell of formulaCells) {
    const rowCells = output.get(cell.row) ?? []
    rowCells.push(cell)
    output.set(cell.row, rowCells)
  }
  return output
}

function compileFirstFormula(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): FirstFormulaPlan | null {
  if (node.kind !== 'BinaryExpr') {
    return null
  }
  const operatorCode = rowPairOperatorCode(node.operator)
  if (operatorCode === undefined) {
    return null
  }
  const left = compileNumericSource(node.left, sheetName, row, formulaColumn, tablesBySheet)
  const right = compileNumericSource(node.right, sheetName, row, formulaColumn, tablesBySheet)
  return left && right ? { operatorCode, left, right } : null
}

function rowPairOperatorCode(operator: string): number | undefined {
  switch (operator) {
    case '+':
      return 1
    case '-':
      return 2
    case '*':
      return 4
    case '/':
      return 5
    default:
      return undefined
  }
}

function compileSecondFormula(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): SecondFormulaPlan | null {
  if (node.kind !== 'BinaryExpr') {
    return null
  }
  const leftReference = compileReferenceColumn(node.left, sheetName, row, formulaColumn, tablesBySheet)
  const rightReference = compileReferenceColumn(node.right, sheetName, row, formulaColumn, tablesBySheet)
  if (node.operator === '*') {
    const leftNumber = numberLiteralValue(node.left)
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { firstFormulaColumn: leftReference, scale: rightNumber, offset: 0 }
    }
    if (rightReference !== null && leftNumber !== null) {
      return { firstFormulaColumn: rightReference, scale: leftNumber, offset: 0 }
    }
  }
  if (node.operator === '+') {
    const leftNumber = numberLiteralValue(node.left)
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { firstFormulaColumn: leftReference, scale: 1, offset: rightNumber }
    }
    if (rightReference !== null && leftNumber !== null) {
      return { firstFormulaColumn: rightReference, scale: 1, offset: leftNumber }
    }
  }
  if (node.operator === '-') {
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { firstFormulaColumn: leftReference, scale: 1, offset: -rightNumber }
    }
  }
  return null
}

function compileNumericSource(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): NumericSource | null {
  const constant = numberLiteralValue(node)
  if (constant !== null) {
    return { constant }
  }
  const column = compileReferenceColumn(node, sheetName, row, formulaColumn, tablesBySheet)
  return column === null ? null : { column }
}

function numberLiteralValue(node: FormulaNode): number | null {
  return node.kind === 'NumberLiteral' && Number.isFinite(node.value) ? node.value : null
}

function compileReferenceColumn(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): number | null {
  if (node.kind === 'CellRef') {
    if (node.sheetName && node.sheetName !== sheetName) {
      return null
    }
    const address = decodeCellAddress(node.ref.replaceAll('$', ''))
    return address.r === row ? address.c : null
  }
  if (node.kind !== 'StructuredRef' || node.endColumnName || (node.section && node.section !== 'this-row')) {
    return null
  }
  const table = findCurrentRowTable(node.tableName, sheetName, row, formulaColumn, tablesBySheet)
  if (!table) {
    return null
  }
  const columnIndex = table.columns.findIndex(
    (column) => column.toLocaleLowerCase('en-US') === normalizeStructuredReferenceColumnName(node.columnName).toLocaleLowerCase('en-US'),
  )
  return columnIndex >= 0 ? table.range.s.c + columnIndex : null
}

function findCurrentRowTable(
  tableName: string,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): NativeTable | null {
  const tables = tablesBySheet.get(sheetName) ?? []
  const matching = tables.filter((table) => {
    const nameMatches =
      tableName.length === 0 ||
      table.name.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US') ||
      table.displayName.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US')
    return nameMatches && rowIsInTableDataBody(table, row)
  })
  return matching.find((table) => formulaColumn >= table.range.s.c && formulaColumn <= table.range.e.c) ?? matching[0] ?? null
}

function rowIsInTableDataBody(table: NativeTable, row: number): boolean {
  const start = table.range.s.r + table.headerRowCount
  const end = table.range.e.r - table.totalsRowCount
  return row >= start && row <= end
}

function normalizeStructuredReferenceColumnName(value: string): string {
  return decodeExcelEscapedText(value).replace(/\r\n?/gu, '\n').trim()
}

function decodeExcelEscapedText(value: string): string {
  const escapedUnderscore = '\uE000'
  return value
    .replace(/_x005F_/giu, escapedUnderscore)
    .replace(/_x([0-9a-fA-F]{4})_/gu, (_match, code: string) => {
      const codePoint = Number.parseInt(code, 16)
      return Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : ''
    })
    .replaceAll(escapedUnderscore, '_')
}

function readNumericSource(rowValues: ReadonlyMap<number, PendingCellValue>, source: NumericSource): number | null {
  if (source.constant !== undefined) {
    return source.constant
  }
  if (source.column === undefined) {
    return null
  }
  const value = resolvedCellValue(rowValues.get(source.column))
  return value.tag === ValueTag.Number && Number.isFinite(value.value) ? value.value : null
}

function resolvedCellValue(value: PendingCellValue | undefined): CellValue {
  if (value === undefined || isSharedStringReference(value)) {
    return { tag: ValueTag.Empty }
  }
  return value
}

function isSharedStringReference(value: PendingCellValue): value is Extract<PendingCellValue, { readonly kind: 'shared-string' }> {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}

function groupCandidates(candidates: readonly RowChainCandidate[]): Map<string, RowChainCandidate[]> {
  const output = new Map<string, RowChainCandidate[]>()
  for (const candidate of candidates) {
    const key = `${String(candidate.firstOperatorCode)}:${String(candidate.secondScale)}:${String(candidate.secondOffset)}`
    const group = output.get(key) ?? []
    group.push(candidate)
    output.set(key, group)
  }
  return output
}

function formulaPatch(cell: NativeFormulaCell, value: number): XlsxSourceLiteralPatch {
  return {
    sheetName: cell.sheetName,
    address: cell.address,
    value,
    preserveFormula: true,
  }
}

function cellKey(cell: NativeFormulaCell): string {
  return `${cell.sheetName}!${cell.address}`
}

function emptyResult(): StreamingNativeWasmRowChainResult {
  return {
    batchCount: 0,
    evaluatedFormulaCellCount: 0,
    patches: [],
    processedCells: new Set(),
  }
}
