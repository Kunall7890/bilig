import { parseFormula, type FormulaNode } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createKernelSync } from '@bilig/wasm-kernel'
import { decodeCellAddress, type XlsxSourceLiteralPatch } from '@bilig/xlsx'

import type { NativeFormulaCell, NativeTable, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'

export interface StreamingNativeWasmRowChainResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

export interface StreamingNativeWasmDirectScalarResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

interface RowChainCandidate {
  readonly firstCell: NativeFormulaCell
  readonly secondCell: NativeFormulaCell
  readonly firstOperatorCode: number
  readonly leftValue: number
  readonly rightValue: number
  readonly rowValues: Map<number, PendingCellValue>
}

interface AffineRowChainCandidate extends RowChainCandidate {
  readonly secondKind: 'affine'
  readonly secondScale: number
  readonly secondOffset: number
}

interface DivideRowChainCandidate extends RowChainCandidate {
  readonly secondKind: 'divide'
  readonly denominatorValue: number
}

interface DirectScalarCandidate {
  readonly cell: NativeFormulaCell
  readonly operatorCode: number
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

interface AffineSecondFormulaPlan {
  readonly kind: 'affine'
  readonly firstFormulaColumn: number
  readonly scale: number
  readonly offset: number
}

interface DivideSecondFormulaPlan {
  readonly kind: 'divide'
  readonly firstFormulaColumn: number
  readonly denominator: NumericSource
}

interface InlineDivideSecondFormulaPlan {
  readonly kind: 'inline-divide'
  readonly firstFormula: FirstFormulaPlan
  readonly denominator: NumericSource
}

type SecondFormulaPlan = AffineSecondFormulaPlan | DivideSecondFormulaPlan | InlineDivideSecondFormulaPlan
type CompiledRowChainCandidate = AffineRowChainCandidate | DivideRowChainCandidate

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
    const firstCandidate = group[0]!
    if (firstCandidate.secondKind === 'affine') {
      const affineGroup = group.filter(isAffineRowChainCandidate)
      kernel.evalDenseDirectScalarRowChainStoreTargetBatch(
        Float64Array.from(affineGroup.map((candidate) => candidate.leftValue)),
        Float64Array.from(affineGroup.map((candidate) => candidate.rightValue)),
        Uint32Array.from(affineGroup.map((_candidate, index) => index * 2)),
        Uint32Array.from(affineGroup.map((_candidate, index) => index * 2 + 1)),
        affineGroup.length,
        firstCandidate.firstOperatorCode,
        firstCandidate.secondScale,
        firstCandidate.secondOffset,
      )
    } else {
      const divideGroup = group.filter(isDivideRowChainCandidate)
      kernel.evalDenseDirectScalarRowChainDivideStoreTargetBatch(
        Float64Array.from(divideGroup.map((candidate) => candidate.leftValue)),
        Float64Array.from(divideGroup.map((candidate) => candidate.rightValue)),
        Float64Array.from(divideGroup.map((candidate) => candidate.denominatorValue)),
        Uint32Array.from(divideGroup.map((_candidate, index) => index * 2)),
        Uint32Array.from(divideGroup.map((_candidate, index) => index * 2 + 1)),
        divideGroup.length,
        firstCandidate.firstOperatorCode,
      )
    }
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

export function evaluateStreamingNativeWasmDirectScalars(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): StreamingNativeWasmDirectScalarResult {
  const candidates = collectDirectScalarCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const kernel = createKernelSync()
  kernel.init(candidates.length, 1, 1, 1, 1)
  kernel.evalDirectScalarStoreTargetBatch(
    Uint32Array.from(candidates.map((_candidate, index) => index)),
    Uint8Array.from(candidates.map((candidate) => candidate.operatorCode)),
    Uint32Array.from(candidates.map(() => 0xffffffff)),
    Uint8Array.from(candidates.map(() => ValueTag.Number)),
    Float64Array.from(candidates.map((candidate) => candidate.leftValue)),
    Uint16Array.from(candidates.map(() => ErrorCode.None)),
    Uint32Array.from(candidates.map(() => 0xffffffff)),
    Uint8Array.from(candidates.map(() => ValueTag.Number)),
    Float64Array.from(candidates.map((candidate) => candidate.rightValue)),
    Uint16Array.from(candidates.map(() => ErrorCode.None)),
    Float64Array.from(candidates.map(() => 0)),
  )
  const tags = kernel.readTags()
  const numbers = kernel.readNumbers()
  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  for (let index = 0; index < candidates.length; index += 1) {
    const candidate = candidates[index]!
    const value = numbers[index]!
    if (tags[index] !== ValueTag.Number || !Number.isFinite(value)) {
      continue
    }
    candidate.rowValues.set(candidate.cell.col, { tag: ValueTag.Number, value })
    patches.push(formulaPatch(candidate.cell, value))
    processedCells.add(cellKey(candidate.cell))
  }
  return {
    batchCount: patches.length > 0 ? 1 : 0,
    evaluatedFormulaCellCount: processedCells.size,
    patches,
    processedCells,
  }
}

function collectRowChainCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
}): CompiledRowChainCandidate[] {
  const candidates: CompiledRowChainCandidate[] = []
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
        const firstCandidate = resolveFirstFormulaCandidate(scan, secondCell, secondPlan, rowFormulaCells, formulaCellsByColumn, args)
        if (!firstCandidate) {
          continue
        }
        const { firstCell, firstPlan } = firstCandidate
        const leftValue = readNumericSource(rowValues, firstPlan.left)
        const rightValue = readNumericSource(rowValues, firstPlan.right)
        if (leftValue === null || rightValue === null || (firstPlan.operatorCode === 5 && rightValue === 0)) {
          continue
        }
        if (secondPlan.kind === 'affine') {
          candidates.push({
            firstCell,
            secondCell,
            firstOperatorCode: firstPlan.operatorCode,
            secondKind: 'affine',
            secondScale: secondPlan.scale,
            secondOffset: secondPlan.offset,
            leftValue,
            rightValue,
            rowValues,
          })
          continue
        }
        const denominatorValue = readNumericSource(rowValues, secondPlan.denominator)
        if (denominatorValue === null || denominatorValue === 0) {
          continue
        }
        candidates.push({
          firstCell,
          secondCell,
          firstOperatorCode: firstPlan.operatorCode,
          secondKind: 'divide',
          denominatorValue,
          leftValue,
          rightValue,
          rowValues,
        })
      }
    }
  }
  return candidates
}

function collectDirectScalarCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): DirectScalarCandidate[] {
  const candidates: DirectScalarCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      if (args.skippedCells.has(cellKey(cell))) {
        continue
      }
      const rowValues = scan.rows.get(cell.row)
      if (!rowValues) {
        continue
      }
      const firstPlan = tryCompileFirstFormula(scan, cell, args)
      const operatorCode = firstPlan ? directScalarOperatorCode(firstPlan.operatorCode) : undefined
      if (!firstPlan || operatorCode === undefined) {
        continue
      }
      const leftValue = readNumericSource(rowValues, firstPlan.left)
      const rightValue = readNumericSource(rowValues, firstPlan.right)
      if (leftValue === null || rightValue === null || (firstPlan.operatorCode === 5 && rightValue === 0)) {
        continue
      }
      candidates.push({
        cell,
        operatorCode,
        leftValue,
        rightValue,
        rowValues,
      })
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

function resolveFirstFormulaCandidate(
  scan: SheetScanState,
  secondCell: NativeFormulaCell,
  secondPlan: SecondFormulaPlan,
  rowFormulaCells: readonly NativeFormulaCell[],
  formulaCellsByColumn: ReadonlyMap<number, NativeFormulaCell>,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): { readonly firstCell: NativeFormulaCell; readonly firstPlan: FirstFormulaPlan } | null {
  if (secondPlan.kind !== 'inline-divide') {
    const firstCell = formulaCellsByColumn.get(secondPlan.firstFormulaColumn)
    if (!firstCell) {
      return null
    }
    const firstPlan = tryCompileFirstFormula(scan, firstCell, args)
    return firstPlan ? { firstCell, firstPlan } : null
  }
  for (const firstCell of rowFormulaCells) {
    if (firstCell.col === secondCell.col) {
      continue
    }
    const firstPlan = tryCompileFirstFormula(scan, firstCell, args)
    if (firstPlan && firstFormulaPlansMatch(firstPlan, secondPlan.firstFormula)) {
      return { firstCell, firstPlan }
    }
  }
  return null
}

function firstFormulaPlansMatch(left: FirstFormulaPlan, right: FirstFormulaPlan): boolean {
  if (left.operatorCode !== right.operatorCode) {
    return false
  }
  if (numericSourcesEqual(left.left, right.left) && numericSourcesEqual(left.right, right.right)) {
    return true
  }
  return (
    isCommutativeRowPairOperator(left.operatorCode) &&
    numericSourcesEqual(left.left, right.right) &&
    numericSourcesEqual(left.right, right.left)
  )
}

function isCommutativeRowPairOperator(operatorCode: number): boolean {
  return operatorCode === 1 || operatorCode === 4
}

function numericSourcesEqual(left: NumericSource, right: NumericSource): boolean {
  return left.constant === right.constant && left.column === right.column
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

function directScalarOperatorCode(rowPairCode: number): number | undefined {
  switch (rowPairCode) {
    case 1:
      return 1
    case 2:
      return 2
    case 4:
      return 3
    case 5:
      return 4
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
      return { kind: 'affine', firstFormulaColumn: leftReference, scale: rightNumber, offset: 0 }
    }
    if (rightReference !== null && leftNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: rightReference, scale: leftNumber, offset: 0 }
    }
  }
  if (node.operator === '+') {
    const leftNumber = numberLiteralValue(node.left)
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: leftReference, scale: 1, offset: rightNumber }
    }
    if (rightReference !== null && leftNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: rightReference, scale: 1, offset: leftNumber }
    }
  }
  if (node.operator === '-') {
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: leftReference, scale: 1, offset: -rightNumber }
    }
  }
  if (node.operator === '/' && leftReference !== null) {
    const denominator = compileNumericSource(node.right, sheetName, row, formulaColumn, tablesBySheet)
    if (denominator) {
      return { kind: 'divide', firstFormulaColumn: leftReference, denominator }
    }
  }
  if (node.operator === '/' && leftReference === null) {
    const firstFormula = compileFirstFormula(node.left, sheetName, row, formulaColumn, tablesBySheet)
    const denominator = compileNumericSource(node.right, sheetName, row, formulaColumn, tablesBySheet)
    if (firstFormula && denominator) {
      return { kind: 'inline-divide', firstFormula, denominator }
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

function groupCandidates(candidates: readonly CompiledRowChainCandidate[]): Map<string, CompiledRowChainCandidate[]> {
  const output = new Map<string, CompiledRowChainCandidate[]>()
  for (const candidate of candidates) {
    const key =
      candidate.secondKind === 'affine'
        ? `affine:${String(candidate.firstOperatorCode)}:${String(candidate.secondScale)}:${String(candidate.secondOffset)}`
        : `divide:${String(candidate.firstOperatorCode)}`
    const group = output.get(key) ?? []
    group.push(candidate)
    output.set(key, group)
  }
  return output
}

function isAffineRowChainCandidate(candidate: CompiledRowChainCandidate): candidate is AffineRowChainCandidate {
  return candidate.secondKind === 'affine'
}

function isDivideRowChainCandidate(candidate: CompiledRowChainCandidate): candidate is DivideRowChainCandidate {
  return candidate.secondKind === 'divide'
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
