import { parseFormula, type FormulaNode } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createKernelSync } from '@bilig/wasm-kernel'

import { decodeCellAddress } from './address.js'
import type { XlsxSourceLiteralPatch } from './source-preserving-literal-patches.js'
import { compileReferenceColumn } from './streaming-native-row-chain-references.js'
import type { StreamingNativeWasmConditionalResult } from './streaming-native-row-chain-wasm.js'
import type { NativeFormulaCell, NativeTable, PendingCellRow, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'

interface ConditionalPickCandidate {
  readonly cell: NativeFormulaCell
  readonly conditions: readonly ConditionalPickCondition[]
  readonly branchValues: readonly CellValue[]
  readonly defaultValue: CellValue
  readonly rowValues: PendingCellRow
}

interface ConditionalPickCondition {
  readonly operatorCode: number
  readonly left: CellValue
  readonly right: CellValue
}

interface ConditionalCompileContext {
  readonly sheetName: string
  readonly row: number
  readonly formulaColumn: number
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly rowValues: PendingCellRow
}

const directConditionOpTruthy = 1
const directConditionOpEq = 2
const directConditionOpNeq = 3
const directConditionOpGt = 4
const directConditionOpGte = 5
const directConditionOpLt = 6
const directConditionOpLte = 7

export function evaluateStreamingNativeWasmConditionals(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): StreamingNativeWasmConditionalResult {
  const candidates = collectConditionalPickCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const stringArena = new KernelStringArena()
  const conditionStarts: number[] = []
  const conditionLengths: number[] = []
  const conditionOps: number[] = []
  const leftTags: number[] = []
  const leftNumbers: number[] = []
  const leftStringIds: number[] = []
  const leftErrors: number[] = []
  const rightTags: number[] = []
  const rightNumbers: number[] = []
  const rightStringIds: number[] = []
  const rightErrors: number[] = []
  const branchTags: number[] = []
  const branchNumbers: number[] = []
  const branchStringIds: number[] = []
  const branchErrors: number[] = []
  const defaultTags: number[] = []
  const defaultNumbers: number[] = []
  const defaultStringIds: number[] = []
  const defaultErrors: number[] = []
  for (const candidate of candidates) {
    conditionStarts.push(conditionOps.length)
    conditionLengths.push(candidate.conditions.length)
    for (let index = 0; index < candidate.conditions.length; index += 1) {
      const condition = candidate.conditions[index]!
      conditionOps.push(condition.operatorCode)
      appendKernelValueFields(condition.left, stringArena, 'compare', leftTags, leftNumbers, leftStringIds, leftErrors)
      appendKernelValueFields(condition.right, stringArena, 'compare', rightTags, rightNumbers, rightStringIds, rightErrors)
      appendKernelValueFields(
        candidate.branchValues[index]!,
        stringArena,
        'output',
        branchTags,
        branchNumbers,
        branchStringIds,
        branchErrors,
      )
    }
    appendKernelValueFields(candidate.defaultValue, stringArena, 'output', defaultTags, defaultNumbers, defaultStringIds, defaultErrors)
  }

  const outTags = new Uint8Array(candidates.length)
  const outNumbers = new Float64Array(candidates.length)
  const outStringIds = new Uint32Array(candidates.length)
  const outErrors = new Uint16Array(candidates.length)
  createKernelSync().evalDirectConditionalPickBatch(
    Uint32Array.from(conditionStarts),
    Uint32Array.from(conditionLengths),
    Uint8Array.from(conditionOps),
    Uint8Array.from(leftTags),
    Float64Array.from(leftNumbers),
    Uint32Array.from(leftStringIds),
    Uint16Array.from(leftErrors),
    Uint8Array.from(rightTags),
    Float64Array.from(rightNumbers),
    Uint32Array.from(rightStringIds),
    Uint16Array.from(rightErrors),
    Uint8Array.from(branchTags),
    Float64Array.from(branchNumbers),
    Uint32Array.from(branchStringIds),
    Uint16Array.from(branchErrors),
    Uint8Array.from(defaultTags),
    Float64Array.from(defaultNumbers),
    Uint32Array.from(defaultStringIds),
    Uint16Array.from(defaultErrors),
    outTags,
    outNumbers,
    outStringIds,
    outErrors,
  )

  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  for (let index = 0; index < candidates.length; index += 1) {
    const value = cellValueFromKernelFields(outTags[index]!, outNumbers[index]!, outStringIds[index]!, outErrors[index]!, stringArena)
    const patchValue = value ? cellValueToPatchValue(value) : undefined
    if (!value || patchValue === undefined) {
      continue
    }
    const candidate = candidates[index]!
    candidate.rowValues.set(candidate.cell.col, value)
    patches.push(formulaPatch(candidate.cell, patchValue))
    processedCells.add(cellKey(candidate.cell))
  }
  return {
    batchCount: processedCells.size > 0 ? 1 : 0,
    evaluatedFormulaCellCount: processedCells.size,
    patches,
    processedCells,
  }
}

function collectConditionalPickCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): ConditionalPickCandidate[] {
  const candidates: ConditionalPickCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      if (args.skippedCells.has(cellKey(cell))) {
        continue
      }
      const rowValues = scan.rows.get(cell.row)
      if (!rowValues) {
        continue
      }
      const candidate = tryCompileConditionalPickCandidate(scan, cell, rowValues, args)
      if (candidate) {
        candidates.push(candidate)
      }
    }
  }
  return candidates
}

function tryCompileConditionalPickCandidate(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  rowValues: PendingCellRow,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): ConditionalPickCandidate | null {
  let node: FormulaNode
  try {
    node = parseFormula(args.resolveFormulaSource(scan, cell))
  } catch {
    return null
  }
  const context = { sheetName: scan.sheetName, row: cell.row, formulaColumn: cell.col, tablesBySheet: args.tablesBySheet, rowValues }
  if (node.kind === 'CallExpr') {
    const callee = node.callee.toLocaleUpperCase('en-US').replace(/^_XLFN\./u, '')
    if (callee === 'IF') {
      return compileNativeIfCandidate(cell, rowValues, node.args, context)
    }
    if (callee === 'IFS') {
      return compileNativeIfsCandidate(cell, rowValues, node.args, context)
    }
  }
  const value = compileScalarValue(node, context)
  return value ? { cell, rowValues, conditions: [], branchValues: [], defaultValue: value } : null
}

function compileNativeIfCandidate(
  cell: NativeFormulaCell,
  rowValues: PendingCellRow,
  args: readonly FormulaNode[],
  context: ConditionalCompileContext,
): ConditionalPickCandidate | null {
  if (args.length < 2 || args.length > 3) {
    return null
  }
  const condition = compileConditionalCondition(args[0]!, context)
  const trueValue = compileScalarValue(args[1]!, context)
  const falseValue = args[2] ? compileScalarValue(args[2], context) : ({ tag: ValueTag.Boolean, value: false } satisfies CellValue)
  return condition && trueValue && falseValue
    ? {
        cell,
        rowValues,
        conditions: [condition],
        branchValues: [trueValue],
        defaultValue: falseValue,
      }
    : null
}

function compileNativeIfsCandidate(
  cell: NativeFormulaCell,
  rowValues: PendingCellRow,
  args: readonly FormulaNode[],
  context: ConditionalCompileContext,
): ConditionalPickCandidate | null {
  if (args.length < 2 || args.length % 2 !== 0) {
    return null
  }
  const conditions: ConditionalPickCondition[] = []
  const branchValues: CellValue[] = []
  for (let index = 0; index < args.length; index += 2) {
    const condition = compileConditionalCondition(args[index]!, context)
    const value = compileScalarValue(args[index + 1]!, context)
    if (!condition || !value) {
      return null
    }
    conditions.push(condition)
    branchValues.push(value)
  }
  return {
    cell,
    rowValues,
    conditions,
    branchValues,
    defaultValue: { tag: ValueTag.Error, code: ErrorCode.NA },
  }
}

function compileConditionalCondition(node: FormulaNode, context: ConditionalCompileContext): ConditionalPickCondition | null {
  if (node.kind === 'BinaryExpr') {
    const operatorCode = directConditionOperatorCode(node.operator)
    if (operatorCode === undefined) {
      return null
    }
    const left = compileScalarValue(node.left, context)
    const right = compileScalarValue(node.right, context)
    if (!left || !right || !conditionOperandsAreSupported(operatorCode, left, right)) {
      return null
    }
    return { operatorCode, left, right }
  }
  const left = compileScalarValue(node, context)
  return left && conditionOperandsAreSupported(directConditionOpTruthy, left, emptyCellValue())
    ? { operatorCode: directConditionOpTruthy, left, right: emptyCellValue() }
    : null
}

function compileScalarValue(node: FormulaNode, context: ConditionalCompileContext): CellValue | null {
  switch (node.kind) {
    case 'NumberLiteral':
      return Number.isFinite(node.value) ? { tag: ValueTag.Number, value: node.value } : null
    case 'BooleanLiteral':
      return { tag: ValueTag.Boolean, value: node.value }
    case 'StringLiteral':
      return stringCellValue(node.value)
    case 'CellRef':
    case 'StructuredRef':
      return readRowLocalScalarValue(node, context)
    case 'ArrayConstant':
    case 'BinaryExpr':
    case 'CallExpr':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'InvokeExpr':
    case 'NameRef':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'UnaryExpr':
      return null
    default:
      return null
  }
}

function readRowLocalScalarValue(
  node: Extract<FormulaNode, { readonly kind: 'CellRef' | 'StructuredRef' }>,
  context: ConditionalCompileContext,
): CellValue | null {
  if (node.kind === 'CellRef') {
    if (node.sheetName && node.sheetName !== context.sheetName) {
      return null
    }
    let address
    try {
      address = decodeCellAddress(node.ref.replaceAll('$', ''))
    } catch {
      return null
    }
    return address.r === context.row ? resolvedCellValue(context.rowValues.get(address.c)) : null
  }
  const column = compileReferenceColumn(node, context.sheetName, context.row, context.formulaColumn, context.tablesBySheet)
  return column === null ? null : resolvedCellValue(context.rowValues.get(column))
}

function directConditionOperatorCode(operator: string): number | undefined {
  switch (operator) {
    case '=':
      return directConditionOpEq
    case '<>':
      return directConditionOpNeq
    case '>':
      return directConditionOpGt
    case '>=':
      return directConditionOpGte
    case '<':
      return directConditionOpLt
    case '<=':
      return directConditionOpLte
    default:
      return undefined
  }
}

function conditionOperandsAreSupported(operatorCode: number, left: CellValue, right: CellValue): boolean {
  if (left.tag === ValueTag.Error || right.tag === ValueTag.Error) {
    return true
  }
  if (operatorCode === directConditionOpTruthy) {
    return left.tag === ValueTag.Number || left.tag === ValueTag.Boolean || left.tag === ValueTag.Empty
  }
  const leftNumeric = isNumericComparable(left)
  const rightNumeric = isNumericComparable(right)
  if (leftNumeric && rightNumeric) {
    return true
  }
  return (
    (operatorCode === directConditionOpEq || operatorCode === directConditionOpNeq) &&
    (left.tag === ValueTag.String || left.tag === ValueTag.Empty) &&
    (right.tag === ValueTag.String || right.tag === ValueTag.Empty)
  )
}

function resolvedCellValue(value: PendingCellValue | undefined): CellValue {
  if (value === undefined || isSharedStringReference(value)) {
    return { tag: ValueTag.Empty }
  }
  return value
}

function emptyCellValue(): CellValue {
  return { tag: ValueTag.Empty }
}

function stringCellValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function isNumericComparable(value: CellValue): boolean {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean || value.tag === ValueTag.Empty
}

type KernelStringMode = 'compare' | 'output'

class KernelStringArena {
  private readonly ids = new Map<string, number>()
  private readonly outputValues = new Map<number, string>()

  idFor(value: string, mode: KernelStringMode): number {
    const key = mode === 'compare' ? value.toLocaleLowerCase('en-US') : value
    if (mode === 'compare' && key.length === 0) {
      return 0
    }
    const existing = this.ids.get(key)
    if (existing !== undefined) {
      if (mode === 'output') {
        this.outputValues.set(existing, value)
      }
      return existing
    }
    const id = this.ids.size + 1
    this.ids.set(key, id)
    if (mode === 'output') {
      this.outputValues.set(id, value)
    }
    return id
  }

  outputValue(id: number): string | undefined {
    return this.outputValues.get(id)
  }
}

function appendKernelValueFields(
  value: CellValue,
  stringArena: KernelStringArena,
  mode: KernelStringMode,
  tags: number[],
  numbers: number[],
  stringIds: number[],
  errors: number[],
): void {
  tags.push(value.tag)
  switch (value.tag) {
    case ValueTag.Number:
      numbers.push(value.value)
      stringIds.push(0)
      errors.push(ErrorCode.None)
      break
    case ValueTag.Boolean:
      numbers.push(value.value ? 1 : 0)
      stringIds.push(0)
      errors.push(ErrorCode.None)
      break
    case ValueTag.String:
      numbers.push(0)
      stringIds.push(stringArena.idFor(value.value, mode))
      errors.push(ErrorCode.None)
      break
    case ValueTag.Error:
      numbers.push(0)
      stringIds.push(0)
      errors.push(value.code)
      break
    case ValueTag.Empty:
      numbers.push(0)
      stringIds.push(0)
      errors.push(ErrorCode.None)
      break
  }
}

function cellValueFromKernelFields(
  tag: number,
  number: number,
  stringId: number,
  error: number,
  stringArena: KernelStringArena,
): CellValue | null {
  if (!isValueTag(tag)) {
    return null
  }
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return Number.isFinite(number) ? { tag: ValueTag.Number, value: number } : null
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: number !== 0 }
    case ValueTag.String: {
      const value = stringArena.outputValue(stringId)
      return value === undefined ? null : stringCellValue(value)
    }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: error as ErrorCode }
    default:
      return null
  }
}

function isValueTag(value: number): value is ValueTag {
  return Number.isInteger(value) && value >= 0 && value <= 4
}

function cellValueToPatchValue(value: CellValue): XlsxSourceLiteralPatch['value'] | undefined {
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
      return undefined
    default:
      return undefined
  }
}

function formulaPatch(cell: NativeFormulaCell, value: XlsxSourceLiteralPatch['value']): XlsxSourceLiteralPatch {
  return {
    sheetName: cell.sheetName,
    address: cell.address,
    value,
    preserveFormula: true,
  }
}

function cellKey(cell: NativeFormulaCell): string {
  return cell.sheetName + '!' + cell.address
}

function emptyResult(): StreamingNativeWasmConditionalResult {
  return {
    batchCount: 0,
    evaluatedFormulaCellCount: 0,
    patches: [],
    processedCells: new Set(),
  }
}

function isSharedStringReference(value: PendingCellValue): value is Extract<PendingCellValue, { readonly kind: 'shared-string' }> {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}
