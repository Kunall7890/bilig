import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { formatAddress, parseCellAddress, parseRangeAddress } from './addressing.js'
import { parseFormula } from './parser.js'
import type { EvaluationContext, ReferenceOperand, StackValue } from './js-evaluator.js'

interface AggregateCandidateValue {
  value: CellValue
  sheetName?: string
  address?: string
}

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}

function isNestedRollupFormula(context: EvaluationContext, sheetName: string, address: string, callees: ReadonlySet<string>): boolean {
  const source = context.resolveFormula?.(sheetName, address)
  if (!source) {
    return false
  }
  try {
    const ast = parseFormula(source)
    if (ast.kind !== 'CallExpr') {
      return false
    }
    const callee = ast.callee.toUpperCase()
    return callees.has(callee)
  } catch {
    return false
  }
}

export const nestedSubtotalCallees = new Set(['SUBTOTAL'])
export const nestedAggregateCallees = new Set(['SUBTOTAL', 'AGGREGATE'])

export function aggregateOptionIgnoresHiddenRows(option: number): boolean {
  return option === 1 || option === 3 || option === 5 || option === 7
}

export function aggregateOptionIgnoresErrors(option: number): boolean {
  return option === 2 || option === 3 || option === 6 || option === 7
}

export function aggregateOptionIgnoresNestedRollups(option: number): boolean {
  return option >= 0 && option <= 3
}

export function firstErrorValue(values: readonly CellValue[]): CellValue | undefined {
  return values.find((value) => value.tag === ValueTag.Error)
}

export function filterNestedRollupCandidates(
  candidates: readonly AggregateCandidateValue[],
  context: EvaluationContext,
  callees: ReadonlySet<string>,
): AggregateCandidateValue[] {
  return candidates.filter(
    (candidate) =>
      !candidate.sheetName || !candidate.address || !isNestedRollupFormula(context, candidate.sheetName, candidate.address, callees),
  )
}

export function collectAggregateCandidates(
  value: StackValue,
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
  ignoreHiddenRows: boolean,
): AggregateCandidateValue[] {
  if (value.kind === 'scalar') {
    if (ref?.kind !== 'cell' || !ref.address) {
      return [{ value: value.value }]
    }
    const sheetName = ref.sheetName ?? context.sheetName
    const cell = parseCellAddress(ref.address, sheetName)
    if (ignoreHiddenRows && context.isRowHidden?.(sheetName, cell.row) === true) {
      return []
    }
    return [{ value: value.value, sheetName, address: formatAddress(cell.row, cell.col) }]
  }
  if (value.kind === 'omitted' || value.kind === 'lambda') {
    return [{ value: valueError() }]
  }
  if (value.kind === 'array') {
    return value.values.map((cellValue) => ({ value: cellValue }))
  }
  if (value.refKind !== 'cells') {
    return value.values.map((cellValue) => ({ value: cellValue }))
  }

  const sheetName = ref?.sheetName ?? value.sheetName ?? context.sheetName
  const start = ref?.kind === 'range' ? ref.start : value.start
  const end = ref?.kind === 'range' ? ref.end : value.end
  if (!start || !end) {
    return value.values.map((cellValue) => ({ value: cellValue }))
  }

  let startRow = 0
  let startCol = 0
  let cols = value.cols
  try {
    const parsed = parseRangeAddress(`${start}:${end}`, sheetName)
    if (parsed.kind !== 'cells') {
      return value.values.map((cellValue) => ({ value: cellValue }))
    }
    startRow = parsed.start.row
    startCol = parsed.start.col
    cols = parsed.end.col - parsed.start.col + 1
  } catch {
    return value.values.map((cellValue) => ({ value: cellValue }))
  }

  const candidates: AggregateCandidateValue[] = []
  for (let index = 0; index < value.values.length; index += 1) {
    const row = startRow + Math.floor(index / cols)
    if (ignoreHiddenRows && context.isRowHidden?.(sheetName, row) === true) {
      continue
    }
    const col = startCol + (index % cols)
    candidates.push({ value: value.values[index]!, sheetName, address: formatAddress(row, col) })
  }
  return candidates
}
