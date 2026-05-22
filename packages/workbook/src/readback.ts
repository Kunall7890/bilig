import { parseCellAddress } from '@bilig/formula'
import type { LiteralInput } from '@bilig/protocol'
import type { WorkbookRef } from './find.js'
import type { WorkbookCheckProof, WorkbookCheckResult } from './result.js'

export interface WorkbookRunReadback {
  readonly target: WorkbookRef
  readonly value?: LiteralInput
  readonly formula?: string | null
  readonly values?: readonly (readonly LiteralInput[])[]
  readonly formulas?: readonly (readonly (string | null)[])[]
  readonly cells?: readonly WorkbookCellReadback[]
}

export interface WorkbookCellReadback {
  readonly sheetName: string
  readonly address: string
  readonly value?: LiteralInput
  readonly formula?: string | null
}

export type WorkbookReadbackIssueCode =
  | 'readback_missing'
  | 'duplicate_readback'
  | 'value_mismatch'
  | 'values_mismatch'
  | 'formula_mismatch'
  | 'formulas_mismatch'

export const workbookReadbackIssueCodes = Object.freeze([
  'readback_missing',
  'duplicate_readback',
  'value_mismatch',
  'values_mismatch',
  'formula_mismatch',
  'formulas_mismatch',
] satisfies readonly WorkbookReadbackIssueCode[])

export function isWorkbookReadbackIssueCode(value: unknown): value is WorkbookReadbackIssueCode {
  return typeof value === 'string' && workbookReadbackIssueCodes.some((code) => code === value)
}

export interface WorkbookReadbackIssue {
  readonly code: WorkbookReadbackIssueCode
  readonly message: string
  readonly path?: string
  readonly check: WorkbookCheckResult
  readonly target?: WorkbookRef
  readonly expected?: LiteralInput | readonly (readonly LiteralInput[])[] | readonly (readonly (string | null)[])[]
  readonly actual?: LiteralInput | readonly (readonly LiteralInput[])[] | readonly (readonly (string | null)[])[]
}

export interface WorkbookReadbackVerification {
  readonly status: 'passed' | 'failed'
  readonly checks: readonly WorkbookCheckResult[]
  readonly issues: readonly WorkbookReadbackIssue[]
}

function refKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
}

function readbackKey(readback: WorkbookRunReadback): string {
  return refKey(readback.target)
}

function findReadback(readbacks: ReadonlyMap<string, WorkbookRunReadback>, target: WorkbookRef): WorkbookRunReadback | undefined {
  return readbacks.get(refKey(target))
}

function hasValue(readback: WorkbookRunReadback): readback is WorkbookRunReadback & { readonly value: LiteralInput } {
  return Object.prototype.hasOwnProperty.call(readback, 'value')
}

function hasFormula(readback: WorkbookRunReadback): readback is WorkbookRunReadback & { readonly formula: string | null } {
  return Object.prototype.hasOwnProperty.call(readback, 'formula')
}

function hasCellValue(cell: WorkbookCellReadback): cell is WorkbookCellReadback & { readonly value: LiteralInput } {
  return Object.prototype.hasOwnProperty.call(cell, 'value')
}

function hasCellFormula(cell: WorkbookCellReadback): cell is WorkbookCellReadback & { readonly formula: string | null } {
  return Object.prototype.hasOwnProperty.call(cell, 'formula')
}

interface CellPosition {
  readonly row: number
  readonly col: number
}

function cellPosition(sheetName: string, address: string): CellPosition | null {
  try {
    const parsed = parseCellAddress(address)
    if (parsed.sheetName !== undefined && parsed.sheetName !== sheetName) {
      return null
    }
    return {
      row: parsed.row,
      col: parsed.col,
    }
  } catch {
    return null
  }
}

function targetCellRange(readback: WorkbookRunReadback): {
  readonly sheetName: string
  readonly start: CellPosition
  readonly end: CellPosition
} | null {
  if (readback.target.kind !== 'range') {
    return null
  }
  const { sheetName, startAddress, endAddress } = readback.target.range
  const start = cellPosition(sheetName, startAddress)
  const end = cellPosition(sheetName, endAddress)
  if (start === null || end === null || end.row < start.row || end.col < start.col) {
    return null
  }
  return { sheetName, start, end }
}

function cellKey(row: number, col: number): string {
  return `${String(row)}:${String(col)}`
}

function indexedCells(readback: WorkbookRunReadback): ReadonlyMap<string, WorkbookCellReadback> | null {
  if (readback.cells === undefined) {
    return null
  }
  const target = targetCellRange(readback)
  if (target === null) {
    return null
  }
  const cells = new Map<string, WorkbookCellReadback>()
  for (const cell of readback.cells) {
    if (cell.sheetName !== target.sheetName) {
      continue
    }
    const position = cellPosition(cell.sheetName, cell.address)
    if (
      position === null ||
      position.row < target.start.row ||
      position.row > target.end.row ||
      position.col < target.start.col ||
      position.col > target.end.col
    ) {
      continue
    }
    const key = cellKey(position.row, position.col)
    if (cells.has(key)) {
      return null
    }
    cells.set(key, cell)
  }
  return cells
}

function cellValuesMatrix(readback: WorkbookRunReadback): readonly (readonly LiteralInput[])[] | undefined {
  const target = targetCellRange(readback)
  const cells = indexedCells(readback)
  if (target === null || cells === null) {
    return undefined
  }
  const rows: LiteralInput[][] = []
  for (let row = target.start.row; row <= target.end.row; row += 1) {
    const values: LiteralInput[] = []
    for (let col = target.start.col; col <= target.end.col; col += 1) {
      const cell = cells.get(cellKey(row, col))
      if (cell === undefined || !hasCellValue(cell)) {
        return undefined
      }
      values.push(cell.value)
    }
    rows.push(values)
  }
  return rows
}

function cellFormulasMatrix(readback: WorkbookRunReadback): readonly (readonly (string | null)[])[] | undefined {
  const target = targetCellRange(readback)
  const cells = indexedCells(readback)
  if (target === null || cells === null) {
    return undefined
  }
  const rows: (string | null)[][] = []
  for (let row = target.start.row; row <= target.end.row; row += 1) {
    const formulas: (string | null)[] = []
    for (let col = target.start.col; col <= target.end.col; col += 1) {
      const cell = cells.get(cellKey(row, col))
      if (cell === undefined || !hasCellFormula(cell)) {
        return undefined
      }
      formulas.push(cell.formula)
    }
    rows.push(formulas)
  }
  return rows
}

function scalarValue(readback: WorkbookRunReadback): LiteralInput | undefined {
  if (hasValue(readback)) {
    return readback.value
  }
  if (readback.values?.length === 1 && readback.values[0]?.length === 1) {
    return readback.values[0][0]
  }
  const cellValues = cellValuesMatrix(readback)
  if (cellValues?.length === 1 && cellValues[0]?.length === 1) {
    return cellValues[0][0]
  }
  return undefined
}

function scalarFormula(readback: WorkbookRunReadback): string | null | undefined {
  if (hasFormula(readback)) {
    return readback.formula
  }
  if (readback.formulas?.length === 1 && readback.formulas[0]?.length === 1) {
    return readback.formulas[0][0]
  }
  const cellFormulas = cellFormulasMatrix(readback)
  if (cellFormulas?.length === 1 && cellFormulas[0]?.length === 1) {
    return cellFormulas[0][0]
  }
  return undefined
}

function valuesMatrix(readback: WorkbookRunReadback): readonly (readonly LiteralInput[])[] | undefined {
  if (readback.values !== undefined) {
    return readback.values
  }
  if (hasValue(readback)) {
    return [[readback.value]]
  }
  return cellValuesMatrix(readback)
}

function formulasMatrix(readback: WorkbookRunReadback): readonly (readonly (string | null)[])[] | undefined {
  if (readback.formulas !== undefined) {
    return readback.formulas
  }
  if (hasFormula(readback)) {
    return [[readback.formula]]
  }
  return cellFormulasMatrix(readback)
}

function sameMatrix<T>(left: readonly (readonly T[])[], right: readonly (readonly T[])[]): boolean {
  return (
    left.length === right.length &&
    left.every(
      (leftRow, rowIndex) =>
        leftRow.length === right[rowIndex]?.length &&
        leftRow.every((leftValue, colIndex) => Object.is(leftValue, right[rowIndex]?.[colIndex])),
    )
  )
}

function issue(input: WorkbookReadbackIssue): WorkbookReadbackIssue {
  return input
}

function missingReadback(check: WorkbookCheckResult): WorkbookReadbackIssue {
  return issue({
    code: 'readback_missing',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    message: `${check.target?.label ?? check.kind} has no readback`,
  })
}

function duplicateReadback(check: WorkbookCheckResult, readback: WorkbookRunReadback): WorkbookReadbackIssue {
  return issue({
    code: 'duplicate_readback',
    check,
    target: readback.target,
    message: `${readback.target.label} has more than one readback`,
  })
}

function valueMismatch(check: WorkbookCheckResult, expected: LiteralInput, actual?: LiteralInput): WorkbookReadbackIssue {
  return issue({
    code: 'value_mismatch',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.target?.label ?? check.kind} has no value readback`
        : `${check.target?.label ?? check.kind} expected value ${JSON.stringify(expected)} but read ${JSON.stringify(actual)}`,
  })
}

function valuesMismatch(
  check: WorkbookCheckResult,
  expected: readonly (readonly LiteralInput[])[],
  actual?: readonly (readonly LiteralInput[])[],
): WorkbookReadbackIssue {
  return issue({
    code: 'values_mismatch',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.target?.label ?? check.kind} has no values readback`
        : `${check.target?.label ?? check.kind} expected values ${JSON.stringify(expected)} but read ${JSON.stringify(actual)}`,
  })
}

function formulaMismatch(check: WorkbookCheckResult, expected: string, actual: string | null | undefined): WorkbookReadbackIssue {
  return issue({
    code: 'formula_mismatch',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.target?.label ?? check.kind} has no formula readback`
        : `${check.target?.label ?? check.kind} expected formula ${expected} but read ${actual ?? 'null'}`,
  })
}

function formulasMismatch(
  check: WorkbookCheckResult,
  expected: readonly (readonly (string | null)[])[],
  actual: readonly (readonly (string | null)[])[] | undefined,
): WorkbookReadbackIssue {
  return issue({
    code: 'formulas_mismatch',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.target?.label ?? check.kind} has no formulas readback`
        : `${check.target?.label ?? check.kind} expected formulas ${JSON.stringify(expected)} but read ${JSON.stringify(actual)}`,
  })
}

function checked(check: WorkbookCheckResult, status: WorkbookCheckResult['status'], proof?: WorkbookCheckProof): WorkbookCheckResult {
  return {
    ...check,
    status,
    ...(proof !== undefined ? { proof } : {}),
  }
}

function valueProof(value: LiteralInput): WorkbookCheckProof {
  return {
    kind: 'value',
    value,
  }
}

function valuesProof(values: readonly (readonly LiteralInput[])[]): WorkbookCheckProof {
  return {
    kind: 'values',
    values,
  }
}

function formulaProof(formula: string | null): WorkbookCheckProof {
  return {
    kind: 'formula',
    formula,
  }
}

function formulasProof(formulas: readonly (readonly (string | null)[])[]): WorkbookCheckProof {
  return {
    kind: 'formulas',
    formulas,
  }
}

function verifyCheck(
  check: WorkbookCheckResult,
  readbacks: ReadonlyMap<string, WorkbookRunReadback>,
): { readonly check: WorkbookCheckResult; readonly issue?: WorkbookReadbackIssue } {
  if (check.expectation === undefined) {
    return { check }
  }

  if (check.target === undefined) {
    const failedCheck = checked(check, 'failed')
    return {
      check: failedCheck,
      issue: missingReadback(failedCheck),
    }
  }

  const readback = findReadback(readbacks, check.target)
  if (readback === undefined) {
    const failedCheck = checked(check, 'failed')
    return {
      check: failedCheck,
      issue: missingReadback(failedCheck),
    }
  }

  if (check.expectation.kind === 'valueEquals') {
    const actual = scalarValue(readback)
    if (actual !== check.expectation.value) {
      const failedCheck = checked(check, 'failed', actual === undefined ? undefined : valueProof(actual))
      return {
        check: failedCheck,
        issue: valueMismatch(failedCheck, check.expectation.value, actual),
      }
    }
    return { check: checked(check, 'passed', valueProof(actual)) }
  }

  if (check.expectation.kind === 'valuesEqual') {
    const actual = valuesMatrix(readback)
    if (actual === undefined || !sameMatrix(actual, check.expectation.values)) {
      const failedCheck = checked(check, 'failed', actual === undefined ? undefined : valuesProof(actual))
      return {
        check: failedCheck,
        issue: valuesMismatch(failedCheck, check.expectation.values, actual),
      }
    }
    return { check: checked(check, 'passed', valuesProof(actual)) }
  }

  if (check.expectation.kind === 'formulaEquals') {
    const actual = scalarFormula(readback)
    if (actual !== check.expectation.formula) {
      const failedCheck = checked(check, 'failed', actual === undefined ? undefined : formulaProof(actual))
      return {
        check: failedCheck,
        issue: formulaMismatch(failedCheck, check.expectation.formula, actual),
      }
    }
    return { check: checked(check, 'passed', formulaProof(actual)) }
  }

  const actual = formulasMatrix(readback)
  if (actual === undefined || !sameMatrix(actual, check.expectation.formulas)) {
    const failedCheck = checked(check, 'failed', actual === undefined ? undefined : formulasProof(actual))
    return {
      check: failedCheck,
      issue: formulasMismatch(failedCheck, check.expectation.formulas, actual),
    }
  }
  return { check: checked(check, 'passed', formulasProof(actual)) }
}

export function verifyWorkbookReadbacks(
  checks: readonly WorkbookCheckResult[],
  readbacks: readonly WorkbookRunReadback[],
): WorkbookReadbackVerification {
  const readbackByTarget = new Map<string, WorkbookRunReadback>()
  const duplicateReadbacks = new Map<string, WorkbookRunReadback>()
  readbacks.forEach((readback) => {
    const key = readbackKey(readback)
    if (readbackByTarget.has(key)) {
      duplicateReadbacks.set(key, readback)
      return
    }
    readbackByTarget.set(key, readback)
  })

  const verifiedChecks: WorkbookCheckResult[] = []
  const issues: WorkbookReadbackIssue[] = []

  checks.forEach((check, checkIndex) => {
    if (check.target !== undefined) {
      const duplicate = duplicateReadbacks.get(refKey(check.target))
      if (duplicate !== undefined) {
        const failedCheck = checked(check, 'failed')
        verifiedChecks.push(failedCheck)
        issues.push({
          ...duplicateReadback(failedCheck, duplicate),
          path: `checks[${String(checkIndex)}]`,
        })
        return
      }
    }
    const result = verifyCheck(check, readbackByTarget)
    verifiedChecks.push(result.check)
    if (result.issue !== undefined) {
      issues.push({
        ...result.issue,
        path: `checks[${String(checkIndex)}]`,
      })
    }
  })

  return {
    status: issues.length === 0 ? 'passed' : 'failed',
    checks: verifiedChecks,
    issues,
  }
}
