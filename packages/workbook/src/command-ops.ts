import { formatAddress, parseCellAddress, parseFormula, serializeFormula } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkbookRef, WorkbookRefData } from './find.js'
import { materializeFormulaLabels, type WorkbookFormulaLabelReplacement } from './formula-usage.js'
import type { WorkbookActionCommand } from './model.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookCommandResolvedRefs, WorkbookResolvedRefValue } from './result.js'

export type WorkbookConcreteCommandOp = Extract<WorkbookOp, { kind: 'setCellFormula' | 'setCellValue' | 'setCellFormat' | 'clearCell' }>

interface WorkbookConcreteCell {
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
}

export interface WorkbookCommandOpsExpectation {
  readonly ops: readonly WorkbookOp[]
}

function concreteSingleCell(target: WorkbookRef): { sheetName: string; address: string } | null {
  if (target.kind !== 'range') {
    return null
  }
  const range = target.range
  return range.startAddress === range.endAddress ? { sheetName: range.sheetName, address: range.startAddress } : null
}

function isRangeRefData(value: WorkbookRefData): value is Extract<WorkbookRefData, { readonly kind: 'range' }> {
  return value.kind === 'range'
}

function concreteRangesFromResolvedRef(value: WorkbookResolvedRefValue | undefined): readonly CellRangeRef[] | null {
  if (value === undefined) {
    return null
  }
  const values = Array.isArray(value) ? value : [value]
  if (!values.every(isRangeRefData)) {
    return null
  }
  return values.map((entry) => entry.range)
}

function cellsFromRange(range: CellRangeRef): readonly WorkbookConcreteCell[] {
  const start = parseCellAddress(range.startAddress)
  const end = parseCellAddress(range.endAddress)
  if (end.row < start.row || end.col < start.col) {
    return []
  }
  const cells: WorkbookConcreteCell[] = []
  for (let row = start.row; row <= end.row; row += 1) {
    for (let col = start.col; col <= end.col; col += 1) {
      cells.push({
        sheetName: range.sheetName,
        address: formatAddress(row, col),
        row,
        col,
      })
    }
  }
  return cells
}

function concreteCellsFromResolvedRefs(resolvedRefs: WorkbookCommandResolvedRefs | undefined): readonly WorkbookConcreteCell[] | null {
  const ranges = concreteRangesFromResolvedRef(resolvedRefs?.target)
  if (ranges === null) {
    return null
  }
  return ranges.flatMap(cellsFromRange)
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function cellSource(cell: WorkbookConcreteCell): string {
  return `${quoteSheetName(cell.sheetName)}!${cell.address}`
}

function rangeSource(range: CellRangeRef): string {
  const sheet = quoteSheetName(range.sheetName)
  return range.startAddress === range.endAddress ? `${sheet}!${range.startAddress}` : `${sheet}!${range.startAddress}:${range.endAddress}`
}

function workbookRefKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
}

function resolvedInputForLabel(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  label: { readonly ref: WorkbookRef },
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
): WorkbookResolvedRefValue | null {
  const inputs = resolvedRefs?.inputs
  if (inputs === undefined) {
    return null
  }
  const labelKey = workbookRefKey(label.ref)
  const inputIndex = command.inputs.findIndex((input) => workbookRefKey(input) === labelKey)
  return inputIndex < 0 ? null : (inputs[inputIndex] ?? null)
}

function labelReplacementForCell(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  label: { readonly name: string; readonly ref: WorkbookRef },
  targetCells: readonly WorkbookConcreteCell[],
  cellIndex: number,
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
): WorkbookFormulaLabelReplacement | null {
  const value = resolvedInputForLabel(command, label, resolvedRefs)
  const inputRanges = concreteRangesFromResolvedRef(value ?? undefined)
  if (inputRanges === null) {
    return null
  }
  const inputCells = inputRanges.flatMap(cellsFromRange)
  if (inputCells.length === targetCells.length) {
    const inputCell = inputCells[cellIndex]
    return inputCell === undefined ? null : { name: label.name, source: cellSource(inputCell) }
  }
  if (inputCells.length === 1) {
    return { name: label.name, source: cellSource(inputCells[0]!) }
  }
  if (inputRanges.length === 1) {
    return { name: label.name, source: rangeSource(inputRanges[0]!) }
  }
  return null
}

function resolvedFormulaLabelsForCell(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  targetCells: readonly WorkbookConcreteCell[],
  cellIndex: number,
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
): readonly WorkbookFormulaLabelReplacement[] | null {
  const replacements: WorkbookFormulaLabelReplacement[] = []
  for (const label of command.labels) {
    const replacement = labelReplacementForCell(command, label, targetCells, cellIndex, resolvedRefs)
    if (replacement === null) {
      return null
    }
    replacements.push(replacement)
  }
  return replacements
}

function formulaForExpectedCell(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
  targetCells: readonly WorkbookConcreteCell[],
  cellIndex: number,
  resolvedRefs: WorkbookCommandResolvedRefs | undefined,
  formulaLabels: readonly WorkbookFormulaLabelReplacement[] = [],
): string | null {
  if (command.labels.length === 0) {
    return command.inputs.length > 0 ? null : command.formula
  }

  const resolvedLabels = resolvedFormulaLabelsForCell(command, targetCells, cellIndex, resolvedRefs)
  if (resolvedLabels !== null) {
    return materializeFormulaLabels(command.formula, resolvedLabels)
  }
  if (targetCells.length === 1 && formulaLabels.length > 0) {
    return materializeFormulaLabels(command.formula, formulaLabels)
  }
  return null
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalValue)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalValue(entry)]),
    )
  }
  return value
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

export function workbookOpMatches(expected: WorkbookOp, actual: WorkbookOp): boolean {
  return canonicalJson(actual) === canonicalJson(expected)
}

export function workbookOpsMatch(left: readonly WorkbookOp[], right: readonly WorkbookOp[]): boolean {
  return (
    left.length === right.length &&
    left.every((op, index) => {
      const other = right[index]
      return other !== undefined && workbookOpMatches(op, other)
    })
  )
}

export function workbookConcreteOpMatches(expected: WorkbookConcreteCommandOp, actual: WorkbookOp): boolean {
  if (expected.kind !== actual.kind) {
    return false
  }
  switch (expected.kind) {
    case 'setCellFormula':
      return (
        actual.kind === 'setCellFormula' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.formula === expected.formula
      )
    case 'setCellValue':
      return (
        actual.kind === 'setCellValue' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.value === expected.value
      )
    case 'setCellFormat':
      return (
        actual.kind === 'setCellFormat' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.format === expected.format
      )
    case 'clearCell':
      return actual.kind === 'clearCell' && actual.sheetName === expected.sheetName && actual.address === expected.address
  }
}

export function expectedConcreteCommandOp(command: WorkbookActionCommand): WorkbookConcreteCommandOp | null {
  if (command.kind === 'op') {
    return null
  }

  const target = concreteSingleCell(command.target)
  if (target === null) {
    return null
  }
  switch (command.kind) {
    case 'writeFormula':
      return {
        kind: 'setCellFormula',
        sheetName: target.sheetName,
        address: target.address,
        formula: command.formula,
      }
    case 'writeValue':
      return {
        kind: 'setCellValue',
        sheetName: target.sheetName,
        address: target.address,
        value: command.value,
      }
    case 'clear':
      return {
        kind: 'clearCell',
        sheetName: target.sheetName,
        address: target.address,
      }
    case 'format':
      if (command.numberFormat === undefined) {
        return null
      }
      return {
        kind: 'setCellFormat',
        sheetName: target.sheetName,
        address: target.address,
        format: command.numberFormat,
      }
  }
}

export function expectedCommandOps(command: WorkbookActionCommand): WorkbookCommandOpsExpectation | null {
  if (command.kind === 'op') {
    return {
      ops: [command.op],
    }
  }

  const op = expectedConcreteCommandOp(command)
  return op === null
    ? null
    : {
        ops: [op],
      }
}

function formulasMatchExpected(expected: string, actual: string, formulaLabels: readonly WorkbookFormulaLabelReplacement[] = []): boolean {
  try {
    const expectedSource = formulaLabels.length > 0 ? materializeFormulaLabels(expected, formulaLabels) : expected
    return serializeFormula(parseFormula(expectedSource)) === serializeFormula(parseFormula(actual))
  } catch {
    return expected === actual
  }
}

export function commandOpsMatchExpected(
  command: WorkbookActionCommand,
  ops: readonly WorkbookOp[],
  formulaLabels: readonly WorkbookFormulaLabelReplacement[] = [],
  resolvedRefs?: WorkbookCommandResolvedRefs,
): boolean {
  if (ops.length === 0) {
    return true
  }
  if (command.kind === 'writeFormula') {
    const expected = expectedConcreteCommandOp(command)
    const concreteCells = expected === null ? concreteCellsFromResolvedRefs(resolvedRefs) : null
    if (expected === null && concreteCells === null) {
      return false
    }
    const expectedCells =
      concreteCells ??
      (expected === null
        ? []
        : [
            {
              sheetName: expected.sheetName,
              address: expected.address,
              ...parseCellAddress(expected.address),
            },
          ])
    return (
      expectedCells.length === ops.length &&
      expectedCells.every((cell, index) => {
        const actual = ops[index]
        const expectedFormula = formulaForExpectedCell(command, expectedCells, index, resolvedRefs, formulaLabels)
        return (
          actual !== undefined &&
          actual.kind === 'setCellFormula' &&
          actual.sheetName === cell.sheetName &&
          actual.address === cell.address &&
          expectedFormula !== null &&
          formulasMatchExpected(expectedFormula, actual.formula)
        )
      })
    )
  }
  const expected = expectedCommandOps(command)
  if (expected === null) {
    const concreteCells = concreteCellsFromResolvedRefs(resolvedRefs)
    if (concreteCells === null) {
      return false
    }
    if (command.kind === 'writeValue') {
      return (
        concreteCells.length === ops.length &&
        concreteCells.every((cell, index) => {
          const actual = ops[index]
          return (
            actual !== undefined &&
            actual.kind === 'setCellValue' &&
            actual.sheetName === cell.sheetName &&
            actual.address === cell.address &&
            actual.value === command.value
          )
        })
      )
    }
    if (command.kind === 'clear') {
      return (
        concreteCells.length === ops.length &&
        concreteCells.every((cell, index) => {
          const actual = ops[index]
          return (
            actual !== undefined && actual.kind === 'clearCell' && actual.sheetName === cell.sheetName && actual.address === cell.address
          )
        })
      )
    }
    if (command.kind === 'format') {
      return (command.style !== undefined || command.numberFormat !== undefined) && ops.length > 0
    }
    return false
  }
  return workbookOpsMatch(expected.ops, ops)
}
