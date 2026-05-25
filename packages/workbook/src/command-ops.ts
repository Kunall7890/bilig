import { formatAddress, parseCellAddress, parseFormula, serializeFormula } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkbookRef, WorkbookRefData } from './find.js'
import { materializeFormulaLabels, type WorkbookFormulaLabelReplacement } from './formula-usage.js'
import type { WorkbookActionCommand } from './model.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookCommandResolvedRefs, WorkbookResolvedRefValue } from './result.js'

export type WorkbookConcreteCommandOp = Extract<WorkbookOp, { kind: 'setCellFormula' | 'setCellValue' | 'setCellFormat' | 'clearCell' }>

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

function cellsFromRange(range: CellRangeRef): readonly { readonly sheetName: string; readonly address: string }[] {
  const start = parseCellAddress(range.startAddress)
  const end = parseCellAddress(range.endAddress)
  if (end.row < start.row || end.col < start.col) {
    return []
  }
  const cells: { sheetName: string; address: string }[] = []
  for (let row = start.row; row <= end.row; row += 1) {
    for (let col = start.col; col <= end.col; col += 1) {
      cells.push({
        sheetName: range.sheetName,
        address: formatAddress(row, col),
      })
    }
  }
  return cells
}

function concreteCellsFromResolvedRefs(resolvedRefs: WorkbookCommandResolvedRefs | undefined):
  | readonly {
      readonly sheetName: string
      readonly address: string
    }[]
  | null {
  const ranges = concreteRangesFromResolvedRef(resolvedRefs?.target)
  if (ranges === null) {
    return null
  }
  return ranges.flatMap(cellsFromRange)
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
    const expectedCells = concreteCells ?? (expected === null ? [] : [{ sheetName: expected.sheetName, address: expected.address }])
    return (
      expectedCells.length === ops.length &&
      expectedCells.every((cell, index) => {
        const actual = ops[index]
        const formulaMatches =
          expectedCells.length === 1
            ? actual?.kind === 'setCellFormula' && formulasMatchExpected(command.formula, actual.formula, formulaLabels)
            : true
        return (
          actual !== undefined &&
          actual.kind === 'setCellFormula' &&
          actual.sheetName === cell.sheetName &&
          actual.address === cell.address &&
          formulaMatches
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
