import type { WorkbookRef } from './find.js'
import type { WorkbookActionCommand } from './model.js'
import type { WorkbookOp } from './ops.js'

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

export function commandOpsMatchExpected(command: WorkbookActionCommand, ops: readonly WorkbookOp[]): boolean {
  if (ops.length === 0) {
    return true
  }
  const expected = expectedCommandOps(command)
  if (expected === null) {
    return true
  }
  return workbookOpsMatch(expected.ops, ops)
}
