import { expect } from 'vitest'
import * as fc from 'fast-check'
import { formatAddress } from '@bilig/formula'
import type { CellNumberFormatInput, CellRangeRef, CellStylePatch, LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

export type CoreCorrectnessAction =
  | { kind: 'values'; range: CellRangeRef; values: LiteralInput[][] }
  | { kind: 'formula'; address: string; formula: string }
  | { kind: 'style'; range: CellRangeRef; patch: CellStylePatch }
  | { kind: 'format'; range: CellRangeRef; format: CellNumberFormatInput }
  | { kind: 'clear'; range: CellRangeRef }
  | { kind: 'fill'; source: CellRangeRef; target: CellRangeRef }
  | { kind: 'insertRows'; start: number; count: number }
  | { kind: 'deleteRows'; start: number; count: number }
  | { kind: 'insertColumns'; start: number; count: number }
  | { kind: 'deleteColumns'; start: number; count: number }

export const sheetName = 'Sheet1'

export function toRangeRef(startRow: number, startCol: number, endRow: number, endCol: number): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  }
}

function buildValueMatrix(height: number, width: number, values: readonly LiteralInput[]): LiteralInput[][] {
  const rows: LiteralInput[][] = []
  let offset = 0
  for (let row = 0; row < height; row += 1) {
    const nextRow: LiteralInput[] = []
    for (let col = 0; col < width; col += 1) {
      nextRow.push(values[offset] ?? null)
      offset += 1
    }
    rows.push(nextRow)
  }
  return rows
}

export function assertSnapshotInvariants(snapshot: WorkbookSnapshot): void {
  const sheetNames = snapshot.sheets.map((sheet) => sheet.name)
  expect(new Set(sheetNames).size).toBe(sheetNames.length)
  snapshot.sheets.forEach((sheet) => {
    const addresses = sheet.cells.map((cell) => cell.address)
    expect(new Set(addresses).size).toBe(addresses.length)
  })
}

export async function createBaselineSnapshot(workbookName: string): Promise<WorkbookSnapshot> {
  const seed = new SpreadsheetEngine({
    workbookName,
    replicaId: `${workbookName}-seed`,
  })
  await seed.ready()
  seed.createSheet(sheetName)
  return seed.exportSnapshot()
}

export function applyAction(engine: SpreadsheetEngine, action: CoreCorrectnessAction): void {
  switch (action.kind) {
    case 'values':
      engine.setRangeValues(action.range, action.values)
      break
    case 'formula':
      engine.setCellFormula(sheetName, action.address, action.formula)
      break
    case 'style':
      engine.setRangeStyle(action.range, action.patch)
      break
    case 'format':
      engine.setRangeNumberFormat(action.range, action.format)
      break
    case 'clear':
      engine.clearRange(action.range)
      break
    case 'fill':
      engine.fillRange(action.source, action.target)
      break
    case 'insertRows':
      engine.insertRows(sheetName, action.start, action.count)
      break
    case 'deleteRows':
      engine.deleteRows(sheetName, action.start, action.count)
      break
    case 'insertColumns':
      engine.insertColumns(sheetName, action.start, action.count)
      break
    case 'deleteColumns':
      engine.deleteColumns(sheetName, action.start, action.count)
      break
  }
}

export function undoAll(engine: SpreadsheetEngine, maxSteps: number): number {
  let steps = 0
  while (engine.undo()) {
    steps += 1
    if (steps > maxSteps) {
      throw new Error(`Undo exceeded expected history budget: ${steps} > ${maxSteps}`)
    }
  }
  return steps
}

export function redoAll(engine: SpreadsheetEngine, maxSteps: number): number {
  let steps = 0
  while (engine.redo()) {
    steps += 1
    if (steps > maxSteps) {
      throw new Error(`Redo exceeded expected history budget: ${steps} > ${maxSteps}`)
    }
  }
  return steps
}

const literalInputArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -10_000, max: 10_000 }),
  fc.boolean(),
  fc.constantFrom('north', 'south', 'ready', 'done'),
  fc.constant(null),
)
const rangeSeedArbitrary = fc.record({
  startRow: fc.integer({ min: 0, max: 5 }),
  startCol: fc.integer({ min: 0, max: 5 }),
  height: fc.integer({ min: 1, max: 2 }),
  width: fc.integer({ min: 1, max: 2 }),
})
const rangeArbitrary = rangeSeedArbitrary.map((value) =>
  toRangeRef(value.startRow, value.startCol, value.startRow + value.height - 1, value.startCol + value.width - 1),
)
const formulaArbitrary = fc
  .tuple(fc.constantFrom('A1', 'B2', 'C3', 'D4', 'E5'), fc.constantFrom('+', '-', '*', '/'), fc.constantFrom('A1', 'B2', 'C3', 'D4', 'E5'))
  .map(([left, operator, right]) => `${left}${operator}${right}`)
const stylePatchArbitrary = fc.constantFrom<CellStylePatch>(
  { fill: { backgroundColor: '#dbeafe' } },
  { font: { bold: true } },
  { alignment: { horizontal: 'right', wrap: true } },
)
const formatInputArbitrary = fc.constantFrom<CellNumberFormatInput>(
  '0.00',
  { kind: 'currency', currency: 'USD', decimals: 2 },
  { kind: 'percent', decimals: 1 },
  { kind: 'text' },
)
const valuesActionArbitrary = rangeSeedArbitrary.chain((range) =>
  fc
    .array(literalInputArbitrary, {
      minLength: range.height * range.width,
      maxLength: range.height * range.width,
    })
    .map((values) => ({
      kind: 'values' as const,
      range: toRangeRef(range.startRow, range.startCol, range.startRow + range.height - 1, range.startCol + range.width - 1),
      values: buildValueMatrix(range.height, range.width, values),
    })),
)
const formulaActionArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 5 }),
    col: fc.integer({ min: 0, max: 5 }),
    formula: formulaArbitrary,
  })
  .map(({ row, col, formula }) => ({
    kind: 'formula' as const,
    address: formatAddress(row, col),
    formula,
  }))
const styleActionArbitrary = fc
  .record({ range: rangeArbitrary, patch: stylePatchArbitrary })
  .map(({ range, patch }) => ({ kind: 'style' as const, range, patch }))
const formatActionArbitrary = fc
  .record({ range: rangeArbitrary, format: formatInputArbitrary })
  .map(({ range, format }) => ({ kind: 'format' as const, range, format }))
const clearActionArbitrary = rangeArbitrary.map((range) => ({ kind: 'clear' as const, range }))
const fillActionArbitrary = rangeSeedArbitrary.chain((source) =>
  fc
    .record({
      targetStartRow: fc.integer({ min: source.startRow, max: 5 }),
      targetStartCol: fc.integer({ min: source.startCol, max: 5 }),
    })
    .map(({ targetStartRow, targetStartCol }) => ({
      kind: 'fill' as const,
      source: toRangeRef(source.startRow, source.startCol, source.startRow + source.height - 1, source.startCol + source.width - 1),
      target: toRangeRef(
        targetStartRow,
        targetStartCol,
        Math.min(5, targetStartRow + source.height - 1),
        Math.min(5, targetStartCol + source.width - 1),
      ),
    })),
)
const axisMutationArbitrary = fc.record({
  start: fc.integer({ min: 0, max: 4 }),
  count: fc.integer({ min: 1, max: 2 }),
})
const insertRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'insertRows' as const,
  start,
  count,
}))
const deleteRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'deleteRows' as const,
  start,
  count,
}))
const insertColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'insertColumns' as const,
  start,
  count,
}))
const deleteColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: 'deleteColumns' as const,
  start,
  count,
}))

export const correctnessActionArbitrary = fc.oneof<CoreCorrectnessAction>(
  valuesActionArbitrary,
  formulaActionArbitrary,
  styleActionArbitrary,
  formatActionArbitrary,
  clearActionArbitrary,
  fillActionArbitrary,
  insertRowsActionArbitrary,
  deleteRowsActionArbitrary,
  insertColumnsActionArbitrary,
  deleteColumnsActionArbitrary,
)
