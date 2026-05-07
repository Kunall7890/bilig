import { ValueTag } from '@bilig/protocol'
import { formatAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { getRuntimeFormulaSource } from '../runtime-formula-source.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'

export function shouldCaptureStoredCell(args: CreateEngineStructureServiceArgs, cellIndex: number): boolean {
  const value = args.state.workbook.cellStore.getValue(cellIndex, () => '')
  const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
  const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
  const formula = args.state.formulas.get(cellIndex)
  if ((flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
    return false
  }
  return !(
    formula === undefined &&
    explicitFormat === undefined &&
    (flags & CellFlags.AuthoredBlank) === 0 &&
    (value.tag === ValueTag.Empty || value.tag === ValueTag.Error)
  )
}

function captureStoredCellState(
  args: CreateEngineStructureServiceArgs,
  cellIndex: number,
  sheetName: string,
  address: string,
  sourceSheetName?: string,
  sourceAddress?: string,
): EngineOp[] {
  return args.captureStoredCellOps(cellIndex, sheetName, address, sourceSheetName, sourceAddress)
}

function captureStoredCellStateForUndo(
  args: CreateEngineStructureServiceArgs,
  cellIndex: number,
  sheetName: string,
  address: string,
): EngineOp[] {
  const formula = args.state.formulas.get(cellIndex)
  if (formula) {
    return [{ kind: 'setCellFormula', sheetName, address, formula: getRuntimeFormulaSource(formula) }]
  }
  const tag: ValueTag = args.state.workbook.cellStore.tags[cellIndex] ?? ValueTag.Empty
  const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
  const shouldRestoreExplicitBlank =
    (args.state.workbook.cellStore.versions[cellIndex] ?? 0) !== 0 || (flags & CellFlags.AuthoredBlank) !== 0
  const ops: EngineOp[] = []
  switch (tag) {
    case ValueTag.Number:
      ops.push({ kind: 'setCellValue', sheetName, address, value: args.state.workbook.cellStore.numbers[cellIndex] ?? 0 })
      break
    case ValueTag.Boolean:
      ops.push({ kind: 'setCellValue', sheetName, address, value: (args.state.workbook.cellStore.numbers[cellIndex] ?? 0) !== 0 })
      break
    case ValueTag.String:
      return captureStoredCellState(args, cellIndex, sheetName, address)
    case ValueTag.Empty:
    case ValueTag.Error:
      ops.push(
        shouldRestoreExplicitBlank ? { kind: 'setCellValue', sheetName, address, value: null } : { kind: 'clearCell', sheetName, address },
      )
      break
  }
  const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
  if (explicitFormat !== undefined) {
    ops.push({ kind: 'setCellFormat', sheetName, address, format: explicitFormat ?? null })
  }
  return ops
}

export function captureAxisRangeCellState(
  args: CreateEngineStructureServiceArgs,
  sheetName: string,
  axis: 'row' | 'column',
  start: number,
  count: number,
): EngineOp[] {
  const sheet = args.state.workbook.getSheet(sheetName)
  if (!sheet) {
    return []
  }
  const axisIds = sheet.logicalAxisMap.snapshot(axis, start, count).map((entry) => entry.id)
  const captured: Array<{ cellIndex: number; row: number; col: number }> = []
  sheet.logical.listResidentCellIndicesUnordered(axis, axisIds).forEach((cellIndex) => {
    if (!shouldCaptureStoredCell(args, cellIndex)) {
      return
    }
    const position = args.state.workbook.getCellPosition(cellIndex)
    if (!position) {
      return
    }
    const { row, col } = position
    const index = axis === 'row' ? row : col
    if (index >= start && index < start + count) {
      captured.push({ cellIndex, row, col })
    }
  })
  if (args.state.counters && captured.length > 0) {
    addEngineCounter(args.state.counters, 'structuralUndoCapturedCells', captured.length)
  }
  return captured
    .toSorted((left, right) => left.row - right.row || left.col - right.col)
    .flatMap(({ cellIndex, row, col }) => captureStoredCellStateForUndo(args, cellIndex, sheetName, formatAddress(row, col)))
}

export function captureSheetCellState(args: CreateEngineStructureServiceArgs, sheetName: string): EngineOp[] {
  const sheet = args.state.workbook.getSheet(sheetName)
  if (!sheet) {
    return []
  }
  const captured: Array<{ cellIndex: number; row: number; col: number }> = []
  sheet.grid.forEachCellEntry((cellIndex, row, col) => {
    if (!shouldCaptureStoredCell(args, cellIndex)) {
      return
    }
    captured.push({ cellIndex, row, col })
  })
  return captured
    .toSorted((left, right) => left.row - right.row || left.col - right.col)
    .flatMap(({ cellIndex }) => captureStoredCellState(args, cellIndex, sheetName, args.state.workbook.getAddress(cellIndex)))
}
