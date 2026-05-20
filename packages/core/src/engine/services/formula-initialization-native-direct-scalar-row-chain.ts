import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type {
  EngineRuntimeState,
  RuntimeDirectScalarDescriptor,
  RuntimeDirectScalarOperand,
  RuntimeFormula,
  U32,
} from '../runtime-state.js'
import {
  directScalarCellNumber,
  rowPairDirectScalarCode,
  rowPairDirectScalarCodeNeedsZeroGuard,
  singleInputAffineDirectScalar,
} from './direct-scalar-helpers.js'

// Below the direct-init ceiling, the JS direct path is faster than the JS/WASM boundary cost.
export const MIN_INITIAL_NATIVE_DIRECT_SCALAR_ROW_CHAIN_BATCH_SIZE = 16_385

interface NativeDirectScalarRowChainState {
  readonly workbook: EngineRuntimeState['workbook']
  readonly strings: EngineRuntimeState['strings']
  readonly wasm: EngineRuntimeState['wasm']
  readonly counters: EngineRuntimeState['counters']
}

export interface InitialNativeDirectScalarRowChainBatch {
  readonly count: number
  readonly add: (
    prepared: { readonly cellIndex: number; readonly sheetId: number; readonly row: number; readonly col: number },
    formula: RuntimeFormula | undefined,
  ) => boolean
  readonly evaluate: () => U32 | undefined
}

interface PendingFirstFormula {
  readonly rowOffset: number
  readonly row: number
  readonly cellIndex: number
  readonly leftValue: number
  readonly rightValue: number
}

function operandCellIndex(operand: RuntimeDirectScalarOperand): number | undefined {
  return operand.kind === 'cell' ? operand.cellIndex : undefined
}

function directScalarCellOperands(directScalar: RuntimeDirectScalarDescriptor): readonly [number, number] | undefined {
  if (directScalar.kind !== 'binary' || directScalar.resultOffset !== undefined) {
    return undefined
  }
  const left = operandCellIndex(directScalar.left)
  const right = operandCellIndex(directScalar.right)
  return left === undefined || right === undefined ? undefined : [left, right]
}

function sameRowCell(
  state: NativeDirectScalarRowChainState,
  cellIndex: number,
  sheetId: number,
  row: number,
  expectedCol: number | undefined,
): boolean {
  const cellStore = state.workbook.cellStore
  return (
    cellStore.sheetIds[cellIndex] === sheetId &&
    cellStore.rows[cellIndex] === row &&
    (expectedCol === undefined || cellStore.cols[cellIndex] === expectedCol)
  )
}

export function createInitialNativeDirectScalarRowChainBatch(args: {
  readonly state: NativeDirectScalarRowChainState
  readonly capacity: number
}): InitialNativeDirectScalarRowChainBatch {
  const rowCapacity = args.capacity / 2
  const leftValues = new Float64Array(rowCapacity)
  const rightValues = new Float64Array(rowCapacity)
  const firstTargets = new Uint32Array(rowCapacity)
  const secondTargets = new Uint32Array(rowCapacity)
  const changedCellIndices = new Uint32Array(args.capacity)
  const cellStore = args.state.workbook.cellStore
  let count = 0
  let pendingFirst: PendingFirstFormula | undefined
  let expectedSheetId: number | undefined
  let firstFormulaCol: number | undefined
  let secondFormulaCol: number | undefined
  let leftInputCol: number | undefined
  let rightInputCol: number | undefined
  let firstFormulaCode = 0
  let secondFormulaScale: number | undefined
  let secondFormulaOffset: number | undefined
  let previousCompletedRow = -1
  let failed = args.capacity < MIN_INITIAL_NATIVE_DIRECT_SCALAR_ROW_CHAIN_BATCH_SIZE || args.capacity % 2 !== 0

  const addFirst = (
    prepared: { readonly cellIndex: number; readonly sheetId: number; readonly row: number; readonly col: number },
    formula: RuntimeFormula,
  ): boolean => {
    const operands = directScalarCellOperands(formula.directScalar!)
    if (!operands || pendingFirst !== undefined || prepared.row <= previousCompletedRow) {
      return false
    }
    const [leftCellIndex, rightCellIndex] = operands
    const nextFirstFormulaCode = rowPairDirectScalarCode(formula.directScalar!, leftCellIndex, rightCellIndex)
    if (nextFirstFormulaCode === 0 || rowPairDirectScalarCodeNeedsZeroGuard(nextFirstFormulaCode)) {
      return false
    }
    const nextLeftInputCol = cellStore.cols[leftCellIndex]
    const nextRightInputCol = cellStore.cols[rightCellIndex]
    if (
      nextLeftInputCol === undefined ||
      nextRightInputCol === undefined ||
      !sameRowCell(args.state, leftCellIndex, prepared.sheetId, prepared.row, nextLeftInputCol) ||
      !sameRowCell(args.state, rightCellIndex, prepared.sheetId, prepared.row, nextRightInputCol)
    ) {
      return false
    }
    if (expectedSheetId === undefined) {
      expectedSheetId = prepared.sheetId
      firstFormulaCol = prepared.col
      leftInputCol = nextLeftInputCol
      rightInputCol = nextRightInputCol
      firstFormulaCode = nextFirstFormulaCode
    } else if (
      prepared.sheetId !== expectedSheetId ||
      prepared.col !== firstFormulaCol ||
      nextLeftInputCol !== leftInputCol ||
      nextRightInputCol !== rightInputCol ||
      nextFirstFormulaCode !== firstFormulaCode
    ) {
      return false
    }
    const leftValue = directScalarCellNumber(cellStore, leftCellIndex)
    const rightValue = directScalarCellNumber(cellStore, rightCellIndex)
    if (leftValue === undefined || rightValue === undefined) {
      return false
    }
    pendingFirst = {
      rowOffset: count / 2,
      row: prepared.row,
      cellIndex: prepared.cellIndex,
      leftValue,
      rightValue,
    }
    return true
  }

  const addSecond = (
    prepared: { readonly cellIndex: number; readonly sheetId: number; readonly row: number; readonly col: number },
    formula: RuntimeFormula,
  ): boolean => {
    if (
      pendingFirst === undefined ||
      expectedSheetId === undefined ||
      prepared.sheetId !== expectedSheetId ||
      prepared.row !== pendingFirst.row ||
      prepared.col <= (firstFormulaCol ?? -1)
    ) {
      return false
    }
    const affine = singleInputAffineDirectScalar(formula.directScalar!, pendingFirst.cellIndex)
    if (!affine) {
      return false
    }
    if (secondFormulaCol === undefined) {
      secondFormulaCol = prepared.col
      secondFormulaScale = affine.scale
      secondFormulaOffset = affine.offset
    } else if (prepared.col !== secondFormulaCol || affine.scale !== secondFormulaScale || affine.offset !== secondFormulaOffset) {
      return false
    }
    const rowOffset = pendingFirst.rowOffset
    leftValues[rowOffset] = pendingFirst.leftValue
    rightValues[rowOffset] = pendingFirst.rightValue
    firstTargets[rowOffset] = pendingFirst.cellIndex
    secondTargets[rowOffset] = prepared.cellIndex
    changedCellIndices[rowOffset * 2] = pendingFirst.cellIndex
    changedCellIndices[rowOffset * 2 + 1] = prepared.cellIndex
    previousCompletedRow = prepared.row
    pendingFirst = undefined
    return true
  }

  return {
    get count() {
      return count
    },
    add(prepared, formula) {
      if (failed || count >= args.capacity || !formula?.directScalar || formula.compiled.volatile || formula.compiled.producesSpill) {
        failed = true
        return false
      }
      const added = count % 2 === 0 ? addFirst(prepared, formula) : addSecond(prepared, formula)
      if (!added) {
        failed = true
        return false
      }
      count += 1
      return true
    },
    evaluate() {
      if (
        failed ||
        count !== args.capacity ||
        pendingFirst !== undefined ||
        expectedSheetId === undefined ||
        firstFormulaCol === undefined ||
        secondFormulaCol === undefined ||
        secondFormulaScale === undefined ||
        secondFormulaOffset === undefined ||
        !args.state.wasm.initSyncIfPossible()
      ) {
        return undefined
      }
      const rowCount = count / 2
      if (
        !args.state.wasm.evalDenseDirectScalarRowChainStoreTargetBatch(
          {
            leftValues: leftValues.subarray(0, rowCount),
            rightValues: rightValues.subarray(0, rowCount),
            firstTargets: firstTargets.subarray(0, rowCount),
            secondTargets: secondTargets.subarray(0, rowCount),
            rowCount,
            firstFormulaCode,
            secondFormulaScale,
            secondFormulaOffset,
          },
          args.state.workbook.cellStore.size,
        )
      ) {
        return undefined
      }

      const notifySheetId = expectedSheetId
      const notifyFirstFormulaCol = firstFormulaCol
      const notifySecondFormulaCol = secondFormulaCol
      const clearFormulaOutputFlags = ~(CellFlags.SpillChild | CellFlags.PivotOutput)
      const changedCells = changedCellIndices.subarray(0, count)
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        for (let index = 0; index < changedCells.length; index += 1) {
          const cellIndex = changedCells[index]!
          cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & clearFormulaOutputFlags
        }
        args.state.wasm.syncToStore(cellStore, changedCells, args.state.strings)
        args.state.workbook.notifyColumnsWritten(notifySheetId, Uint32Array.of(notifyFirstFormulaCol, notifySecondFormulaCol))
      })
      addEngineCounter(args.state.counters, 'nativeDirectScalarInitialEvaluations', count)
      return changedCells
    },
  }
}
