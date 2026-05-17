import type { EngineOpBatch } from '@bilig/workbook'
import type { EngineEvent, LiteralInput } from '@bilig/protocol'
import type { SheetRecord } from '../../workbook-store.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import {
  canFastPathOperationLiteralOverwrite,
  cellsShareOperationVersionColumn,
  isOperationClearCellNoOp,
  isOperationNullLiteralWriteNoOp,
  writeFastPathOperationLiteralToExistingCell,
  writeOperationNumericLiteralToExistingCell,
  writeTrustedOperationExistingNumericLiteralToCell,
} from './operation-literal-write-helpers.js'

export function createOperationServiceRuntimeHelpers(args: {
  readonly state: Pick<EngineRuntimeState, 'batchListeners' | 'getLastMetrics' | 'workbook' | 'strings' | 'formulas'>
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
}): {
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => EngineEvent['metrics']
  readonly writeNumericLiteralToExistingCell: (cellIndex: number, value: number) => void
  readonly writeTrustedExistingNumericLiteralToCell: (cellIndex: number, sheet: SheetRecord, col: number, value: number) => void
  readonly writeFastPathLiteralToExistingCell: (cellIndex: number, value: LiteralInput) => void
  readonly cellsShareVersionColumn: (leftCellIndex: number, rightCellIndex: number) => boolean
  readonly withOptionalColumnVersionBatch: (shouldBatch: boolean, execute: () => void) => void
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly isNullLiteralWriteNoOp: (cellIndex: number) => boolean
  readonly isClearCellNoOp: (cellIndex: number) => boolean
} {
  const singleCellKernelSync = new Uint32Array(1)
  return {
    emitBatch(batch) {
      args.state.batchListeners.forEach((listener) => {
        listener(batch)
      })
    },
    deferSingleCellKernelSync(cellIndex) {
      singleCellKernelSync[0] = cellIndex
      args.deferKernelSync(singleCellKernelSync)
    },
    makeSingleLiteralSkipMetrics() {
      const previousMetrics = args.state.getLastMetrics()
      return {
        batchId: previousMetrics.batchId + 1,
        changedInputCount: 1,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      }
    },
    writeNumericLiteralToExistingCell(cellIndex, value) {
      writeOperationNumericLiteralToExistingCell({ workbook: args.state.workbook, cellIndex, value })
    },
    writeTrustedExistingNumericLiteralToCell(cellIndex, sheet, col, value) {
      writeTrustedOperationExistingNumericLiteralToCell({ cellStore: args.state.workbook.cellStore, cellIndex, sheet, col, value })
    },
    writeFastPathLiteralToExistingCell(cellIndex, value) {
      writeFastPathOperationLiteralToExistingCell({ workbook: args.state.workbook, strings: args.state.strings, cellIndex, value })
    },
    cellsShareVersionColumn(leftCellIndex, rightCellIndex) {
      return cellsShareOperationVersionColumn({ workbook: args.state.workbook, leftCellIndex, rightCellIndex })
    },
    withOptionalColumnVersionBatch(shouldBatch, execute) {
      if (shouldBatch) {
        args.state.workbook.withBatchedColumnVersionUpdates(execute)
        return
      }
      execute()
    },
    canFastPathLiteralOverwrite(cellIndex) {
      return canFastPathOperationLiteralOverwrite({ cellStore: args.state.workbook.cellStore, formulas: args.state.formulas, cellIndex })
    },
    isNullLiteralWriteNoOp(cellIndex) {
      return isOperationNullLiteralWriteNoOp({ state: args.state, cellIndex })
    },
    isClearCellNoOp(cellIndex) {
      return isOperationClearCellNoOp({ state: args.state, cellIndex })
    },
  }
}
