import { CellFlags } from '../../cell-store.js'
import type { EngineRuntimeState, RuntimeFormula, U32 } from '../runtime-state.js'
import type { InitialFormulaCellIndexList } from './formula-initialization-refs.js'

export interface InitialDirectScalarPreEvaluationTracker {
  readonly cellCount: number
  readonly cellIndices: U32 | undefined
  readonly allCellsReusable: boolean
  readonly noteCell: (cellIndex: number, runtimeFormula: RuntimeFormula) => void
  readonly noteReusableCells: (cellIndices: U32) => void
}

export function createInitialDirectScalarPreEvaluationTracker(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly refsLength: number
  readonly targetCellIndices: U32
  readonly maxTargetCellIndex: number
  readonly orderedPreparedCellList: () => InitialFormulaCellIndexList
  readonly orderedPreparedCellCount: () => number
}): InitialDirectScalarPreEvaluationTracker {
  let cellBuffer: Uint32Array | undefined
  let cellCount = 0
  let reusableCells: Uint8Array | undefined
  let allCellsReusable = true

  const materializeCellBuffer = (): Uint32Array => {
    if (cellBuffer) {
      return cellBuffer
    }
    cellBuffer = new Uint32Array(Math.max(args.refsLength, 1))
    cellBuffer.set(args.targetCellIndices.subarray(0, cellCount))
    return cellBuffer
  }
  const pushCell = (cellIndex: number): void => {
    if (cellBuffer === undefined && cellIndex === args.targetCellIndices[cellCount]) {
      cellCount += 1
      return
    }
    let buffer = materializeCellBuffer()
    if (cellCount === buffer.length) {
      const next = new Uint32Array(buffer.length * 2)
      next.set(buffer)
      cellBuffer = next
      buffer = next
    }
    buffer[cellCount] = cellIndex
    cellCount += 1
  }
  const materializeReusableCells = (): Uint8Array => {
    if (reusableCells) {
      return reusableCells
    }
    reusableCells = new Uint8Array(args.maxTargetCellIndex + 1)
    return reusableCells
  }
  const noteCell = (cellIndex: number, runtimeFormula: RuntimeFormula): void => {
    pushCell(cellIndex)
    const reusable = materializeReusableCells()
    let canReuse = true
    for (let index = 0; index < runtimeFormula.dependencyIndices.length; index += 1) {
      const dependencyCellIndex = runtimeFormula.dependencyIndices[index]!
      if (
        ((args.state.workbook.cellStore.flags[dependencyCellIndex] ?? 0) & CellFlags.HasFormula) !== 0 &&
        reusable[dependencyCellIndex] !== 1
      ) {
        canReuse = false
        break
      }
    }
    if (canReuse) {
      reusable[cellIndex] = 1
    } else {
      allCellsReusable = false
    }
  }
  const noteReusableCells = (cellIndices: U32): void => {
    if (cellIndices.length === 0) {
      return
    }
    const reusable = materializeReusableCells()
    for (let index = 0; index < cellIndices.length; index += 1) {
      reusable[cellIndices[index]!] = 1
    }
    const existingPreEvaluated = cellCount > 0 ? (cellBuffer ?? args.targetCellIndices) : undefined
    const preEvaluatedCells = new Uint8Array(args.maxTargetCellIndex + 1)
    for (let index = 0; index < cellCount; index += 1) {
      preEvaluatedCells[existingPreEvaluated![index]!] = 1
    }
    for (let index = 0; index < cellIndices.length; index += 1) {
      preEvaluatedCells[cellIndices[index]!] = 1
    }
    let preEvaluatedCount = 0
    const orderedPreparedCells = args.orderedPreparedCellList()
    const orderedPreparedCellCount = args.orderedPreparedCellCount()
    const orderedPreEvaluatedCells = new Uint32Array(cellCount + cellIndices.length)
    for (let index = 0; index < orderedPreparedCellCount; index += 1) {
      const cellIndex = orderedPreparedCells[index]!
      if (preEvaluatedCells[cellIndex] === 1) {
        orderedPreEvaluatedCells[preEvaluatedCount] = cellIndex
        preEvaluatedCount += 1
      }
    }
    cellBuffer = orderedPreEvaluatedCells
    cellCount = preEvaluatedCount
  }

  return {
    get cellCount() {
      return cellCount
    },
    get cellIndices() {
      return cellCount === 0 ? undefined : (cellBuffer ?? args.targetCellIndices)
    },
    get allCellsReusable() {
      return allCellsReusable
    },
    noteCell,
    noteReusableCells,
  }
}
