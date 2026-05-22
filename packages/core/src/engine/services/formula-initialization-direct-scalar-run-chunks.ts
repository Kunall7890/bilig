import { CellFlags } from '../../cell-store.js'
import type {
  EngineRuntimeState,
  RuntimeDirectScalarDescriptor,
  RuntimeDirectScalarOperand,
  RuntimeFormula,
  U32,
} from '../runtime-state.js'
import {
  createInitialNativeDirectScalarBatch,
  MAX_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE,
  MIN_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE,
} from './formula-initialization-native-direct-scalar.js'

interface InitialDirectScalarRunChunkState {
  readonly workbook: EngineRuntimeState['workbook']
  readonly wasm: EngineRuntimeState['wasm']
  readonly counters: EngineRuntimeState['counters']
}

export interface InitialDirectScalarRunChunkCandidate {
  readonly cellIndex: number
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly runtimeFormula: RuntimeFormula
}

interface InitialDirectScalarRunChunk {
  readonly candidates: InitialDirectScalarRunChunkCandidate[]
}

export interface InitialDirectScalarRunChunkCollector {
  readonly candidateCount: number
  readonly hasNativeChunks: () => boolean
  readonly add: (candidate: InitialDirectScalarRunChunkCandidate) => boolean
  readonly evaluate: () => U32 | undefined
}

export interface InitialDirectScalarRunChunkCollectorOptions {
  readonly minChunkSize?: number
  readonly maxChunkSize?: number
}

function operandCanUseChunk(
  operand: RuntimeDirectScalarOperand,
  selectedCells: ReadonlySet<number>,
  cellFlags: ArrayLike<number | undefined>,
): boolean {
  switch (operand.kind) {
    case 'literal-number':
    case 'error':
      return true
    case 'cell':
      return selectedCells.has(operand.cellIndex) || ((cellFlags[operand.cellIndex] ?? 0) & CellFlags.HasFormula) === 0
  }
}

function formulaCanUseChunk(
  directScalar: RuntimeDirectScalarDescriptor,
  selectedCells: ReadonlySet<number>,
  cellFlags: ArrayLike<number | undefined>,
): boolean {
  if (directScalar.kind === 'abs') {
    return operandCanUseChunk(directScalar.operand, selectedCells, cellFlags)
  }
  return operandCanUseChunk(directScalar.left, selectedCells, cellFlags) && operandCanUseChunk(directScalar.right, selectedCells, cellFlags)
}

function materializeChunkPlans(args: {
  readonly candidates: readonly InitialDirectScalarRunChunkCandidate[]
  readonly cellFlags: ArrayLike<number | undefined>
  readonly minChunkSize: number
  readonly maxChunkSize: number
}): InitialDirectScalarRunChunk[] {
  const chunks: InitialDirectScalarRunChunk[] = []
  let current: InitialDirectScalarRunChunkCandidate[] = []
  let selectedCells = new Set<number>()

  const flushCurrent = (): void => {
    if (current.length >= args.minChunkSize) {
      chunks.push({ candidates: current })
    }
    current = []
    selectedCells = new Set()
  }

  for (const candidate of args.candidates) {
    if (current.length === args.maxChunkSize) {
      flushCurrent()
    }
    const directScalar = candidate.runtimeFormula.directScalar
    if (
      !directScalar ||
      candidate.runtimeFormula.compiled.volatile ||
      candidate.runtimeFormula.compiled.producesSpill ||
      !formulaCanUseChunk(directScalar, selectedCells, args.cellFlags)
    ) {
      continue
    }
    current.push(candidate)
    selectedCells.add(candidate.cellIndex)
  }
  flushCurrent()
  return chunks
}

export function createInitialDirectScalarRunChunkCollector(args: {
  readonly state: InitialDirectScalarRunChunkState
  readonly options?: InitialDirectScalarRunChunkCollectorOptions
}): InitialDirectScalarRunChunkCollector {
  const minChunkSize = args.options?.minChunkSize ?? MIN_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE
  const maxChunkSize = args.options?.maxChunkSize ?? MAX_INITIAL_NATIVE_DIRECT_SCALAR_BATCH_SIZE
  const candidates: InitialDirectScalarRunChunkCandidate[] = []
  let chunks: InitialDirectScalarRunChunk[] | undefined

  const chunkPlans = (): InitialDirectScalarRunChunk[] => {
    if (chunks) {
      return chunks
    }
    chunks = materializeChunkPlans({
      candidates,
      cellFlags: args.state.workbook.cellStore.flags,
      minChunkSize,
      maxChunkSize,
    })
    return chunks
  }

  return {
    get candidateCount() {
      return candidates.length
    },
    add(candidate) {
      if (!candidate.runtimeFormula.directScalar) {
        return false
      }
      candidates.push(candidate)
      chunks = undefined
      return true
    },
    hasNativeChunks() {
      return chunkPlans().length > 0
    },
    evaluate() {
      const plans = chunkPlans()
      if (plans.length === 0) {
        return undefined
      }
      const changedChunks: U32[] = []
      let changedCellCount = 0
      for (const plan of plans) {
        const batch = createInitialNativeDirectScalarBatch({
          state: args.state,
          capacity: plan.candidates.length,
        })
        let addedAll = true
        for (const candidate of plan.candidates) {
          const directScalar = candidate.runtimeFormula.directScalar
          if (
            !directScalar ||
            !batch.add(
              {
                cellIndex: candidate.cellIndex,
                sheetId: candidate.sheetId,
                col: candidate.col,
              },
              directScalar,
            )
          ) {
            addedAll = false
            break
          }
        }
        if (!addedAll) {
          continue
        }
        const changedCells = batch.evaluate()
        if (!changedCells || changedCells.length === 0) {
          continue
        }
        changedChunks.push(changedCells)
        changedCellCount += changedCells.length
      }
      if (changedCellCount === 0) {
        return undefined
      }
      if (changedChunks.length === 1) {
        return changedChunks[0]!
      }
      const changedCells = new Uint32Array(changedCellCount)
      let offset = 0
      for (const chunk of changedChunks) {
        changedCells.set(chunk, offset)
        offset += chunk.length
      }
      return changedCells
    },
  }
}
