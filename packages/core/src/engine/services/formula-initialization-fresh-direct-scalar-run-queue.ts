import type { EngineCounters } from '../../perf/engine-counters.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { FreshDirectScalarFormulaBindingInput, FreshDirectScalarFormulaBindingMember } from './formula-binding-service-types.js'
import type { InitialResolvedFormulaEntry } from './formula-initialization-refs.js'
import { createInitialFreshDirectScalarFormulaBindingMember } from './formula-initialization-fresh-direct-scalar-binding.js'

export interface InitialFreshDirectScalarFormulaRunQueue {
  readonly flush: () => void
  readonly queue: (prepared: InitialResolvedFormulaEntry) => void
}

export function createInitialFreshDirectScalarFormulaRunQueue(args: {
  readonly bindFreshDirectScalarFormulaRun: ((run: FreshDirectScalarFormulaBindingInput) => void) | undefined
  readonly clearPendingFormulaCell: (cellIndex: number) => void
  readonly counters: EngineCounters
  readonly noteInitializedFormula: (prepared: InitialResolvedFormulaEntry, runtimeFormula: RuntimeFormula | undefined) => void
  readonly readRuntimeFormula: (cellIndex: number) => RuntimeFormula | undefined
}): InitialFreshDirectScalarFormulaRunQueue {
  let pending:
    | {
        readonly sheetId: number
        readonly ownerSheetName: string
        readonly cellIndices: number[]
        readonly members: FreshDirectScalarFormulaBindingMember[]
        readonly prepared: InitialResolvedFormulaEntry[]
      }
    | undefined

  const flush = (): void => {
    const run = pending
    if (run === undefined) {
      return
    }
    pending = undefined
    const bindFreshDirectScalarFormulaRun = args.bindFreshDirectScalarFormulaRun
    if (bindFreshDirectScalarFormulaRun === undefined) {
      return
    }
    if (run.prepared.length === 1) {
      const prepared = run.prepared[0]!
      bindFreshDirectScalarFormulaRun({
        sheetId: prepared.sheetId,
        ownerSheetName: prepared.ownerSheetName,
        cellIndex: prepared.cellIndex,
        member: run.members[0]!,
      })
    } else {
      bindFreshDirectScalarFormulaRun({
        sheetId: run.sheetId,
        ownerSheetName: run.ownerSheetName,
        cellIndices: run.cellIndices,
        members: run.members,
      })
    }
    addEngineCounter(args.counters, 'initialFreshDirectScalarFastBindings', run.prepared.length)
    for (const prepared of run.prepared) {
      args.noteInitializedFormula(prepared, args.readRuntimeFormula(prepared.cellIndex))
      args.clearPendingFormulaCell(prepared.cellIndex)
    }
  }

  return {
    flush,
    queue(prepared) {
      if (pending !== undefined && (pending.sheetId !== prepared.sheetId || pending.ownerSheetName !== prepared.ownerSheetName)) {
        flush()
      }
      pending ??= {
        sheetId: prepared.sheetId,
        ownerSheetName: prepared.ownerSheetName,
        cellIndices: [],
        members: [],
        prepared: [],
      }
      pending.cellIndices.push(prepared.cellIndex)
      pending.members.push(createInitialFreshDirectScalarFormulaBindingMember(prepared))
      pending.prepared.push(prepared)
    },
  }
}
