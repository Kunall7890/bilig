import type { EngineCounters } from '../../perf/engine-counters.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { FreshDirectScalarFormulaBindingInput } from './formula-binding-service-types.js'
import type { InitialResolvedFormulaEntry } from './formula-initialization-refs.js'
import {
  assertInitialFreshDirectScalarFormulaBindingMember,
  type InitialFreshDirectScalarFormulaBindingMember,
} from './formula-initialization-fresh-direct-scalar-binding.js'

export interface InitialFreshDirectScalarFormulaRunQueue {
  readonly flush: () => void
  readonly queue: (prepared: InitialResolvedFormulaEntry) => void
}

export function createInitialFreshDirectScalarFormulaRunQueue(args: {
  readonly bindFreshDirectScalarFormulaRun: ((run: FreshDirectScalarFormulaBindingInput) => void) | undefined
  readonly clearPendingFormulaCell: (cellIndex: number) => void
  readonly capacity?: number
  readonly counters: EngineCounters
  readonly noteInitializedFormula: (prepared: InitialResolvedFormulaEntry, runtimeFormula: RuntimeFormula | undefined) => void
  readonly readRuntimeFormula: (cellIndex: number) => RuntimeFormula | undefined
}): InitialFreshDirectScalarFormulaRunQueue {
  let pending:
    | {
        readonly sheetId: number
        readonly ownerSheetName: string
        cellIndices: Uint32Array
        cellIndexCount: number
        readonly members: InitialFreshDirectScalarFormulaBindingMember[]
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
    if (run.members.length === 1) {
      const prepared = run.members[0]!
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
        cellIndices: run.cellIndexCount === run.cellIndices.length ? run.cellIndices : run.cellIndices.subarray(0, run.cellIndexCount),
        members: run.members,
      })
    }
    addEngineCounter(args.counters, 'initialFreshDirectScalarFastBindings', run.members.length)
    for (const prepared of run.members) {
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
        cellIndices: new Uint32Array(Math.max(1, args.capacity ?? 16)),
        cellIndexCount: 0,
        members: preallocatedRunArray<InitialFreshDirectScalarFormulaBindingMember>(args.capacity),
      }
      assertInitialFreshDirectScalarFormulaBindingMember(prepared)
      const member = prepared
      if (pending.cellIndexCount === pending.cellIndices.length) {
        const next = new Uint32Array(pending.cellIndices.length * 2)
        next.set(pending.cellIndices)
        pending.cellIndices = next
      }
      pending.cellIndices[pending.cellIndexCount] = member.cellIndex
      pending.cellIndexCount += 1
      pending.members.push(member)
    },
  }
}

function preallocatedRunArray<T>(capacity: number | undefined): T[] {
  const values: T[] = []
  if (capacity !== undefined && capacity > 0) {
    values.length = capacity
    values.length = 0
  }
  return values
}
