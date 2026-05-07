import { parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import type { WorkbookStore } from '../../workbook-store.js'
import type { CommitOp, EngineRuntimeState, TransactionRecord } from '../runtime-state.js'
import { collectLiveCreatedSheetNames } from './mutation-cell-content-helpers.js'
import { createLazyRenderCommitTransactionRecord, type RenderCommitCellMutation } from './mutation-transaction-records.js'

interface MutationRenderCommitFastPathRuntime {
  readonly state: Pick<EngineRuntimeState, 'undoStack' | 'redoStack' | 'getTransactionReplayDepth' | 'setTransactionReplayDepth'> & {
    readonly workbook: WorkbookStore
  }
  readonly ops: readonly CommitOp[]
  readonly hasExternallyVisibleBatchRequirement: () => boolean
  readonly restoreCellOpFromRef: (ref: EngineCellMutationRef) => EngineOp
  readonly executeLocalNowWithCustomApply: (
    ops: EngineOp[],
    potentialNewCells: number | undefined,
    applyForward: (forward: TransactionRecord) => void,
    options: {
      readonly returnUndoOps?: boolean
      readonly reuseForwardOps?: boolean
    },
  ) => readonly EngineOp[] | null
  readonly executeTransactionNow: (record: TransactionRecord, source: 'local' | 'restore' | 'undo' | 'redo') => void
  readonly applyCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    options: {
      readonly captureUndo?: boolean
      readonly potentialNewCells?: number
      readonly source?: 'local' | 'restore'
      readonly returnUndoOps?: boolean
      readonly reuseRefs?: boolean
    },
  ) => readonly EngineOp[] | null
}

export function tryExecuteMutationRenderCommitFastPath(args: MutationRenderCommitFastPathRuntime): boolean {
  if (args.hasExternallyVisibleBatchRequirement()) {
    return false
  }

  const createdSheetNames = collectLiveCreatedSheetNames(args.state.workbook.sheetsByName.keys(), args.ops)
  const prefixOps: EngineOp[] = []
  const cellMutations: RenderCommitCellMutation[] = []
  let potentialNewCells = 0
  let sawCellMutation = false

  for (let index = 0; index < args.ops.length; index += 1) {
    const op = args.ops[index]
    if (!op) {
      continue
    }
    switch (op.kind) {
      case 'upsertWorkbook':
        if (sawCellMutation) {
          return false
        }
        if (op.name) {
          prefixOps.push({ kind: 'upsertWorkbook', name: op.name })
        }
        break
      case 'upsertSheet':
        if (sawCellMutation) {
          return false
        }
        if (op.name) {
          prefixOps.push({ kind: 'upsertSheet', name: op.name, order: op.order ?? 0 })
        }
        break
      case 'renameSheet':
        if (sawCellMutation) {
          return false
        }
        if (op.oldName && op.newName) {
          prefixOps.push({
            kind: 'renameSheet',
            oldName: op.oldName,
            newName: op.newName,
          })
        }
        break
      case 'deleteSheet':
        if (sawCellMutation) {
          return false
        }
        if (op.name) {
          prefixOps.push({ kind: 'deleteSheet', name: op.name })
        }
        break
      case 'upsertCell': {
        if (!op.sheetName || !op.addr || op.format !== undefined) {
          return false
        }
        const preparedCellAddress = parseCellAddress(op.addr, op.sheetName)
        cellMutations.push({
          sheetName: op.sheetName,
          mutation:
            op.formula !== undefined
              ? { kind: 'setCellFormula', row: preparedCellAddress.row, col: preparedCellAddress.col, formula: op.formula }
              : { kind: 'setCellValue', row: preparedCellAddress.row, col: preparedCellAddress.col, value: op.value ?? null },
        })
        potentialNewCells += 1
        sawCellMutation = true
        break
      }
      case 'deleteCell': {
        if (!op.sheetName || !op.addr) {
          return false
        }
        const preparedCellAddress = parseCellAddress(op.addr, op.sheetName)
        cellMutations.push({
          sheetName: op.sheetName,
          mutation: { kind: 'clearCell', row: preparedCellAddress.row, col: preparedCellAddress.col },
        })
        sawCellMutation = true
        break
      }
      default:
        return false
    }
  }

  if (cellMutations.length === 0) {
    return false
  }

  const priorReplayDepth = args.state.getTransactionReplayDepth()
  let prefixUndoOps: readonly EngineOp[] | null = null
  let cellUndoOps: readonly EngineOp[] | null = null
  args.state.setTransactionReplayDepth(priorReplayDepth + 1)
  try {
    if (prefixOps.length > 0) {
      prefixUndoOps = args.executeLocalNowWithCustomApply(
        prefixOps,
        undefined,
        (forward) => {
          args.executeTransactionNow(forward, 'local')
        },
        { returnUndoOps: true, reuseForwardOps: true },
      )
    }

    const refs: EngineCellMutationRef[] = Array.from({ length: cellMutations.length })
    const sheetIdByName = new Map<string, number>()
    const createdSheetMutationFlags = new Uint8Array(cellMutations.length)
    let sawExistingSheetMutation = false
    for (let index = 0; index < cellMutations.length; index += 1) {
      const mutation = cellMutations[index]!
      let sheetId = sheetIdByName.get(mutation.sheetName)
      if (sheetId === undefined) {
        const sheet = args.state.workbook.getSheet(mutation.sheetName)
        if (!sheet) {
          throw new Error(`Unknown sheet: ${mutation.sheetName}`)
        }
        sheetId = sheet.id
        sheetIdByName.set(mutation.sheetName, sheetId)
      }
      const targetsCreatedSheet = createdSheetNames.has(mutation.sheetName)
      createdSheetMutationFlags[index] = targetsCreatedSheet ? 1 : 0
      sawExistingSheetMutation ||= !targetsCreatedSheet
      refs[index] = {
        sheetId,
        mutation: mutation.mutation,
      }
    }

    const inverseOps: EngineOp[] = []
    if (sawExistingSheetMutation) {
      for (let index = refs.length - 1; index >= 0; index -= 1) {
        if (createdSheetMutationFlags[index] === 1) {
          continue
        }
        inverseOps.push(args.restoreCellOpFromRef(refs[index]!))
      }
    }
    cellUndoOps = inverseOps

    args.applyCellMutationsAtNow(refs, {
      captureUndo: false,
      source: 'local',
      potentialNewCells,
      returnUndoOps: false,
      reuseRefs: true,
    })
  } finally {
    args.state.setTransactionReplayDepth(priorReplayDepth)
  }

  if (priorReplayDepth === 0) {
    const inverseOps = [...(cellUndoOps ?? []), ...(prefixUndoOps ?? [])]
    args.state.undoStack.push({
      forward: createLazyRenderCommitTransactionRecord(prefixOps, cellMutations, potentialNewCells),
      inverse: {
        kind: 'ops',
        ops: inverseOps,
        potentialNewCells: inverseOps.length,
      },
    })
    args.state.redoStack.length = 0
  }

  return true
}
