import { ValueTag, type CellValue, type RecalcMetrics } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import type { U32 } from '../runtime-state.js'
import type { SheetRecord } from '../../workbook-store.js'
import { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import {
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  hasCompleteDirectFormulaDeltas,
} from './direct-formula-recalc-helpers.js'
import {
  canTrustPhysicalTrackedChangeSplit,
  makeExistingNumericMutationResult,
  tagTrustedPhysicalTrackedChanges,
} from './operation-change-helpers.js'
import { emitOperationTrackedCellsBatch } from './operation-tracked-event-helpers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

export interface OperationTrustedDirectScalarClosureExistingNumericMutationRequest {
  readonly existingIndex: number
  readonly sheet: SheetRecord
  readonly sheetId: number
  readonly col: number
  readonly value: number
  readonly oldNumber: number
  readonly hasTrackedEventListeners: boolean
}

export interface OperationDirectScalarClosureFastPathArgs {
  readonly state: Pick<CreateEngineOperationServiceArgs['state'], 'workbook' | 'formulas' | 'counters' | 'events' | 'setLastMetrics'>
  readonly tryMarkDirectScalarLinearDeltaClosure: (
    rootCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ) => boolean
  readonly writeTrustedExistingNumericLiteralToCell: (existingIndex: number, sheet: SheetRecord, col: number, value: number) => void
  readonly tryApplyDirectScalarDeltas: (collection: DirectFormulaIndexCollection, captureChanged?: boolean) => U32 | undefined
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export function tryApplyTrustedDirectScalarClosureExistingNumericMutation(
  args: OperationDirectScalarClosureFastPathArgs,
  request: OperationTrustedDirectScalarClosureExistingNumericMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const postRecalcDirectFormulaIndices = new DirectFormulaIndexCollection()
  const oldValue: CellValue = { tag: ValueTag.Number, value: request.oldNumber }
  const newValue: CellValue = { tag: ValueTag.Number, value: request.value }
  if (!args.tryMarkDirectScalarLinearDeltaClosure(request.existingIndex, oldValue, newValue, postRecalcDirectFormulaIndices)) {
    return null
  }
  if (!hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices)) {
    return null
  }
  countDirectFormulaDeltaSkip(args.state.formulas, postRecalcDirectFormulaIndices, args.state.counters)
  args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
  const directChanged = args.tryApplyDirectScalarDeltas(postRecalcDirectFormulaIndices, true)
  if (directChanged === undefined) {
    throw new Error('Failed to apply direct scalar closure delta')
  }
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = args.makeSingleLiteralSkipMetrics()
  args.state.setLastMetrics(lastMetrics)
  const changed = composeSingleDisjointExplicitEventChanges(request.existingIndex, directChanged)
  if (changed.length > 4 && canTrustPhysicalTrackedChangeSplit(changed, request.sheetId, 1, args.state.workbook)) {
    tagTrustedPhysicalTrackedChanges(changed, request.sheetId, 1)
  }
  if (request.hasTrackedEventListeners) {
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices: changed,
      metrics: lastMetrics,
    })
  }
  return makeExistingNumericMutationResult(changed, 1)
}
