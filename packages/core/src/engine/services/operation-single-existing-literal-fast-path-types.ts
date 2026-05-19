import type { CellValue, EngineEvent, LiteralInput } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import type { SheetRecord } from '../../workbook-store.js'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor, U32 } from '../runtime-state.js'
import type { DirectFormulaIndexCollection, DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import type { OperationLookupPlanner } from './operation-lookup-planner.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'

export type MutationSource = 'local' | 'restore' | 'undo' | 'redo'

type SingleExistingLiteralState = Pick<
  EngineRuntimeState,
  'workbook' | 'strings' | 'events' | 'formulas' | 'counters' | 'trackReplicaVersions' | 'getLastMetrics' | 'setLastMetrics'
>

export interface OperationSingleExistingLiteralFastPathArgs {
  readonly state: SingleExistingLiteralState
  readonly hasVolatileFormulas: (() => boolean) | undefined
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly noteAggregateLiteralWrite: (request: {
    readonly sheetName: string
    readonly row: number
    readonly col: number
    readonly oldValue: CellValue
    readonly newValue: CellValue
  }) => void
  readonly evaluateDirectFormula: (cellIndex: number) => readonly number[] | undefined
  readonly invalidateExactLookupColumn: (request: { readonly sheetName: string; readonly col: number }) => void
  readonly invalidateSortedLookupColumn: (request: { readonly sheetName: string; readonly col: number }) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedColumnDependentsAnywhere: () => boolean
  readonly collectRegionFormulaDependentsForCell: (sheetName: string, row: number, col: number) => U32
  readonly collectSingleRegionFormulaDependentForCell: (sheetName: string, row: number, col: number) => number
  readonly collectSingleRegionFormulaDependentForCellAt?: ((sheetId: number, row: number, col: number) => number) | undefined
  readonly canSkipApproximateLookupNewNumericColumnWrite: (sheetId: number, col: number, row: number) => boolean
  readonly writeNumericLiteralToExistingCell: (cellIndex: number, value: number) => void
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => EngineEvent['metrics']
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly directScalarCellNumericValue: (cellIndex: number) => number | undefined
  readonly tryApplySingleDirectAggregateLiteralMutationFastPath: (request: {
    readonly existingIndex: number
    readonly sheetId?: number
    readonly sheetName: string
    readonly row: number
    readonly col: number
    readonly value: LiteralInput
    readonly delta: number
    readonly emitTracked: boolean
    readonly singleRangeEntityDependent?: number
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation: (request: {
    readonly existingIndex: number
    readonly rangeEntityDependent: number
    readonly sheet: SheetRecord
    readonly sheetId: number
    readonly col: number
    readonly value: number
    readonly delta: number
    readonly hasExactLookupDependents: boolean
    readonly hasSortedLookupDependents: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedColumnDirectAggregateExistingNumericMutation: (request: {
    readonly existingIndex: number
    readonly sheet: SheetRecord
    readonly sheetId: number
    readonly sheetName: string
    readonly row: number
    readonly col: number
    readonly value: number
    readonly delta: number
    readonly hasExactLookupDependents: boolean
    readonly hasSortedLookupDependents: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedDirectScalarClosureExistingNumericMutation: (request: {
    readonly existingIndex: number
    readonly sheet: SheetRecord
    readonly sheetId: number
    readonly col: number
    readonly value: number
    readonly oldNumber: number
    readonly hasTrackedEventListeners: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedFormulaLeafExistingNumericMutation: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly sheet: SheetRecord
    readonly col: number
    readonly value: number
    readonly oldNumber: number
    readonly hasTrackedEventListeners: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyFormulaLeafExistingLiteralMutation: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly value: LiteralInput
    readonly hasTrackedEventListeners: boolean
  }) => EngineExistingNumericCellMutationResult | null
  readonly planExactLookupNumericColumnWrite: OperationLookupPlanner['planExactLookupNumericColumnWrite']
  readonly planApproximateLookupNumericColumnWrite: OperationLookupPlanner['planApproximateLookupNumericColumnWrite']
  readonly patchUniformLookupTailWrites: (request: {
    readonly sheetId: number
    readonly col: number
    readonly row: number
    readonly oldNumeric: number
    readonly newNumeric: number
    readonly exact: boolean
    readonly sorted: boolean
  }) => { readonly exact: boolean; readonly sorted: boolean }
  readonly tryApplySingleKernelSyncOnlyLiteralMutationFastPath: (request: {
    readonly existingIndex: number
    readonly value: LiteralInput
    readonly emitTracked: boolean
  }) => boolean
  readonly tryApplySingleDirectFormulaLiteralMutationWithoutEvents: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly value: LiteralInput
    readonly oldNumber: number
    readonly newNumber: number
    readonly exactLookupValue: number | undefined
    readonly approximateLookupValue: number | undefined
  }) => boolean
  readonly tryApplySingleDirectScalarLiteralMutationWithoutEvents: (request: {
    readonly existingIndex: number
    readonly value: LiteralInput
    readonly oldNumber: number
    readonly newNumber: number
  }) => boolean
  readonly tryApplySingleDirectScalarLiteralMutationWithoutEventsAndReturnChanged: (request: {
    readonly existingIndex: number
    readonly value: LiteralInput
    readonly oldNumber: number
    readonly newNumber: number
  }) => U32 | null
  readonly tryApplySingleDirectLookupOperandMutationFastPath: (request: {
    readonly existingIndex: number
    readonly formulaCellIndex: number
    readonly value: LiteralInput
    readonly exactLookupValue: number | undefined
    readonly approximateLookupValue: number | undefined
    readonly emitTracked: boolean
    readonly lookupSheetHint?: SheetRecord | undefined
    readonly trustedInputSheet?: SheetRecord | undefined
    readonly trustedInputCol?: number | undefined
  }) => EngineExistingNumericCellMutationResult | null
  readonly markPostRecalcDirectScalarNumericDependents: (
    cellIndex: number,
    oldNumber: number,
    newNumber: number,
    collection: DirectFormulaIndexCollection,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
  ) => boolean
  readonly tryMarkDirectScalarLinearDeltaClosure: (
    cellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    collection: DirectFormulaIndexCollection,
  ) => boolean
  readonly collectSingleAffectedDirectRangeDependent: (request: {
    readonly sheetName: string
    readonly sheetId?: number
    readonly row: number
    readonly col: number
  }) => number
  readonly collectAffectedDirectRangeDependents: (request: {
    readonly sheetName: string
    readonly row: number
    readonly col: number
  }) => readonly number[]
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, value: DirectScalarCurrentOperand) => boolean
  readonly applyDirectFormulaNumericDelta: (cellIndex: number, delta: number) => boolean
  readonly applyDirectScalarCurrentValue: (cellIndex: number, directScalar: RuntimeDirectScalarDescriptor) => boolean
  readonly tryApplyDirectScalarDeltas: (collection: DirectFormulaIndexCollection, collectChanged?: boolean) => U32 | undefined
  readonly tryApplyDirectFormulaDeltas: (collection: DirectFormulaIndexCollection, collectChanged?: boolean) => U32 | undefined
  readonly countPostRecalcDirectFormulaMetric: (cellIndex: number, counts: DirectFormulaMetricCounts) => void
  readonly hasDynamicFormulaDependents: (cellIndex: number) => boolean
}
