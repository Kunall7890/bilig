import type { Effect } from 'effect'
import type { CompiledFormula, StructuralAxisTransform } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import type { EngineCounters } from '../../perf/engine-counters.js'
import type { RangeRegistry } from '../../range-registry.js'
import type { FormulaTable } from '../../formula-table.js'
import type { FormulaFamily, FormulaFamilyStructuralSourceTransform } from '../../formula/formula-family-store.js'
import type { WorkbookPivotRecord, WorkbookStore } from '../../workbook-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import type { EngineStructureError } from '../errors.js'

export type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns'
  }
>

export type StructuralAxisMutationSource = 'local' | 'remote' | 'restore' | 'undo' | 'redo'

export interface StructuralAxisApplyOptions {
  readonly source?: StructuralAxisMutationSource
}

export interface EngineStructureState {
  readonly workbook: WorkbookStore
  readonly formulas: FormulaTable<RuntimeFormula>
  readonly ranges: RangeRegistry
  readonly pivotOutputOwners: Map<number, string>
  readonly counters?: EngineCounters
}

export interface StructuralFormulaRebindInput {
  readonly cellIndex: number
  readonly ownerSheetName: string
  readonly ownerRow: number
  readonly ownerCol: number
  readonly source: string
  readonly compiled?: CompiledFormula
  readonly templateId?: number
  readonly preservesBinding?: boolean
  readonly preservesValue?: boolean
}

export interface StructuralAxisOpResult {
  readonly transaction: StructuralTransaction
  readonly changedCellIndices: number[]
  readonly precomputedChangedInputCellIndices: number[]
  readonly formulaCellIndices: number[]
  readonly topologyChanged: boolean
  readonly graphRefreshRequired: boolean
}

export interface EngineStructureService {
  readonly captureSheetCellState: (sheetName: string) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly captureColumnRangeCellState: (sheetName: string, start: number, count: number) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly materializeDeferredStructuralFormulaSourcesNow: () => void
  readonly materializeDeferredStructuralFormulaSources: () => Effect.Effect<void, EngineStructureError>
  readonly applyStructuralAxisOpNow: (op: StructuralAxisOp, options?: StructuralAxisApplyOptions) => StructuralAxisOpResult
  readonly applyStructuralAxisOp: (
    op: StructuralAxisOp,
    options?: StructuralAxisApplyOptions,
  ) => Effect.Effect<StructuralAxisOpResult, EngineStructureError>
}

export interface CreateEngineStructureServiceArgs {
  readonly state: EngineStructureState
  readonly captureStoredCellOps: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
  readonly clearOwnedSpill: (cellIndex: number) => readonly number[]
  readonly writeTableHeaderCell: (sheetName: string, row: number, col: number, value: string) => number | undefined
  readonly removeFormula: (cellIndex: number) => boolean
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[]
  readonly refreshRangeDependencies: (rangeIndices: readonly number[]) => void
  readonly retargetRangeDependencies: (transaction: StructuralTransaction, rangeIndices: readonly number[]) => void
  readonly rebindFormulaCells: (inputs: readonly StructuralFormulaRebindInput[]) => void
  readonly retargetDirectAggregateFormulaForStructuralTransform: (
    input: StructuralFormulaRebindInput,
    targetSheetName: string,
    transform: StructuralAxisTransform,
  ) => boolean
  readonly retargetDirectAggregateFormulasForStructuralTransform: (
    inputs: readonly {
      readonly cellIndex: number
      readonly ownerSheetName: string
      readonly preservesValue?: boolean
    }[],
    targetSheetName: string,
    transform: StructuralAxisTransform,
  ) => readonly number[]
  readonly rewriteFormulaCompiledPreservingBinding: (input: StructuralFormulaRebindInput) => boolean
  readonly collectFormulaCellsOwnedBySheet: (sheetName: string) => readonly number[]
  readonly forEachFormulaCellOwnedBySheet: (sheetName: string, fn: (cellIndex: number) => void) => void
  readonly countFormulaSheetMembers: (sheetId: number) => number
  readonly canUseFormulaFamilyIndex: () => boolean
  readonly tryDeferFormulaFamilyStructuralSourceTransforms: (
    sheetId: number,
    transform: FormulaFamilyStructuralSourceTransform,
    canDeferCellIndex: (cellIndex: number) => boolean,
  ) => number | undefined
  readonly forEachFormulaFamily: (fn: (family: FormulaFamily) => void) => void
  readonly setFormulaFamilyStructuralSourceTransform: (
    familyId: number,
    transform: NonNullable<RuntimeFormula['structuralSourceTransform']>,
  ) => void
  readonly consumeFormulaFamilyStructuralSourceTransforms: () => readonly {
    readonly cellIndices: readonly number[]
    readonly transform: NonNullable<RuntimeFormula['structuralSourceTransform']>
  }[]
  readonly collectFormulaCellsReferencingSheet: (sheetName: string) => readonly number[]
  readonly collectFormulaCellsForDefinedNames: (names: readonly string[]) => readonly number[]
  readonly collectFormulaCellsForTables: (tableNames: readonly string[]) => readonly number[]
}
