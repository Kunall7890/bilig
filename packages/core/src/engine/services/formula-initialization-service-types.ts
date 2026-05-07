import type { Effect } from 'effect'
import type { CompiledFormula } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import type { EngineCellMutationRef, EngineFormulaSourceRefs } from '../../cell-mutations-at.js'
import type { EngineMutationError } from '../errors.js'

export interface EngineFormulaInitializationService {
  readonly initializeCellFormulasAt: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeCellFormulasAtNow: (refs: readonly EngineCellMutationRef[], potentialNewCells?: number) => void
  readonly initializeFormulaSourcesAtNow: (refs: EngineFormulaSourceRefs, potentialNewCells?: number) => void
  readonly initializePreparedCellFormulasAt: (
    refs: readonly PreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializePreparedCellFormulasAtNow: (refs: readonly PreparedFormulaInitializationRef[], potentialNewCells?: number) => void
  readonly initializeHydratedPreparedCellFormulasAt: (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => Effect.Effect<void, EngineMutationError>
  readonly initializeHydratedPreparedCellFormulasAtNow: (
    refs: readonly HydratedPreparedFormulaInitializationRef[],
    potentialNewCells?: number,
  ) => void
}

export interface PreparedFormulaInitializationRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId?: number
  readonly cellIndex?: number
}

export interface HydratedPreparedFormulaInitializationRef extends PreparedFormulaInitializationRef {
  readonly value: CellValue
}
