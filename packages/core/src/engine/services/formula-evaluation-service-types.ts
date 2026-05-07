import type { Effect } from 'effect'
import type { CellValue } from '@bilig/protocol'
import type { FormulaNode } from '@bilig/formula'
import type { EngineFormulaEvaluationError } from '../errors.js'

export interface EngineFormulaEvaluationService {
  readonly evaluateDirectLookupFormula: (cellIndex: number) => Effect.Effect<number[] | undefined, EngineFormulaEvaluationError>
  readonly evaluateDirectLookupFormulaNow: (cellIndex: number) => number[] | undefined
  readonly evaluateUnsupportedFormula: (cellIndex: number) => Effect.Effect<number[], EngineFormulaEvaluationError>
  readonly resolveStructuredReference: (
    tableName: string,
    columnName: string,
  ) => Effect.Effect<FormulaNode | undefined, EngineFormulaEvaluationError>
  readonly resolveSpillReference: (
    currentSheetName: string,
    sheetName: string | undefined,
    address: string,
  ) => Effect.Effect<FormulaNode | undefined, EngineFormulaEvaluationError>
  readonly resolveMultipleOperations: (request: {
    formulaSheetName: string
    formulaAddress: string
    rowCellSheetName: string
    rowCellAddress: string
    rowReplacementSheetName: string
    rowReplacementAddress: string
    columnCellSheetName?: string
    columnCellAddress?: string
    columnReplacementSheetName?: string
    columnReplacementAddress?: string
  }) => Effect.Effect<CellValue, EngineFormulaEvaluationError>
}
