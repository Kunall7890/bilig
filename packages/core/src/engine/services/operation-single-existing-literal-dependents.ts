import { isRangeEntity } from '../../entity-ids.js'
import { hasOperationCompactedRangeDependencies } from './operation-cell-lifecycle-helpers.js'

interface OperationSingleExistingLiteralFormulaShape {
  readonly dependencyIndices: Uint32Array
  readonly rangeDependencies: Uint32Array
  readonly graphRangeDependencies: Uint32Array
}

interface OperationSingleExistingLiteralFormulaTable {
  readonly size: number
  get(cellIndex: number): OperationSingleExistingLiteralFormulaShape | undefined
  entries(): IterableIterator<[number, OperationSingleExistingLiteralFormulaShape]>
}

export interface DynamicFormulaDependentArgs {
  readonly state: { readonly formulas: OperationSingleExistingLiteralFormulaTable }
  readonly collectSingleRegionFormulaDependentForCellAt?: ((sheetId: number, row: number, col: number) => number | undefined) | undefined
  readonly hasDynamicFormulaDependents: (cellIndex: number) => boolean
}

export interface FormulaLeafRangeDependentArgs {
  readonly state: { readonly formulas: OperationSingleExistingLiteralFormulaTable }
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly collectSingleRegionFormulaDependentForCellAt?: ((sheetId: number, row: number, col: number) => number | undefined) | undefined
  readonly collectSingleRegionFormulaDependentForCell: (sheetName: string, row: number, col: number) => number
  readonly collectRegionFormulaDependentsForCell: (sheetName: string, row: number, col: number) => ArrayLike<number>
}

export function hasKnownDynamicFormulaDependentForSingleExistingLiteral(
  args: DynamicFormulaDependentArgs,
  request: {
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly existingIndex: number
    readonly singleExistingCellDependent: number
  },
): boolean {
  if (request.singleExistingCellDependent === -2) {
    return args.hasDynamicFormulaDependents(request.existingIndex)
  }
  if (request.singleExistingCellDependent >= 0 && !isRangeEntity(request.singleExistingCellDependent) && args.state.formulas.size === 1) {
    const formula = args.state.formulas.get(request.singleExistingCellDependent)
    return formula !== undefined && hasOperationCompactedRangeDependencies(formula)
  }
  const singleRegionFormulaDependent = args.collectSingleRegionFormulaDependentForCellAt?.(request.sheetId, request.row, request.col)
  if (singleRegionFormulaDependent === undefined || singleRegionFormulaDependent === -2) {
    return args.hasDynamicFormulaDependents(request.existingIndex)
  }

  let firstFormulaCellIndex = -1
  let secondFormulaCellIndex = -1
  const pushFormulaCellIndex = (candidate: number): void => {
    if (candidate < 0 || isRangeEntity(candidate)) {
      return
    }
    if (firstFormulaCellIndex === -1) {
      firstFormulaCellIndex = candidate
      return
    }
    if (candidate !== firstFormulaCellIndex) {
      secondFormulaCellIndex = candidate
    }
  }

  pushFormulaCellIndex(request.singleExistingCellDependent)
  pushFormulaCellIndex(singleRegionFormulaDependent)
  if (firstFormulaCellIndex === -1) {
    return false
  }
  const firstFormula = args.state.formulas.get(firstFormulaCellIndex)
  if (firstFormula !== undefined && hasOperationCompactedRangeDependencies(firstFormula)) {
    return true
  }
  const secondFormula = secondFormulaCellIndex === -1 ? undefined : args.state.formulas.get(secondFormulaCellIndex)
  return secondFormula !== undefined && hasOperationCompactedRangeDependencies(secondFormula)
}

export function collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(
  args: FormulaLeafRangeDependentArgs,
  request: {
    readonly existingIndex: number
    readonly singleExistingCellDependent: number
    readonly sheetId: number
    readonly sheetName: string
    readonly row: number
    readonly col: number
    readonly formulaScanLimit: number
  },
): number {
  if (isRangeEntity(request.singleExistingCellDependent)) {
    const rangeFormulaDependent = args.getSingleEntityDependent(request.singleExistingCellDependent)
    if (rangeFormulaDependent !== -1) {
      return rangeFormulaDependent
    }
  }
  const indexedSingle =
    args.collectSingleRegionFormulaDependentForCellAt?.(request.sheetId, request.row, request.col) ??
    args.collectSingleRegionFormulaDependentForCell(request.sheetName, request.row, request.col)
  if (indexedSingle >= 0 && !isRangeEntity(indexedSingle)) {
    return indexedSingle
  }
  const dependents = args.collectRegionFormulaDependentsForCell(request.sheetName, request.row, request.col)
  let singleFormulaCellIndex = -1
  for (let index = 0; index < dependents.length; index += 1) {
    const candidate = dependents[index]!
    if (candidate < 0 || isRangeEntity(candidate)) {
      continue
    }
    if (singleFormulaCellIndex !== -1 && singleFormulaCellIndex !== candidate) {
      return -2
    }
    singleFormulaCellIndex = candidate
  }
  if (singleFormulaCellIndex !== -1 || args.state.formulas.size > request.formulaScanLimit) {
    return singleFormulaCellIndex
  }
  for (const [formulaCellIndex, formula] of args.state.formulas.entries()) {
    const dependencyIndices = formula.dependencyIndices
    for (let index = 0; index < dependencyIndices.length; index += 1) {
      if (dependencyIndices[index] !== request.existingIndex) {
        continue
      }
      if (singleFormulaCellIndex !== -1 && singleFormulaCellIndex !== formulaCellIndex) {
        return -2
      }
      singleFormulaCellIndex = formulaCellIndex
      break
    }
  }
  return singleFormulaCellIndex !== -1 ? singleFormulaCellIndex : indexedSingle === -2 ? -2 : -1
}
