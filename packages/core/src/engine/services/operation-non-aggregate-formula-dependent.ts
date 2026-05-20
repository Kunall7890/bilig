import type { FormulaTable } from '../../formula-table.js'
import type {
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectScalarDescriptor,
  RuntimeDirectScalarOperand,
  RuntimeFormula,
} from '../runtime-state.js'

function directScalarOperandTouchesCell(operand: RuntimeDirectScalarOperand, cellIndex: number): boolean {
  return operand.kind === 'cell' && operand.cellIndex === cellIndex
}

function directScalarTouchesCell(directScalar: RuntimeDirectScalarDescriptor, cellIndex: number): boolean {
  return directScalar.kind === 'abs'
    ? directScalarOperandTouchesCell(directScalar.operand, cellIndex)
    : directScalarOperandTouchesCell(directScalar.left, cellIndex) || directScalarOperandTouchesCell(directScalar.right, cellIndex)
}

function directCriteriaTouchesCell(directCriteria: RuntimeDirectCriteriaDescriptor, cellIndex: number): boolean {
  if (directCriteria.offsetOperand !== undefined && directScalarOperandTouchesCell(directCriteria.offsetOperand, cellIndex)) {
    return true
  }
  for (const pair of directCriteria.criteriaPairs) {
    const criterion = pair.criterion
    if (
      (criterion.kind === 'cell' || criterion.kind === 'cell-string-concat' || criterion.kind === 'cell-month-boundary-string-concat') &&
      criterion.cellIndex === cellIndex
    ) {
      return true
    }
  }
  return (
    directCriteria.resultTransforms?.some((transform) => transform.kind === 'if-empty-cell' && transform.cellIndex === cellIndex) ?? false
  )
}

export function hasNonAggregateFormulaDependentForCell(
  formulas: FormulaTable<RuntimeFormula>,
  existingIndex: number,
  scanLimit: number,
): boolean {
  if (formulas.size > scanLimit) {
    return true
  }
  for (const formula of formulas.values()) {
    if (formula.directAggregate !== undefined) {
      continue
    }
    const directScalar = formula.directScalar
    if (directScalar !== undefined && directScalarTouchesCell(directScalar, existingIndex)) {
      return true
    }
    const directLookup = formula.directLookup
    if (directLookup !== undefined && directLookup.operandCellIndex === existingIndex) {
      return true
    }
    const directCriteria = formula.directCriteria
    if (directCriteria !== undefined && directCriteriaTouchesCell(directCriteria, existingIndex)) {
      return true
    }
    if (formula.inlineScalarPlanCellIndices?.includes(existingIndex)) {
      return true
    }
    for (let index = 0; index < formula.dependencyIndices.length; index += 1) {
      if (formula.dependencyIndices[index] === existingIndex) {
        return true
      }
    }
  }
  return false
}
