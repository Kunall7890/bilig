import type { RuntimeDirectAggregateDescriptor } from '../runtime-state.js'
import type { CreateEngineFormulaBindingServiceArgs, FormulaOwnerPosition } from './formula-binding-service-types.js'

export function directAggregateContainsFormulaOwnerCell(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  directAggregate: RuntimeDirectAggregateDescriptor | undefined,
  cellIndex: number,
  ownerPosition?: FormulaOwnerPosition,
): boolean {
  if (!directAggregate) {
    return false
  }
  const ownerSheetId = serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]
  const aggregateSheet = serviceArgs.state.workbook.getSheet(directAggregate.sheetName)
  if (ownerSheetId === undefined || !aggregateSheet || aggregateSheet.id !== ownerSheetId) {
    return false
  }
  const resolvedOwnerPosition = ownerPosition ?? serviceArgs.state.workbook.getCellPosition(cellIndex)
  return (
    resolvedOwnerPosition !== undefined &&
    resolvedOwnerPosition.row >= directAggregate.rowStart &&
    resolvedOwnerPosition.row <= directAggregate.rowEnd &&
    resolvedOwnerPosition.col >= directAggregate.col &&
    resolvedOwnerPosition.col <= directAggregate.colEnd
  )
}
