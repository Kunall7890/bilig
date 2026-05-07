import type { RuntimeDirectAggregateDescriptor } from '../runtime-state.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export function directAggregateContainsFormulaOwnerCell(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  directAggregate: RuntimeDirectAggregateDescriptor | undefined,
  cellIndex: number,
): boolean {
  if (!directAggregate) {
    return false
  }
  const ownerSheetId = serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]
  const aggregateSheet = serviceArgs.state.workbook.getSheet(directAggregate.sheetName)
  if (ownerSheetId === undefined || !aggregateSheet || aggregateSheet.id !== ownerSheetId) {
    return false
  }
  const ownerPosition = serviceArgs.state.workbook.getCellPosition(cellIndex)
  return (
    ownerPosition !== undefined &&
    ownerPosition.row >= directAggregate.rowStart &&
    ownerPosition.row <= directAggregate.rowEnd &&
    ownerPosition.col >= directAggregate.col &&
    ownerPosition.col <= directAggregate.colEnd
  )
}
