import type { StructuralAxisTransform } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import { emptyValue } from '../../engine-value-utils.js'
import { mapStructuralAxisIndex } from '../../engine-structural-utils.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'

export function collectStructuralRangeDependencies(
  args: CreateEngineStructureServiceArgs,
  formulaCellIndices: readonly number[],
): number[] {
  const rangeIndices = new Set<number>()
  formulaCellIndices.forEach((cellIndex) => {
    const formula = args.state.formulas.get(cellIndex)
    if (!formula) {
      return
    }
    formula.rangeDependencies.forEach((rangeIndex) => {
      rangeIndices.add(rangeIndex)
    })
  })
  return [...rangeIndices]
}

export function clearSpillArtifactsForSheet(args: CreateEngineStructureServiceArgs, sheetName: string): number[] {
  const changedCellIndices: number[] = []
  args.state.workbook.listSpills().forEach((spill) => {
    if (spill.sheetName !== sheetName) {
      return
    }
    const ownerCellIndex = args.state.workbook.getCellIndex(spill.sheetName, spill.address)
    if (ownerCellIndex === undefined) {
      args.state.workbook.deleteSpill(spill.sheetName, spill.address)
      return
    }
    changedCellIndices.push(...args.clearOwnedSpill(ownerCellIndex))
  })
  return changedCellIndices
}

export function clearPivotOutputsForSheet(args: CreateEngineStructureServiceArgs, sheetName: string): void {
  args.state.workbook
    .listPivots()
    .filter((pivot) => pivot.sheetName === sheetName)
    .forEach((pivot) => {
      args.clearOwnedPivot(pivot)
    })
}

function clearDerivedCellArtifacts(args: CreateEngineStructureServiceArgs, cellIndex: number): void {
  args.state.pivotOutputOwners.delete(cellIndex)
}

export function clearRemovedCellRuntimeState(args: CreateEngineStructureServiceArgs, cellIndex: number): void {
  const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
  const hasDerivedState =
    (flags & (CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0 ||
    args.state.workbook.getCellFormat(cellIndex) !== undefined ||
    args.state.pivotOutputOwners.has(cellIndex)
  if (!hasDerivedState) {
    args.state.workbook.cellStore.flags[cellIndex] = flags & ~CellFlags.Materialized
    return
  }
  clearDerivedCellArtifacts(args, cellIndex)
  args.removeFormula(cellIndex)
  args.state.workbook.setCellFormat(cellIndex, null)
  args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
  args.state.workbook.cellStore.flags[cellIndex] =
    flags &
    ~(CellFlags.Materialized | CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
}

export function isCellIndexMapped(args: CreateEngineStructureServiceArgs, cellIndex: number): boolean {
  const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
  const position = args.state.workbook.getCellPosition(cellIndex)
  if (sheetId === undefined || !position || !Number.isFinite(position.row) || !Number.isFinite(position.col)) {
    return false
  }
  const sheet = args.state.workbook.getSheetById(sheetId)
  return (
    (sheet?.structureVersion === 1 ? sheet.grid.getPhysical(position.row, position.col) : sheet?.grid.get(position.row, position.col)) ===
    cellIndex
  )
}

export function structuralAxisIndexAffected(axisIndex: number, transform: StructuralAxisTransform): boolean {
  const nextIndex = mapStructuralAxisIndex(axisIndex, transform)
  return nextIndex === undefined || nextIndex !== axisIndex
}
