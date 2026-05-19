import type { StructuralAxisTransform } from '@bilig/formula'
import type { RuntimeDirectAggregateDescriptor } from '../runtime-state.js'
import { rewriteDirectAggregateDescriptorForStructuralTransform } from './formula-binding-direct-descriptors.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export interface FormulaBindingDirectAggregateRetargeter {
  readonly retargetDirectAggregateFormulaForStructuralTransformNow: (
    cellIndex: number,
    ownerSheetName: string,
    targetSheetName: string,
    transform: StructuralAxisTransform,
    preservesValue: boolean,
  ) => boolean
  readonly retargetDirectAggregateFormulasForStructuralTransformNow: (
    inputs: readonly {
      readonly cellIndex: number
      readonly ownerSheetName: string
      readonly preservesValue: boolean
    }[],
    targetSheetName: string,
    transform: StructuralAxisTransform,
  ) => readonly number[]
}

export function createFormulaBindingDirectAggregateRetargeter(
  args: CreateEngineFormulaBindingServiceArgs,
): FormulaBindingDirectAggregateRetargeter {
  const directAggregateContainsFormulaOwner = (
    cellIndex: number,
    ownerSheetName: string,
    aggregate: RuntimeDirectAggregateDescriptor,
  ): boolean => {
    if (aggregate.sheetName !== ownerSheetName) {
      return false
    }
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    if (sheetId === undefined) {
      return false
    }
    const ownerSheet = args.state.workbook.getSheetById(sheetId)
    const ownerPosition =
      ownerSheet?.structureVersion === 1
        ? {
            row: args.state.workbook.cellStore.rows[cellIndex],
            col: args.state.workbook.cellStore.cols[cellIndex],
          }
        : args.state.workbook.getCellPosition(cellIndex)
    if (ownerPosition?.row === undefined || ownerPosition.col === undefined) {
      return false
    }
    const aggregateColEnd = aggregate.colEnd ?? aggregate.col
    return (
      ownerPosition.row >= aggregate.rowStart &&
      ownerPosition.row <= aggregate.rowEnd &&
      ownerPosition.col >= aggregate.col &&
      ownerPosition.col <= aggregateColEnd
    )
  }

  const retargetDirectAggregateFormulaForStructuralTransformNow = (
    cellIndex: number,
    ownerSheetName: string,
    targetSheetName: string,
    transform: StructuralAxisTransform,
    preservesValue: boolean,
  ): boolean => {
    const existing = args.state.formulas.get(cellIndex)
    if (!existing?.directAggregate) {
      return false
    }
    const previousDirectAggregate = existing.directAggregate
    const nextDirectAggregate = rewriteDirectAggregateDescriptorForStructuralTransform({
      descriptor: previousDirectAggregate,
      targetSheetName,
      transform,
      regionGraph: args.regionGraph,
    })
    if (!nextDirectAggregate) {
      return false
    }
    if (directAggregateContainsFormulaOwner(cellIndex, ownerSheetName, nextDirectAggregate)) {
      return false
    }
    existing.directAggregate = nextDirectAggregate
    existing.structuralSourceTransform = {
      ownerSheetName,
      targetSheetName,
      transform,
      preservesValue,
    }
    args.regionGraph.replaceSingleFormulaSubscription(cellIndex, previousDirectAggregate.regionId, nextDirectAggregate.regionId)
    return true
  }

  const retargetDirectAggregateFormulasForStructuralTransformNow = (
    inputs: readonly {
      readonly cellIndex: number
      readonly ownerSheetName: string
      readonly preservesValue: boolean
    }[],
    targetSheetName: string,
    transform: StructuralAxisTransform,
  ): readonly number[] => {
    if (inputs.length === 0) {
      return []
    }
    const retargetedCellIndices: number[] = []
    const replacements: Array<{ formulaCellIndex: number; previousRegionId: number; nextRegionId: number }> = []
    for (let index = 0; index < inputs.length; index += 1) {
      const { cellIndex, ownerSheetName, preservesValue } = inputs[index]!
      const existing = args.state.formulas.get(cellIndex)
      if (!existing?.directAggregate) {
        continue
      }
      const previousDirectAggregate = existing.directAggregate
      const nextDirectAggregate = rewriteDirectAggregateDescriptorForStructuralTransform({
        descriptor: previousDirectAggregate,
        targetSheetName,
        transform,
        regionGraph: args.regionGraph,
      })
      if (!nextDirectAggregate) {
        continue
      }
      if (directAggregateContainsFormulaOwner(cellIndex, ownerSheetName, nextDirectAggregate)) {
        continue
      }
      existing.directAggregate = nextDirectAggregate
      existing.structuralSourceTransform = {
        ownerSheetName,
        targetSheetName,
        transform,
        preservesValue,
      }
      retargetedCellIndices.push(cellIndex)
      replacements.push({
        formulaCellIndex: cellIndex,
        previousRegionId: previousDirectAggregate.regionId,
        nextRegionId: nextDirectAggregate.regionId,
      })
    }
    args.regionGraph.replaceSingleFormulaSubscriptions(replacements)
    return retargetedCellIndices
  }

  return {
    retargetDirectAggregateFormulaForStructuralTransformNow,
    retargetDirectAggregateFormulasForStructuralTransformNow,
  }
}
