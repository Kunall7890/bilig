import { getRuntimeFormulaSource } from '../runtime-formula-source.js'
import type { U32 } from '../runtime-state.js'
import { collectTrackedDependents } from './formula-binding-dependency-helpers.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export function createFormulaBindingRebinds(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly bindFormulaNow: (cellIndex: number, ownerSheetName: string, source: string) => boolean
}): {
  readonly rebindFormulaCellsNow: (candidates: readonly number[], formulaChangedCount: number) => number
  readonly rebindTrackedDependentsNow: (registry: Map<string, Set<number>>, keys: readonly string[], formulaChangedCount: number) => number
  readonly rebindFormulasForSheetNow: (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32) => number
} {
  const rebindFormulaCellsNow = (candidates: readonly number[], formulaChangedCount: number): number => {
    candidates.forEach((cellIndex) => {
      const formula = args.serviceArgs.state.formulas.get(cellIndex)
      const ownerSheetName = args.serviceArgs.state.workbook.getSheetNameById(
        args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]!,
      )
      if (formula && ownerSheetName) {
        args.bindFormulaNow(cellIndex, ownerSheetName, getRuntimeFormulaSource(formula))
      }
      formulaChangedCount = args.serviceArgs.markFormulaChanged(cellIndex, formulaChangedCount)
    })
    return formulaChangedCount
  }

  const rebindTrackedDependentsNow = (registry: Map<string, Set<number>>, keys: readonly string[], formulaChangedCount: number): number =>
    rebindFormulaCellsNow(collectTrackedDependents(registry, keys), formulaChangedCount)

  const rebindFormulasForSheetNow = (sheetName: string, formulaChangedCount: number, candidates?: readonly number[] | U32): number => {
    if (candidates) {
      for (let index = 0; index < candidates.length; index += 1) {
        const cellIndex = candidates[index]!
        const formula = args.serviceArgs.state.formulas.get(cellIndex)
        if (!formula) {
          continue
        }
        const ownerSheetName = args.serviceArgs.state.workbook.getSheetNameById(
          args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]!,
        )
        if (!ownerSheetName) {
          continue
        }
        const touchesSheet = formula.compiled.deps.some((dep) => {
          if (!dep.includes('!')) {
            return false
          }
          const [qualifiedSheet] = dep.split('!')
          return qualifiedSheet?.replace(/^'(.*)'$/, '$1') === sheetName
        })
        if (!touchesSheet) {
          continue
        }
        args.bindFormulaNow(cellIndex, ownerSheetName, getRuntimeFormulaSource(formula))
        formulaChangedCount = args.serviceArgs.markFormulaChanged(cellIndex, formulaChangedCount)
      }
      return formulaChangedCount
    }

    args.serviceArgs.state.formulas.forEach((formula, cellIndex) => {
      if (!formula) {
        return
      }
      const ownerSheetName = args.serviceArgs.state.workbook.getSheetNameById(
        args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]!,
      )
      if (!ownerSheetName) {
        return
      }
      const touchesSheet = formula.compiled.deps.some((dep) => {
        if (!dep.includes('!')) {
          return false
        }
        const [qualifiedSheet] = dep.split('!')
        return qualifiedSheet?.replace(/^'(.*)'$/, '$1') === sheetName
      })
      if (!touchesSheet) {
        return
      }
      args.bindFormulaNow(cellIndex, ownerSheetName, getRuntimeFormulaSource(formula))
      formulaChangedCount = args.serviceArgs.markFormulaChanged(cellIndex, formulaChangedCount)
    })

    return formulaChangedCount
  }

  return {
    rebindFormulaCellsNow,
    rebindTrackedDependentsNow,
    rebindFormulasForSheetNow,
  }
}
