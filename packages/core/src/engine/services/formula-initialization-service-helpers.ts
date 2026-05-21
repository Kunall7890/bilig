import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineFormulaInitializationServiceArgs } from './formula-initialization-service-types.js'

type FormulaInitializationState = EngineFormulaInitializationServiceArgs['state']

export function scanFormulaInitializationCycleMembers(state: FormulaInitializationState): boolean {
  addEngineCounter(state.counters, 'cycleFormulaScans')
  let found = false
  state.formulas.forEach((_formula, cellIndex) => {
    if (((state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
      found = true
    }
  })
  return found
}

export function createFormulaInitializationSheetNameResolver(state: FormulaInitializationState): (sheetId: number) => string {
  const sheetNameById = new Map<number, string>()
  return (sheetId: number): string => {
    const cached = sheetNameById.get(sheetId)
    if (cached !== undefined) {
      return cached
    }
    const sheet = state.workbook.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    sheetNameById.set(sheetId, sheet.name)
    return sheet.name
  }
}
