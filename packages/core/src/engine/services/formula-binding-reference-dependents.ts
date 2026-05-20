import { tableDependencyKey } from '../../engine-metadata-utils.js'
import { normalizeDefinedName } from '../../workbook-store.js'
import { collectTrackedDependents } from './formula-binding-dependency-helpers.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

type FormulaReverseState = CreateEngineFormulaBindingServiceArgs['reverseState']

export function collectDefinedNameDependents(reverseState: FormulaReverseState, names: readonly string[]): number[] {
  return collectTrackedDependents(
    reverseState.reverseDefinedNameEdges,
    names.map((name) => normalizeDefinedName(name)),
  )
}

export function collectTableDependents(reverseState: FormulaReverseState, tableNames: readonly string[]): number[] {
  return collectTrackedDependents(
    reverseState.reverseTableEdges,
    tableNames.map((name) => tableDependencyKey(name)),
  )
}
