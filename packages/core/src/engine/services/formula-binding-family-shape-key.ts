import type { FormulaMode } from '@bilig/protocol'
import { buildFormulaFamilyShapeKey } from '../../formula/formula-family-deps.js'
import type { RuntimeFormula } from '../runtime-state.js'

interface CachedFormulaFamilyShapeKey {
  readonly mode: FormulaMode
  readonly depsLength: number
  readonly symbolicRangesLength: number
  readonly symbolicNamesLength: number
  readonly symbolicTablesLength: number
  readonly symbolicSpillsLength: number
  readonly dependencyCount: number
  readonly rangeDependencyCount: number
  readonly directAggregateKind: string | undefined
  readonly directLookupKind: string | undefined
  readonly directScalarKind: string | undefined
  readonly directCriteriaKind: string | undefined
  readonly shapeKey: string
}

export type FormulaBindingFamilyShapeKeyCache = Map<number, CachedFormulaFamilyShapeKey>

export function getFormulaBindingFamilyShapeKey(cache: FormulaBindingFamilyShapeKeyCache, formula: RuntimeFormula): string {
  const templateId = formula.templateId
  const directAggregateKind = formula.directAggregate?.aggregateKind
  const directLookupKind = formula.directLookup?.kind
  const directScalarKind = formula.directScalar?.kind
  const directCriteriaKind = formula.directCriteria?.aggregateKind
  if (templateId !== undefined) {
    const cached = cache.get(templateId)
    if (
      cached &&
      cached.mode === formula.compiled.mode &&
      cached.depsLength === formula.compiled.deps.length &&
      cached.symbolicRangesLength === formula.compiled.symbolicRanges.length &&
      cached.symbolicNamesLength === formula.compiled.symbolicNames.length &&
      cached.symbolicTablesLength === formula.compiled.symbolicTables.length &&
      cached.symbolicSpillsLength === formula.compiled.symbolicSpills.length &&
      cached.dependencyCount === formula.dependencyIndices.length &&
      cached.rangeDependencyCount === formula.rangeDependencies.length &&
      cached.directAggregateKind === directAggregateKind &&
      cached.directLookupKind === directLookupKind &&
      cached.directScalarKind === directScalarKind &&
      cached.directCriteriaKind === directCriteriaKind
    ) {
      return cached.shapeKey
    }
  }
  const shapeKey = buildFormulaFamilyShapeKey({
    compiled: formula.compiled,
    dependencyCount: formula.dependencyIndices.length,
    rangeDependencyCount: formula.rangeDependencies.length,
    directAggregateKind,
    directLookupKind,
    directScalarKind,
    directCriteriaKind,
  })
  if (templateId !== undefined) {
    cache.set(templateId, {
      mode: formula.compiled.mode,
      depsLength: formula.compiled.deps.length,
      symbolicRangesLength: formula.compiled.symbolicRanges.length,
      symbolicNamesLength: formula.compiled.symbolicNames.length,
      symbolicTablesLength: formula.compiled.symbolicTables.length,
      symbolicSpillsLength: formula.compiled.symbolicSpills.length,
      dependencyCount: formula.dependencyIndices.length,
      rangeDependencyCount: formula.rangeDependencies.length,
      directAggregateKind,
      directLookupKind,
      directScalarKind,
      directCriteriaKind,
      shapeKey,
    })
  }
  return shapeKey
}
