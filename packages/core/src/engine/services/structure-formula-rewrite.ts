import {
  type CompiledFormula,
  parseCellAddress,
  rewriteCompiledFormulaForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import {
  rewriteTemplateForStructuralTransform,
  retargetStructurallyRewrittenTemplateInstance,
  type StructurallyRewrittenTemplate,
} from '../../formula/structural-retargeting.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { CreateEngineStructureServiceArgs } from './structure-service-types.js'

export type StructuralFormulaRewriteCache = Map<string, StructurallyRewrittenTemplate | null>

export function rewriteStructuralFormulaCompiled(
  formula: RuntimeFormula,
  ownerSheetName: string,
  sheetName: string,
  transform: StructuralAxisTransform,
): ReturnType<typeof rewriteCompiledFormulaForStructuralTransform> | undefined {
  if (
    formula.directAggregate &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0
  ) {
    const directAggregateCandidate = formula.compiled.directAggregateCandidate
    const rangeIndex = directAggregateCandidate?.symbolicRangeIndex
    const parsedRange = rangeIndex === undefined ? undefined : formula.compiled.parsedSymbolicRanges?.[rangeIndex]
    if (
      directAggregateCandidate &&
      rangeIndex !== undefined &&
      parsedRange &&
      parsedRange.refKind === 'cells' &&
      (parsedRange.sheetName ?? ownerSheetName) === sheetName
    ) {
      const candidateRangeIndex = rangeIndex
      const nextRange = rewriteRangeForStructuralTransform(parsedRange.startAddress, parsedRange.endAddress, transform)
      if (nextRange) {
        const rangePrefix = parsedRange.address.includes('!') ? parsedRange.address.slice(0, parsedRange.address.lastIndexOf('!') + 1) : ''
        const nextAddress = `${rangePrefix}${nextRange.startAddress}:${nextRange.endAddress}`
        const nextStart = parseCellAddress(nextRange.startAddress, parsedRange.sheetName ?? ownerSheetName)
        const nextEnd = parseCellAddress(nextRange.endAddress, parsedRange.sheetName ?? ownerSheetName)
        const nextParsedRange = {
          ...parsedRange,
          address: nextAddress,
          startAddress: nextRange.startAddress,
          endAddress: nextRange.endAddress,
          startRow: nextStart.row,
          endRow: nextEnd.row,
          startCol: nextStart.col,
          endCol: nextEnd.col,
        }
        const nextParsedSymbolicRanges = formula.compiled.parsedSymbolicRanges?.slice()
        if (nextParsedSymbolicRanges) {
          nextParsedSymbolicRanges[candidateRangeIndex] = nextParsedRange
        }
        const nextParsedDeps = formula.compiled.parsedDeps?.map((dependency) =>
          dependency.kind === 'range' && dependency.address === parsedRange.address
            ? {
                ...dependency,
                address: nextAddress,
                startAddress: nextRange.startAddress,
                endAddress: nextRange.endAddress,
                startRow: nextStart.row,
                endRow: nextEnd.row,
                startCol: nextStart.col,
                endCol: nextEnd.col,
              }
            : dependency,
        )
        return {
          source: `${directAggregateCandidate.callee}(${nextAddress})`,
          compiled: {
            ...formula.compiled,
            source: `${directAggregateCandidate.callee}(${nextAddress})`,
            astMatchesSource: false,
            deps: formula.compiled.deps.map((dependency) => (dependency === parsedRange.address ? nextAddress : dependency)),
            symbolicRanges: formula.compiled.symbolicRanges.map((range, index) => (index === candidateRangeIndex ? nextAddress : range)),
            ...(nextParsedDeps ? { parsedDeps: nextParsedDeps } : {}),
            ...(nextParsedSymbolicRanges ? { parsedSymbolicRanges: nextParsedSymbolicRanges } : {}),
          },
          reusedProgram: true,
        }
      }
    }
  }
  const rewritten = rewriteCompiledFormulaForStructuralTransform(formula.compiled, ownerSheetName, sheetName, transform)
  if (rewritten.source === formula.source) {
    return undefined
  }
  return rewritten
}

export function rewriteFormulaSourceFallback(
  source: string,
  ownerSheetName: string,
  sheetName: string,
  transform: StructuralAxisTransform,
): string {
  return rewriteFormulaForStructuralTransform(source, ownerSheetName, sheetName, transform)
}

export function structuralRewritePreservesDirectCellDependencies(
  args: CreateEngineStructureServiceArgs,
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula },
  ownerSheetName: string,
): boolean {
  if (
    formula.directLookup !== undefined ||
    formula.directAggregate !== undefined ||
    formula.directCriteria !== undefined ||
    formula.rangeDependencies.length !== 0 ||
    formula.dependencyIndices.length === 0
  ) {
    return false
  }
  const parsedDeps = rewritten.compiled.parsedDeps
  if (!parsedDeps || parsedDeps.length !== formula.dependencyIndices.length) {
    return false
  }
  for (let index = 0; index < parsedDeps.length; index += 1) {
    const dependency = parsedDeps[index]!
    if (dependency.kind !== 'cell') {
      return false
    }
    const dependencySheetName = dependency.sheetName ?? ownerSheetName
    const dependencyCellIndex = args.state.workbook.getCellIndex(dependencySheetName, dependency.address)
    if (dependencyCellIndex === undefined || dependencyCellIndex !== formula.dependencyIndices[index]) {
      return false
    }
  }
  return true
}

export function rewriteFormulaFromTemplate(
  cache: StructuralFormulaRewriteCache,
  formula: RuntimeFormula,
  representative: {
    readonly templateId: number
    readonly ownerSheetName: string
    readonly targetSheetName: string
    readonly representativeRow: number
    readonly representativeCol: number
    readonly ownerRow: number
    readonly ownerCol: number
  },
  targetSheetName: string,
  transform: StructuralAxisTransform,
): { source: string; compiled: CompiledFormula; reusedProgram: boolean } | undefined {
  if (formula.directAggregate !== undefined || formula.directCriteria !== undefined) {
    return undefined
  }
  if (formula.compiled.astMatchesSource === false) {
    return undefined
  }
  const cacheKey =
    `${representative.templateId}:${representative.ownerSheetName}:${targetSheetName}:${transform.kind}:${transform.axis}:${transform.start}:${transform.count}:` +
    `${transform.kind === 'move' ? transform.target : ''}`
  let rewrittenTemplate = cache.get(cacheKey)
  if (rewrittenTemplate === undefined) {
    rewrittenTemplate =
      rewriteTemplateForStructuralTransform({
        template: {
          id: representative.templateId,
          templateKey: `runtime-template-${representative.templateId}`,
          baseSource: formula.source,
          baseRow: representative.representativeRow,
          baseCol: representative.representativeCol,
          compiled: formula.compiled,
        },
        ownerSheetName: representative.ownerSheetName,
        targetSheetName,
        transform,
      }) ?? null
    cache.set(cacheKey, rewrittenTemplate)
  }
  if (rewrittenTemplate === null) {
    return undefined
  }

  try {
    return retargetStructurallyRewrittenTemplateInstance({
      rewrittenTemplate,
      ownerRow: representative.ownerRow,
      ownerCol: representative.ownerCol,
    })
  } catch {
    return undefined
  }
}
