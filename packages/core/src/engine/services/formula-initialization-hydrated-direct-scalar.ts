import { addEngineCounter } from '../../perf/engine-counters.js'
import type {
  EngineFormulaInitializationServiceArgs,
  HydratedPreparedFormulaInitializationRef,
} from './formula-initialization-service-types.js'
import type { FreshDirectAggregateFormulaBindingMember } from './formula-binding-service-types.js'
import { unwrapDirectScalarBinaryNode } from './formula-binding-direct-scalar.js'

interface FreshDirectAggregateBindingSource {
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: HydratedPreparedFormulaInitializationRef['compiled']
  readonly templateId?: number
}

export function tryBindHydratedFreshDirectFormula(
  serviceArgs: EngineFormulaInitializationServiceArgs,
  hadExistingFormulas: boolean,
  cellIndex: number,
  ownerSheetName: string,
  ref: HydratedPreparedFormulaInitializationRef,
): boolean {
  return (
    tryBindHydratedFreshDirectScalarFormula(serviceArgs, hadExistingFormulas, cellIndex, ownerSheetName, ref) ||
    tryBindHydratedFreshDirectAggregateFormula(serviceArgs, hadExistingFormulas, cellIndex, ownerSheetName, ref)
  )
}

function tryBindHydratedFreshDirectScalarFormula(
  serviceArgs: EngineFormulaInitializationServiceArgs,
  hadExistingFormulas: boolean,
  cellIndex: number,
  ownerSheetName: string,
  ref: HydratedPreparedFormulaInitializationRef,
): boolean {
  if (
    hadExistingFormulas ||
    serviceArgs.bindFreshDirectScalarFormulaRun === undefined ||
    ref.templateId === undefined ||
    ref.preserveCachedValueOnFullRecalc === true ||
    !canUseFreshDirectScalarFormulaBinding(ref.compiled)
  ) {
    return false
  }
  serviceArgs.bindFreshDirectScalarFormulaRun({
    sheetId: ref.sheetId,
    ownerSheetName,
    cellIndex,
    member: {
      row: ref.row,
      col: ref.col,
      source: ref.source,
      compiled: ref.compiled,
      templateId: ref.templateId,
    },
  })
  addEngineCounter(serviceArgs.state.counters, 'runtimeHydratedDirectScalarFastBindings')
  return true
}

function tryBindHydratedFreshDirectAggregateFormula(
  serviceArgs: EngineFormulaInitializationServiceArgs,
  hadExistingFormulas: boolean,
  cellIndex: number,
  ownerSheetName: string,
  ref: HydratedPreparedFormulaInitializationRef,
): boolean {
  if (
    hadExistingFormulas ||
    serviceArgs.bindFreshDirectAggregateFormulaRun === undefined ||
    ref.templateId === undefined ||
    ref.preserveCachedValueOnFullRecalc === true
  ) {
    return false
  }
  const member = buildFreshDirectAggregateMember(ownerSheetName, ref)
  if (member === undefined) {
    return false
  }
  serviceArgs.bindFreshDirectAggregateFormulaRun({
    sheetId: ref.sheetId,
    ownerSheetName,
    cellIndex,
    member,
  })
  addEngineCounter(serviceArgs.state.counters, 'runtimeHydratedDirectAggregateFastBindings')
  return true
}

export function buildFreshDirectAggregateMember(
  ownerSheetName: string,
  ref: FreshDirectAggregateBindingSource,
): FreshDirectAggregateFormulaBindingMember | undefined {
  const compiled = ref.compiled
  if (
    compiled.volatile ||
    compiled.producesSpill ||
    compiled.symbolicRanges.length !== 1 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0
  ) {
    return undefined
  }
  const aggregate = compiled.directAggregateCandidate
  const range = aggregate === undefined ? undefined : compiled.parsedSymbolicRanges?.[aggregate.symbolicRangeIndex]
  const aggregateSheetName = range?.sheetName ?? ownerSheetName
  if (
    aggregate === undefined ||
    range === undefined ||
    ref.templateId === undefined ||
    range.refKind !== 'cells' ||
    range.startRow > range.endRow ||
    range.startCol > range.endCol ||
    (aggregateSheetName === ownerSheetName &&
      range.startRow <= ref.row &&
      ref.row <= range.endRow &&
      range.startCol <= ref.col &&
      ref.col <= range.endCol)
  ) {
    return undefined
  }
  return {
    row: ref.row,
    col: ref.col,
    source: ref.source,
    compiled,
    templateId: ref.templateId,
    aggregateKind: aggregate.aggregateKind,
    aggregateSheetName,
    aggregateRowStart: range.startRow,
    aggregateRowEnd: range.endRow,
    aggregateColStart: range.startCol,
    aggregateColEnd: range.endCol,
    resultOffset: normalizeDirectAggregateResultOffset(aggregate.resultOffset),
  }
}

function normalizeDirectAggregateResultOffset(offset: number | undefined): number | undefined {
  return offset === undefined || Object.is(offset, 0) ? undefined : offset
}

export function canUseFreshDirectScalarFormulaBinding(compiled: HydratedPreparedFormulaInitializationRef['compiled']): boolean {
  if (
    compiled.volatile ||
    compiled.producesSpill ||
    compiled.symbolicRanges.length !== 0 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0 ||
    hasExternalWorkbookCellReference(compiled)
  ) {
    return false
  }
  const node = unwrapDirectScalarBinaryNode(compiled.optimizedAst).node
  if (node.kind === 'BinaryExpr' && (node.operator === '+' || node.operator === '-' || node.operator === '*' || node.operator === '/')) {
    return isFreshDirectScalarOperand(node.left) && isFreshDirectScalarOperand(node.right)
  }
  return (
    node.kind === 'CallExpr' &&
    node.callee.trim().toUpperCase() === 'ABS' &&
    node.args.length === 1 &&
    isFreshDirectScalarOperand(node.args[0]!)
  )
}

function isFreshDirectScalarOperand(node: HydratedPreparedFormulaInitializationRef['compiled']['optimizedAst']): boolean {
  return node.kind === 'NumberLiteral' || node.kind === 'CellRef'
}

function hasExternalWorkbookCellReference(compiled: HydratedPreparedFormulaInitializationRef['compiled']): boolean {
  const parsedRefs = compiled.parsedSymbolicRefs
  if (parsedRefs !== undefined && parsedRefs.length === compiled.symbolicRefs.length) {
    for (let index = 0; index < parsedRefs.length; index += 1) {
      const sheetName = parsedRefs[index]?.sheetName
      if (sheetName !== undefined && isExternalWorkbookSheetName(sheetName)) {
        return true
      }
    }
    return false
  }
  return compiled.symbolicRefs.some(symbolicRefStartsWithExternalWorkbook)
}

function isExternalWorkbookSheetName(sheetName: string): boolean {
  return /^\[\d+\]/.test(sheetName)
}

function symbolicRefStartsWithExternalWorkbook(reference: string): boolean {
  const trimmed = reference.trim()
  return /^\[\d+\]/.test(trimmed) || /^'\[\d+\]/.test(trimmed)
}
