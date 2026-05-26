import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { InitialFormulaCellIndexList } from './formula-initialization-refs.js'
import type { EngineFormulaInitializationServiceArgs } from './formula-initialization-service-types.js'

export type InitialPrefixAggregateKind = 'sum' | 'count' | 'average' | 'min' | 'max'

export interface InitialPrefixAggregateGroup {
  readonly sheetName: string
  readonly col: number
  readonly colEnd: number
  readonly aggregateKind: InitialPrefixAggregateKind
  maxRowEnd: number
  lastRowEnd: number
  formulasAreOrdered: boolean
  readonly formulas: Array<{ cellIndex: number; rowEnd: number; resultOffset?: number }>
}

// Native marshaling loses below the existing JS direct-evaluation limit; use it only to avoid the over-limit recalc path.
const NATIVE_INITIAL_PREFIX_AGGREGATE_MIN_FORMULAS = 16_385
const NATIVE_INITIAL_PREFIX_AGGREGATE_MAX_VALUES = 2_000_000
const NATIVE_DIRECT_AGGREGATE_OP_SUM = 1
const NATIVE_DIRECT_AGGREGATE_OP_AVERAGE = 2
const NATIVE_DIRECT_AGGREGATE_OP_COUNT = 3
const NATIVE_DIRECT_AGGREGATE_OP_MIN = 4
const NATIVE_DIRECT_AGGREGATE_OP_MAX = 5

type InitialPrefixAggregateArgs = Pick<EngineFormulaInitializationServiceArgs, 'state' | 'checkEvaluationBudget'>

export function canEvaluateInitialPrefixAggregateGroupsNatively(
  args: InitialPrefixAggregateArgs,
  orderedCellIndices: InitialFormulaCellIndexList,
  options?: { readonly requireAllCells?: boolean },
): boolean {
  const collected = collectInitialPrefixAggregateGroups(args, orderedCellIndices)
  if (!collected || collected.groups.size === 0) {
    return false
  }
  if (options?.requireAllCells === true && collected.matchedFormulaCount !== orderedCellIndices.length) {
    return false
  }
  for (const group of collected.groups.values()) {
    if (!canUseNativeInitialPrefixAggregateGroup(args, group)) {
      return false
    }
  }
  return true
}

export function evaluateInitialPrefixAggregateGroups(
  args: InitialPrefixAggregateArgs,
  orderedCellIndices: InitialFormulaCellIndexList,
  pushChangedCellIndex: (cellIndex: number) => void,
  writeFormulaValue: (cellIndex: number, value: CellValue) => void,
): Set<number> | undefined {
  const collected = collectInitialPrefixAggregateGroups(args, orderedCellIndices)
  if (!collected || collected.groups.size === 0) {
    return undefined
  }

  const handled = new Set<number>()
  collected.groups.forEach((group) => {
    if (tryEvaluateNativeInitialPrefixAggregateGroup(args, group, handled, pushChangedCellIndex, writeFormulaValue)) {
      return
    }
    evaluateInitialPrefixAggregateGroupInJs(args, group, handled, pushChangedCellIndex, writeFormulaValue)
  })
  return handled.size === 0 ? undefined : handled
}

function collectInitialPrefixAggregateGroups(
  args: InitialPrefixAggregateArgs,
  orderedCellIndices: InitialFormulaCellIndexList,
): { readonly groups: Map<string, InitialPrefixAggregateGroup>; readonly matchedFormulaCount: number } | undefined {
  const groups = new Map<string, InitialPrefixAggregateGroup>()
  let matchedFormulaCount = 0
  for (let index = 0; index < orderedCellIndices.length; index += 1) {
    args.checkEvaluationBudget()
    const cellIndex = orderedCellIndices[index]!
    const formula = args.state.formulas.get(cellIndex)
    const aggregate = formula?.directAggregate
    if (!formula || !aggregate || aggregate.rowStart !== 0 || formula.dependencyIndices.length !== 0) {
      continue
    }
    matchedFormulaCount += 1
    const key = `${aggregate.sheetName}\t${aggregate.col}\t${aggregate.colEnd}\t${aggregate.aggregateKind}`
    let group = groups.get(key)
    if (!group) {
      group = {
        sheetName: aggregate.sheetName,
        col: aggregate.col,
        colEnd: aggregate.colEnd,
        aggregateKind: aggregate.aggregateKind,
        maxRowEnd: aggregate.rowEnd,
        lastRowEnd: aggregate.rowEnd,
        formulasAreOrdered: true,
        formulas: [],
      }
      groups.set(key, group)
    } else {
      group.maxRowEnd = Math.max(group.maxRowEnd, aggregate.rowEnd)
      if (aggregate.rowEnd < group.lastRowEnd) {
        group.formulasAreOrdered = false
      }
      group.lastRowEnd = aggregate.rowEnd
    }
    group.formulas.push({
      cellIndex,
      rowEnd: aggregate.rowEnd,
      ...(aggregate.resultOffset !== undefined ? { resultOffset: aggregate.resultOffset } : {}),
    })
  }
  return groups.size === 0 ? undefined : { groups, matchedFormulaCount }
}

function canUseNativeInitialPrefixAggregateGroup(args: InitialPrefixAggregateArgs, group: InitialPrefixAggregateGroup): boolean {
  if (group.formulas.length < NATIVE_INITIAL_PREFIX_AGGREGATE_MIN_FORMULAS || group.colEnd < group.col) {
    return false
  }
  const sheet = args.state.workbook.getSheet(group.sheetName)
  if (!sheet || !args.state.wasm.initSyncIfPossible()) {
    return false
  }
  const rowCount = group.maxRowEnd + 1
  const colCount = group.colEnd - group.col + 1
  const valueCount = rowCount * colCount
  if (
    rowCount <= 0 ||
    colCount <= 0 ||
    !Number.isSafeInteger(valueCount) ||
    valueCount <= 0 ||
    valueCount > NATIVE_INITIAL_PREFIX_AGGREGATE_MAX_VALUES
  ) {
    return false
  }
  for (let row = 0; row <= group.maxRowEnd; row += 1) {
    args.checkEvaluationBudget()
    for (let col = group.col; col <= group.colEnd; col += 1) {
      const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
      if (memberCellIndex !== -1 && ((args.state.workbook.cellStore.flags[memberCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
        return false
      }
    }
  }
  return true
}

function tryEvaluateNativeInitialPrefixAggregateGroup(
  args: InitialPrefixAggregateArgs,
  group: InitialPrefixAggregateGroup,
  handled: Set<number>,
  pushChangedCellIndex: (cellIndex: number) => void,
  writeFormulaValue: (cellIndex: number, value: CellValue) => void,
): boolean {
  if (!canUseNativeInitialPrefixAggregateGroup(args, group)) {
    return false
  }
  const sheet = args.state.workbook.getSheet(group.sheetName)
  if (!sheet) {
    return false
  }
  const formulas = orderedGroupFormulas(group)
  const rowCount = group.maxRowEnd + 1
  const colCount = group.colEnd - group.col + 1
  const valueCount = rowCount * colCount
  const tags = new Uint8Array(valueCount)
  const numbers = new Float64Array(valueCount)
  const errors = new Uint16Array(valueCount)
  let offset = 0
  for (let row = 0; row < rowCount; row += 1) {
    args.checkEvaluationBudget()
    for (let col = group.col; col <= group.colEnd; col += 1) {
      const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
      if (memberCellIndex === -1) {
        tags[offset] = ValueTag.Empty
      } else {
        if (((args.state.workbook.cellStore.flags[memberCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
          return false
        }
        tags[offset] = (args.state.workbook.cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
        numbers[offset] = args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0
        errors[offset] = (args.state.workbook.cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None
      }
      offset += 1
    }
  }

  const formulaRowEnds = new Uint32Array(formulas.length)
  const resultOffsets = new Float64Array(formulas.length)
  for (let index = 0; index < formulas.length; index += 1) {
    formulaRowEnds[index] = formulas[index]!.rowEnd
    resultOffsets[index] = formulas[index]!.resultOffset ?? 0
  }
  const outTags = new Uint8Array(formulas.length)
  const outNumbers = new Float64Array(formulas.length)
  const outErrors = new Uint16Array(formulas.length)
  if (
    !args.state.wasm.evalAnchoredPrefixAggregateBatch({
      aggregateKind: nativeDirectAggregateKind(group.aggregateKind),
      tags,
      numbers,
      errors,
      rowCount,
      colCount,
      formulaRowEnds,
      resultOffsets,
      outTags,
      outNumbers,
      outErrors,
    })
  ) {
    return false
  }

  for (let index = 0; index < formulas.length; index += 1) {
    const formula = formulas[index]!
    const tag = (outTags[index] as ValueTag | undefined) ?? ValueTag.Empty
    const value =
      tag === ValueTag.Number
        ? { tag: ValueTag.Number as const, value: outNumbers[index] ?? 0 }
        : tag === ValueTag.Error
          ? { tag: ValueTag.Error as const, code: (outErrors[index] as ErrorCode | undefined) ?? ErrorCode.None }
          : { tag: ValueTag.Empty as const }
    writeFormulaValue(formula.cellIndex, value)
    handled.add(formula.cellIndex)
    pushChangedCellIndex(formula.cellIndex)
  }
  addEngineCounter(args.state.counters, 'nativeDirectAggregatePrefixEvaluations', formulas.length)
  return true
}

function evaluateInitialPrefixAggregateGroupInJs(
  args: InitialPrefixAggregateArgs,
  group: InitialPrefixAggregateGroup,
  handled: Set<number>,
  pushChangedCellIndex: (cellIndex: number) => void,
  writeFormulaValue: (cellIndex: number, value: CellValue) => void,
): void {
  const sheet = args.state.workbook.getSheet(group.sheetName)
  if (!sheet) {
    return
  }
  const formulas = orderedGroupFormulas(group)
  let sum = 0
  let count = 0
  let averageCount = 0
  let errorCode = ErrorCode.None
  let errorCount = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  let formulaIndex = 0
  let encounteredFormulaMember = false
  for (let row = 0; row <= group.maxRowEnd && !encounteredFormulaMember; row += 1) {
    args.checkEvaluationBudget()
    for (let col = group.col; col <= group.colEnd; col += 1) {
      const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
      if (memberCellIndex !== -1) {
        if (((args.state.workbook.cellStore.flags[memberCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
          encounteredFormulaMember = true
          break
        }
        const tag = (args.state.workbook.cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
        if (tag === ValueTag.Number) {
          const numeric = args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0
          sum += numeric
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, numeric)
          maximum = Math.max(maximum, numeric)
        } else if (tag === ValueTag.Boolean) {
          const numeric = (args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
          sum += numeric
          count += 1
          averageCount += 1
          minimum = Math.min(minimum, numeric)
          maximum = Math.max(maximum, numeric)
        } else if (tag === ValueTag.Empty) {
          minimum = Math.min(minimum, 0)
          maximum = Math.max(maximum, 0)
        } else if (tag === ValueTag.Error) {
          errorCode ||= (args.state.workbook.cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None
          errorCount += 1
        }
      }
    }
    while (formulaIndex < formulas.length && formulas[formulaIndex]!.rowEnd <= row) {
      const formula = formulas[formulaIndex]!
      const aggregateValue = aggregateValueFromState(group.aggregateKind, sum, count, averageCount, errorCode, errorCount, minimum, maximum)
      const value =
        formula.resultOffset !== undefined && aggregateValue.tag === ValueTag.Number
          ? { tag: ValueTag.Number as const, value: aggregateValue.value + formula.resultOffset }
          : aggregateValue
      writeFormulaValue(formula.cellIndex, value)
      handled.add(formula.cellIndex)
      pushChangedCellIndex(formula.cellIndex)
      formulaIndex += 1
    }
  }
}

function orderedGroupFormulas(group: InitialPrefixAggregateGroup): readonly InitialPrefixAggregateGroup['formulas'][number][] {
  return group.formulasAreOrdered ? group.formulas : group.formulas.toSorted((left, right) => left.rowEnd - right.rowEnd)
}

function aggregateValueFromState(
  aggregateKind: InitialPrefixAggregateKind,
  sum: number,
  count: number,
  averageCount: number,
  errorCode: ErrorCode,
  errorCount: number,
  minimum: number,
  maximum: number,
): CellValue {
  return aggregateKind === 'sum'
    ? errorCount > 0 && errorCode !== ErrorCode.None
      ? { tag: ValueTag.Error as const, code: errorCode }
      : { tag: ValueTag.Number as const, value: sum }
    : aggregateKind === 'count'
      ? { tag: ValueTag.Number as const, value: count }
      : aggregateKind === 'average'
        ? errorCount > 0 && errorCode !== ErrorCode.None
          ? { tag: ValueTag.Error as const, code: errorCode }
          : averageCount === 0
            ? { tag: ValueTag.Error as const, code: ErrorCode.Div0 }
            : { tag: ValueTag.Number as const, value: sum / averageCount }
        : aggregateKind === 'min'
          ? { tag: ValueTag.Number as const, value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum }
          : { tag: ValueTag.Number as const, value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum }
}

function nativeDirectAggregateKind(kind: InitialPrefixAggregateKind): number {
  switch (kind) {
    case 'sum':
      return NATIVE_DIRECT_AGGREGATE_OP_SUM
    case 'average':
      return NATIVE_DIRECT_AGGREGATE_OP_AVERAGE
    case 'count':
      return NATIVE_DIRECT_AGGREGATE_OP_COUNT
    case 'min':
      return NATIVE_DIRECT_AGGREGATE_OP_MIN
    case 'max':
      return NATIVE_DIRECT_AGGREGATE_OP_MAX
  }
}
