import { ErrorCode, MAX_COLS, MAX_ROWS, ValueTag, type CellValue } from '@bilig/protocol'
import { formatAddress, parseCellAddress, parseRangeAddress } from './addressing.js'
import { getBuiltin } from './builtins.js'
import { getLookupBuiltin, type RangeBuiltinArgument } from './builtins/lookup.js'
import { evaluateGroupBy, evaluatePivotBy } from './group-pivot-evaluator.js'
import {
  aggregateOptionIgnoresErrors,
  aggregateOptionIgnoresHiddenRows,
  aggregateOptionIgnoresNestedRollups,
  collectAggregateCandidates,
  filterNestedRollupCandidates,
  firstErrorValue,
  nestedAggregateCallees,
  nestedSubtotalCallees,
} from './js-evaluator-aggregate-special-calls.js'
import { isArrayValue } from './runtime-values.js'
import { evaluateIndirectWorkbookSpecialCall, type WorkbookReferenceCallDeps } from './js-evaluator-workbook-reference-calls.js'
import type { EvaluationContext, ReferenceOperand, StackValue } from './js-evaluator.js'

interface MatrixLikeValue {
  rows: number
  cols: number
  values: readonly CellValue[]
}

interface WorkbookSpecialCallDeps extends WorkbookReferenceCallDeps {
  error: (code: ErrorCode) => CellValue
  stackScalar: (value: CellValue, blankReference?: boolean) => StackValue
  toStringValue: (value: CellValue) => string
  isSingleCellValue: (value: StackValue) => CellValue | undefined
  matrixFromStackValue: (value: StackValue) => MatrixLikeValue | undefined
  scalarIntegerArgument: (value: StackValue | undefined) => number | undefined
  vectorIntegerArgument: (value: StackValue | undefined) => number[] | undefined
  aggregateRangeSubset: (
    functionArg: StackValue,
    subset: readonly CellValue[],
    context: EvaluationContext,
    totalSet?: readonly CellValue[],
  ) => CellValue
}

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}

type WholeAxisRefKind = 'rows' | 'cols'
type WholeAxisRangeAddress = Extract<ReturnType<typeof parseRangeAddress>, { kind: WholeAxisRefKind }>

interface WholeAxisReference {
  sheetName: string
  start: string
  end: string
  refKind: WholeAxisRefKind
  parsed: WholeAxisRangeAddress
}

interface ReferenceBounds {
  sheetName: string
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

type ScalarArgumentResult = { kind: 'ok'; value: CellValue | undefined } | { kind: 'error'; value: CellValue }
type IntegerArgumentResult = { kind: 'omitted' } | { kind: 'ok'; value: number } | { kind: 'error'; value: CellValue }

function isWholeAxisRefKind(refKind: 'cells' | 'rows' | 'cols' | undefined): refKind is WholeAxisRefKind {
  return refKind === 'rows' || refKind === 'cols'
}

function wholeAxisReferenceFromArg(
  value: StackValue | undefined,
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
  deps: WorkbookSpecialCallDeps,
): WholeAxisReference | undefined {
  const refKind = isWholeAxisRefKind(ref?.refKind)
    ? ref.refKind
    : value?.kind === 'range' && isWholeAxisRefKind(value.refKind)
      ? value.refKind
      : undefined
  if (!refKind) {
    return undefined
  }

  const start = (ref?.kind === 'range' ? ref.start : undefined) ?? (value?.kind === 'range' ? value.start : undefined)
  const end = (ref?.kind === 'range' ? ref.end : undefined) ?? (value?.kind === 'range' ? value.end : undefined)
  if (!start || !end) {
    return undefined
  }

  const sheetName = deps.referenceSheetName(ref, context) ?? (value?.kind === 'range' ? value.sheetName : undefined) ?? context.sheetName
  try {
    const parsed = parseRangeAddress(`${start}:${end}`, sheetName)
    if (parsed.kind !== refKind) {
      return undefined
    }
    return {
      sheetName,
      start,
      end,
      refKind,
      parsed,
    }
  } catch {
    return undefined
  }
}

function scalarLookupArgument(value: StackValue | undefined, deps: WorkbookSpecialCallDeps): ScalarArgumentResult {
  if (!value || (value.kind === 'omitted' && value.source === 'argument')) {
    return { kind: 'ok', value: undefined }
  }
  const scalar = deps.isSingleCellValue(value)
  if (!scalar) {
    return { kind: 'error', value: valueError() }
  }
  return { kind: 'ok', value: scalar }
}

function integerIndexArgument(value: StackValue | undefined, deps: WorkbookSpecialCallDeps): IntegerArgumentResult {
  if (!value || (value.kind === 'omitted' && value.source === 'argument')) {
    return { kind: 'omitted' }
  }
  const scalar = deps.isSingleCellValue(value)
  if (!scalar) {
    return { kind: 'error', value: valueError() }
  }
  if (scalar.tag === ValueTag.Error) {
    return { kind: 'error', value: scalar }
  }
  const integer = deps.scalarIntegerArgument(value)
  return integer === undefined ? { kind: 'error', value: valueError() } : { kind: 'ok', value: integer }
}

function referenceBoundsFromArg(
  value: StackValue | undefined,
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
): ReferenceBounds | undefined {
  const sheetName = ref?.sheetName ?? (value?.kind === 'range' ? value.sheetName : undefined) ?? context.sheetName
  if (ref?.kind === 'cell' && ref.address) {
    const parsed = parseCellAddress(ref.address, sheetName)
    return {
      sheetName: parsed.sheetName ?? sheetName,
      rowStart: parsed.row,
      rowEnd: parsed.row,
      colStart: parsed.col,
      colEnd: parsed.col,
    }
  }

  const start = (ref?.kind === 'range' ? ref.start : undefined) ?? (value?.kind === 'range' ? value.start : undefined)
  const end = (ref?.kind === 'range' ? ref.end : undefined) ?? (value?.kind === 'range' ? value.end : undefined)
  const refKind =
    ref?.kind === 'row' ? 'rows' : ref?.kind === 'col' ? 'cols' : (ref?.refKind ?? (value?.kind === 'range' ? value.refKind : undefined))
  const address = ref?.kind === 'row' || ref?.kind === 'col' ? ref.address : undefined
  const rangeStart = start ?? address
  const rangeEnd = end ?? address
  if (!rangeStart || !rangeEnd || !refKind) {
    return undefined
  }

  const parsed = parseRangeAddress(`${rangeStart}:${rangeEnd}`, sheetName)
  if (parsed.kind !== refKind) {
    return undefined
  }
  if (parsed.kind === 'cells') {
    return {
      sheetName: parsed.sheetName ?? sheetName,
      rowStart: parsed.start.row,
      rowEnd: parsed.end.row,
      colStart: parsed.start.col,
      colEnd: parsed.end.col,
    }
  }
  if (parsed.kind === 'rows') {
    return {
      sheetName: parsed.sheetName ?? sheetName,
      rowStart: parsed.start.row,
      rowEnd: parsed.end.row,
      colStart: 0,
      colEnd: MAX_COLS - 1,
    }
  }
  return {
    sheetName: parsed.sheetName ?? sheetName,
    rowStart: 0,
    rowEnd: MAX_ROWS - 1,
    colStart: parsed.start.col,
    colEnd: parsed.end.col,
  }
}

function stackCellRange(
  bounds: Pick<ReferenceBounds, 'sheetName' | 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>,
  context: EvaluationContext,
): StackValue {
  const start = formatAddress(bounds.rowStart, bounds.colStart)
  const end = formatAddress(bounds.rowEnd, bounds.colEnd)
  const values =
    bounds.rowStart === bounds.rowEnd && bounds.colStart === bounds.colEnd
      ? [context.resolveCell(bounds.sheetName, start)]
      : context.resolveRange(bounds.sheetName, start, end, 'cells')
  const blankReference = bounds.rowStart === bounds.rowEnd && bounds.colStart === bounds.colEnd && values[0]?.tag === ValueTag.Empty
  return {
    kind: 'range',
    values,
    refKind: 'cells',
    rows: bounds.rowEnd - bounds.rowStart + 1,
    cols: bounds.colEnd - bounds.colStart + 1,
    sheetName: bounds.sheetName,
    start,
    end,
    ...(blankReference ? { blankReference: true } : {}),
  }
}

function wholeAxisLookupRange(
  reference: WholeAxisReference,
  context: EvaluationContext,
  deps: WorkbookSpecialCallDeps,
): RangeBuiltinArgument | CellValue {
  const values = context.resolveRange(reference.sheetName, reference.start, reference.end, reference.refKind)
  if (reference.parsed.kind === 'rows') {
    const rowCount = reference.parsed.end.row - reference.parsed.start.row + 1
    if (rowCount !== 1) {
      return deps.error(ErrorCode.Value)
    }
    return {
      kind: 'range',
      values,
      refKind: 'cells',
      rows: 1,
      cols: values.length,
      sheetName: reference.sheetName,
      start: formatAddress(reference.parsed.start.row, 0),
      end: formatAddress(reference.parsed.start.row, Math.max(values.length - 1, 0)),
    }
  }

  const colCount = reference.parsed.end.col - reference.parsed.start.col + 1
  if (colCount !== 1) {
    return deps.error(ErrorCode.Value)
  }
  return {
    kind: 'range',
    values,
    refKind: 'cells',
    rows: values.length,
    cols: 1,
    sheetName: reference.sheetName,
    start: formatAddress(0, reference.parsed.start.col),
    end: formatAddress(Math.max(values.length - 1, 0), reference.parsed.start.col),
  }
}

function wholeAxisTableRange(
  reference: WholeAxisReference,
  context: EvaluationContext,
  deps: WorkbookSpecialCallDeps,
): RangeBuiltinArgument | CellValue {
  const values = context.resolveRange(reference.sheetName, reference.start, reference.end, reference.refKind)
  if (reference.parsed.kind === 'cols') {
    const cols = reference.parsed.end.col - reference.parsed.start.col + 1
    if (cols <= 0 || values.length % cols !== 0) {
      return deps.error(ErrorCode.Value)
    }
    const rows = values.length / cols
    return {
      kind: 'range',
      values,
      refKind: 'cells',
      rows,
      cols,
      sheetName: reference.sheetName,
      start: formatAddress(0, reference.parsed.start.col),
      end: formatAddress(Math.max(rows - 1, 0), reference.parsed.end.col),
    }
  }

  const rows = reference.parsed.end.row - reference.parsed.start.row + 1
  if (rows <= 0 || values.length % rows !== 0) {
    return deps.error(ErrorCode.Value)
  }
  const cols = values.length / rows
  return {
    kind: 'range',
    values,
    refKind: 'cells',
    rows,
    cols,
    sheetName: reference.sheetName,
    start: formatAddress(reference.parsed.start.row, 0),
    end: formatAddress(reference.parsed.end.row, Math.max(cols - 1, 0)),
  }
}

function wholeAxisResidentShape(
  value: StackValue | undefined,
  ref: ReferenceOperand | undefined,
  context: EvaluationContext,
  deps: WorkbookSpecialCallDeps,
): { rows: number; cols: number } | undefined {
  const reference = wholeAxisReferenceFromArg(value, ref, context, deps)
  if (!reference) {
    return undefined
  }
  const values =
    value?.kind === 'range' ? value.values : context.resolveRange(reference.sheetName, reference.start, reference.end, reference.refKind)

  if (reference.parsed.kind === 'cols') {
    const cols = reference.parsed.end.col - reference.parsed.start.col + 1
    return {
      rows: cols <= 0 ? 0 : Math.floor(values.length / cols),
      cols,
    }
  }

  const rows = reference.parsed.end.row - reference.parsed.start.row + 1
  return {
    rows,
    cols: rows <= 0 ? 0 : Math.floor(values.length / rows),
  }
}

function evaluateWholeAxisShape(
  callee: 'ROWS' | 'COLUMNS',
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  const shape = wholeAxisResidentShape(rawArgs[0], argRefs[0], context, deps)
  if (!shape) {
    return undefined
  }
  if (rawArgs.length !== 1) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  return deps.stackScalar({
    tag: ValueTag.Number,
    value: callee === 'ROWS' ? shape.rows : shape.cols,
  })
}

function evaluateWholeAxisTableLookup(
  callee: 'VLOOKUP' | 'HLOOKUP',
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  const reference = wholeAxisReferenceFromArg(rawArgs[1], argRefs[1], context, deps)
  if (!reference) {
    return undefined
  }
  if (rawArgs.length < 3 || rawArgs.length > 4) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const lookupBuiltin = context.resolveLookupBuiltin?.(callee) ?? getLookupBuiltin(callee)
  if (!lookupBuiltin) {
    return undefined
  }

  const lookupValue = scalarLookupArgument(rawArgs[0], deps)
  if (lookupValue.kind === 'error') {
    return deps.stackScalar(lookupValue.value)
  }
  if (lookupValue.value === undefined) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const indexValue = scalarLookupArgument(rawArgs[2], deps)
  if (indexValue.kind === 'error') {
    return deps.stackScalar(indexValue.value)
  }

  const rangeLookupValue = scalarLookupArgument(rawArgs[3], deps)
  if (rangeLookupValue.kind === 'error') {
    return deps.stackScalar(rangeLookupValue.value)
  }

  const tableRange = wholeAxisTableRange(reference, context, deps)
  if ('tag' in tableRange) {
    return deps.stackScalar(tableRange)
  }

  const result = lookupBuiltin(lookupValue.value, tableRange, indexValue.value, rangeLookupValue.value)
  return isArrayValue(result) ? result : deps.stackScalar(result)
}

function evaluateWholeAxisMatch(
  callee: 'MATCH' | 'XMATCH',
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  const reference = wholeAxisReferenceFromArg(rawArgs[1], argRefs[1], context, deps)
  if (!reference) {
    return undefined
  }
  if (rawArgs.length < 2 || (callee === 'MATCH' ? rawArgs.length > 3 : rawArgs.length > 4)) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  const lookupValue = deps.isSingleCellValue(rawArgs[0]!)
  if (!lookupValue) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  const lookupRange = wholeAxisLookupRange(reference, context, deps)
  if ('tag' in lookupRange) {
    return deps.stackScalar(lookupRange)
  }

  const lookupBuiltin = context.resolveLookupBuiltin?.(callee) ?? getLookupBuiltin(callee)
  if (!lookupBuiltin) {
    return undefined
  }

  const firstOptional = scalarLookupArgument(rawArgs[2], deps)
  if (firstOptional.kind === 'error') {
    return deps.stackScalar(firstOptional.value)
  }
  if (callee === 'MATCH') {
    const result = lookupBuiltin(lookupValue, lookupRange, firstOptional.value)
    return isArrayValue(result) ? result : deps.stackScalar(result)
  }

  const secondOptional = scalarLookupArgument(rawArgs[3], deps)
  if (secondOptional.kind === 'error') {
    return deps.stackScalar(secondOptional.value)
  }
  const result = lookupBuiltin(lookupValue, lookupRange, firstOptional.value, secondOptional.value)
  return isArrayValue(result) ? result : deps.stackScalar(result)
}

function evaluateReferenceIndex(
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  const bounds = referenceBoundsFromArg(rawArgs[0], argRefs[0], context)
  if (!bounds) {
    return undefined
  }
  if (rawArgs.length < 1 || rawArgs.length > 3) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const rowArg = integerIndexArgument(rawArgs[1], deps)
  if (rowArg.kind === 'error') {
    return deps.stackScalar(rowArg.value)
  }
  const colArg = integerIndexArgument(rawArgs[2], deps)
  if (colArg.kind === 'error') {
    return deps.stackScalar(colArg.value)
  }

  const rawRowNum = rowArg.kind === 'ok' ? rowArg.value : undefined
  const rawColNum = colArg.kind === 'ok' ? colArg.value : undefined
  if (rawRowNum === undefined && rawColNum === undefined) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const rowCount = bounds.rowEnd - bounds.rowStart + 1
  const colCount = bounds.colEnd - bounds.colStart + 1
  let rowNum = rawRowNum ?? 0
  let colNum = rawColNum ?? 0
  if (rowCount === 1 && rawColNum === undefined && rawRowNum !== undefined && rawRowNum !== 0) {
    rowNum = 1
    colNum = rawRowNum
  }
  if (colCount === 1 && rawColNum === undefined && rawRowNum !== undefined && rawRowNum !== 0) {
    colNum = 1
  }
  if (rowNum < 0 || colNum < 0 || rowNum > rowCount || colNum > colCount) {
    return deps.stackScalar(deps.error(ErrorCode.Ref))
  }

  if (rowNum === 0 && colNum === 0) {
    return stackCellRange(bounds, context)
  }
  if (rowNum === 0) {
    return stackCellRange(
      {
        sheetName: bounds.sheetName,
        rowStart: bounds.rowStart,
        rowEnd: bounds.rowEnd,
        colStart: bounds.colStart + colNum - 1,
        colEnd: bounds.colStart + colNum - 1,
      },
      context,
    )
  }
  if (colNum === 0) {
    return stackCellRange(
      {
        sheetName: bounds.sheetName,
        rowStart: bounds.rowStart + rowNum - 1,
        rowEnd: bounds.rowStart + rowNum - 1,
        colStart: bounds.colStart,
        colEnd: bounds.colEnd,
      },
      context,
    )
  }

  return stackCellRange(
    {
      sheetName: bounds.sheetName,
      rowStart: bounds.rowStart + rowNum - 1,
      rowEnd: bounds.rowStart + rowNum - 1,
      colStart: bounds.colStart + colNum - 1,
      colEnd: bounds.colStart + colNum - 1,
    },
    context,
  )
}

function evaluateReferenceOffset(
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  const bounds = referenceBoundsFromArg(rawArgs[0], argRefs[0], context)
  if (!bounds) {
    return undefined
  }
  if (rawArgs.length < 3 || rawArgs.length > 5) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const rowArg = integerIndexArgument(rawArgs[1], deps)
  if (rowArg.kind === 'error') {
    return deps.stackScalar(rowArg.value)
  }
  const colArg = integerIndexArgument(rawArgs[2], deps)
  if (colArg.kind === 'error') {
    return deps.stackScalar(colArg.value)
  }
  if (rowArg.kind !== 'ok' || colArg.kind !== 'ok') {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const referenceRows = bounds.rowEnd - bounds.rowStart + 1
  const referenceCols = bounds.colEnd - bounds.colStart + 1
  const heightArg = rawArgs.length >= 4 ? integerIndexArgument(rawArgs[3], deps) : ({ kind: 'ok', value: referenceRows } as const)
  if (heightArg.kind === 'error') {
    return deps.stackScalar(heightArg.value)
  }
  const widthArg = rawArgs.length >= 5 ? integerIndexArgument(rawArgs[4], deps) : ({ kind: 'ok', value: referenceCols } as const)
  if (widthArg.kind === 'error') {
    return deps.stackScalar(widthArg.value)
  }
  if (heightArg.kind !== 'ok' || widthArg.kind !== 'ok') {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  if (heightArg.value < 1 || widthArg.value < 1) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }

  const rowStart = bounds.rowStart + rowArg.value
  const colStart = bounds.colStart + colArg.value
  const rowEnd = rowStart + heightArg.value - 1
  const colEnd = colStart + widthArg.value - 1
  if (rowStart < 0 || colStart < 0 || rowEnd >= MAX_ROWS || colEnd >= MAX_COLS || rowEnd < rowStart || colEnd < colStart) {
    return deps.stackScalar(deps.error(ErrorCode.Ref))
  }

  return stackCellRange(
    {
      sheetName: bounds.sheetName,
      rowStart,
      rowEnd,
      colStart,
      colEnd,
    },
    context,
  )
}

export function evaluateWorkbookSpecialCall(
  callee: string,
  rawArgs: StackValue[],
  context: EvaluationContext,
  argRefs: readonly (ReferenceOperand | undefined)[],
  deps: WorkbookSpecialCallDeps,
): StackValue | undefined {
  switch (callee) {
    case 'VLOOKUP':
    case 'HLOOKUP':
      return evaluateWholeAxisTableLookup(callee, rawArgs, context, argRefs, deps)
    case 'MATCH':
    case 'XMATCH':
      return evaluateWholeAxisMatch(callee, rawArgs, context, argRefs, deps)
    case 'INDEX':
      return evaluateReferenceIndex(rawArgs, context, argRefs, deps)
    case 'OFFSET':
      return evaluateReferenceOffset(rawArgs, context, argRefs, deps)
    case 'ROWS':
    case 'COLUMNS':
      return evaluateWholeAxisShape(callee, rawArgs, context, argRefs, deps)
    case 'SUBTOTAL': {
      const functionNum = deps.scalarIntegerArgument(rawArgs[0])
      if (functionNum === undefined) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const ignoreHiddenRows = functionNum > 100 && context.isRowHidden !== undefined
      if (!ignoreHiddenRows && !context.resolveFormula) {
        return undefined
      }
      const subtotal = getBuiltin('SUBTOTAL')
      if (!subtotal) {
        return undefined
      }
      const candidates = rawArgs
        .slice(1)
        .flatMap((value, index) => collectAggregateCandidates(value, argRefs[index + 1], context, ignoreHiddenRows))
      const values = filterNestedRollupCandidates(candidates, context, nestedSubtotalCallees).map((candidate) => candidate.value)
      const result = subtotal({ tag: ValueTag.Number, value: functionNum }, ...values)
      return isArrayValue(result) ? result : deps.stackScalar(result)
    }
    case 'AGGREGATE': {
      const functionNum = deps.scalarIntegerArgument(rawArgs[0])
      const option = deps.scalarIntegerArgument(rawArgs[1])
      if (functionNum === undefined || option === undefined || option < 0 || option > 7) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const aggregate = getBuiltin('AGGREGATE')
      if (!aggregate) {
        return undefined
      }
      const candidates = rawArgs
        .slice(2)
        .flatMap((value, index) => collectAggregateCandidates(value, argRefs[index + 2], context, aggregateOptionIgnoresHiddenRows(option)))
      const nestedFilteredCandidates = aggregateOptionIgnoresNestedRollups(option)
        ? filterNestedRollupCandidates(candidates, context, nestedAggregateCallees)
        : candidates
      let values = nestedFilteredCandidates.map((candidate) => candidate.value)
      if (aggregateOptionIgnoresErrors(option)) {
        values = values.filter((value) => value.tag !== ValueTag.Error)
      } else {
        const error = firstErrorValue(values)
        if (error) {
          return deps.stackScalar(error)
        }
      }
      const result = aggregate({ tag: ValueTag.Number, value: functionNum }, { tag: ValueTag.Number, value: option }, ...values)
      return isArrayValue(result) ? result : deps.stackScalar(result)
    }
    case 'GETPIVOTDATA': {
      if (rawArgs.length < 2 || (rawArgs.length - 2) % 2 !== 0) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const dataFieldValue = deps.isSingleCellValue(rawArgs[0]!)
      const address = deps.referenceTopLeftAddress(argRefs[1])
      const sheetName = deps.referenceSheetName(argRefs[1], context)
      if (!dataFieldValue) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      if (!address || !sheetName) {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }
      const filters: Array<{ field: string; item: CellValue }> = []
      for (let index = 2; index < rawArgs.length; index += 2) {
        const fieldValue = deps.isSingleCellValue(rawArgs[index]!)
        const itemValue = deps.isSingleCellValue(rawArgs[index + 1]!)
        if (!fieldValue || !itemValue) {
          return deps.stackScalar(deps.error(ErrorCode.Value))
        }
        filters.push({ field: deps.toStringValue(fieldValue), item: itemValue })
      }
      return deps.stackScalar(
        context.resolvePivotData?.({
          dataField: deps.toStringValue(dataFieldValue),
          sheetName,
          address,
          filters,
        }) ?? deps.error(ErrorCode.Ref),
      )
    }
    case 'GROUPBY': {
      if (rawArgs.length < 3 || rawArgs.length > 8) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const rowFields = deps.matrixFromStackValue(rawArgs[0]!)
      const values = deps.matrixFromStackValue(rawArgs[1]!)
      if (!rowFields || !values) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const sortOrder =
        deps.vectorIntegerArgument(rawArgs[5]) ?? (rawArgs[5] ? [deps.scalarIntegerArgument(rawArgs[5]) ?? Number.NaN] : undefined)
      const fieldHeadersMode = deps.scalarIntegerArgument(rawArgs[3])
      const totalDepth = deps.scalarIntegerArgument(rawArgs[4])
      const filterArray = rawArgs[6] ? deps.matrixFromStackValue(rawArgs[6]) : undefined
      const fieldRelationship = deps.scalarIntegerArgument(rawArgs[7])
      const result = evaluateGroupBy(rowFields, values, {
        aggregate: (subset: readonly CellValue[], totalSet?: readonly CellValue[]) =>
          deps.aggregateRangeSubset(rawArgs[2]!, subset, context, totalSet),
        ...(fieldHeadersMode !== undefined ? { fieldHeadersMode } : {}),
        ...(totalDepth !== undefined ? { totalDepth } : {}),
        ...(sortOrder?.every(Number.isFinite) ? { sortOrder } : {}),
        ...(filterArray !== undefined ? { filterArray } : {}),
        ...(fieldRelationship !== undefined ? { fieldRelationship } : {}),
      })
      return isArrayValue(result) ? result : deps.stackScalar(result)
    }
    case 'PIVOTBY': {
      if (rawArgs.length < 4 || rawArgs.length > 11) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const rowFields = deps.matrixFromStackValue(rawArgs[0]!)
      const colFields = deps.matrixFromStackValue(rawArgs[1]!)
      const values = deps.matrixFromStackValue(rawArgs[2]!)
      if (!rowFields || !colFields || !values) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const rowSortOrder =
        deps.vectorIntegerArgument(rawArgs[6]) ?? (rawArgs[6] ? [deps.scalarIntegerArgument(rawArgs[6]) ?? Number.NaN] : undefined)
      const colSortOrder =
        deps.vectorIntegerArgument(rawArgs[8]) ?? (rawArgs[8] ? [deps.scalarIntegerArgument(rawArgs[8]) ?? Number.NaN] : undefined)
      const fieldHeadersMode = deps.scalarIntegerArgument(rawArgs[4])
      const rowTotalDepth = deps.scalarIntegerArgument(rawArgs[5])
      const colTotalDepth = deps.scalarIntegerArgument(rawArgs[7])
      const filterArray = rawArgs[9] ? deps.matrixFromStackValue(rawArgs[9]) : undefined
      const relativeTo = deps.scalarIntegerArgument(rawArgs[10])
      const result = evaluatePivotBy(rowFields, colFields, values, {
        aggregate: (subset: readonly CellValue[], totalSet?: readonly CellValue[]) =>
          deps.aggregateRangeSubset(rawArgs[3]!, subset, context, totalSet),
        ...(fieldHeadersMode !== undefined ? { fieldHeadersMode } : {}),
        ...(rowTotalDepth !== undefined ? { rowTotalDepth } : {}),
        ...(rowSortOrder?.every(Number.isFinite) ? { rowSortOrder } : {}),
        ...(colTotalDepth !== undefined ? { colTotalDepth } : {}),
        ...(colSortOrder?.every(Number.isFinite) ? { colSortOrder } : {}),
        ...(filterArray !== undefined ? { filterArray } : {}),
        ...(relativeTo !== undefined ? { relativeTo } : {}),
      })
      return isArrayValue(result) ? result : deps.stackScalar(result)
    }
    case 'MULTIPLE.OPERATIONS': {
      if (rawArgs.length !== 3 && rawArgs.length !== 5) {
        return deps.stackScalar(deps.error(ErrorCode.Value))
      }
      const formulaAddress = deps.referenceTopLeftAddress(argRefs[0])
      const formulaSheetName = deps.referenceSheetName(argRefs[0], context)
      const rowCellAddress = deps.referenceTopLeftAddress(argRefs[1])
      const rowCellSheetName = deps.referenceSheetName(argRefs[1], context)
      const rowReplacementAddress = deps.referenceTopLeftAddress(argRefs[2])
      const rowReplacementSheetName = deps.referenceSheetName(argRefs[2], context)
      if (
        !formulaAddress ||
        !formulaSheetName ||
        !rowCellAddress ||
        !rowCellSheetName ||
        !rowReplacementAddress ||
        !rowReplacementSheetName
      ) {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }
      const columnCellAddress = rawArgs.length === 5 ? deps.referenceTopLeftAddress(argRefs[3]) : undefined
      const columnCellSheetName = rawArgs.length === 5 ? deps.referenceSheetName(argRefs[3], context) : undefined
      const columnReplacementAddress = rawArgs.length === 5 ? deps.referenceTopLeftAddress(argRefs[4]) : undefined
      const columnReplacementSheetName = rawArgs.length === 5 ? deps.referenceSheetName(argRefs[4], context) : undefined
      if (
        rawArgs.length === 5 &&
        (!columnCellAddress || !columnCellSheetName || !columnReplacementAddress || !columnReplacementSheetName)
      ) {
        return deps.stackScalar(deps.error(ErrorCode.Ref))
      }
      return deps.stackScalar(
        context.resolveMultipleOperations?.({
          formulaSheetName,
          formulaAddress,
          rowCellSheetName,
          rowCellAddress,
          rowReplacementSheetName,
          rowReplacementAddress,
          ...(columnCellSheetName ? { columnCellSheetName } : {}),
          ...(columnCellAddress ? { columnCellAddress } : {}),
          ...(columnReplacementSheetName ? { columnReplacementSheetName } : {}),
          ...(columnReplacementAddress ? { columnReplacementAddress } : {}),
        }) ?? deps.error(ErrorCode.Ref),
      )
    }
    case 'INDIRECT':
      return evaluateIndirectWorkbookSpecialCall(rawArgs, context, deps)
    default:
      return undefined
  }
}
