import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { parseCellAddress, parseRangeAddress } from './addressing.js'
import type { EvaluationContext, ReferenceOperand, StackValue } from './js-evaluator.js'

export interface WorkbookReferenceCallDeps {
  error: (code: ErrorCode) => CellValue
  stackScalar: (value: CellValue, blankReference?: boolean) => StackValue
  toStringValue: (value: CellValue) => string
  referenceTopLeftAddress: (ref: ReferenceOperand | undefined) => string | undefined
  referenceSheetName: (ref: ReferenceOperand | undefined, context: EvaluationContext) => string | undefined
  coerceScalarTextArgument: (value: StackValue | undefined) => string | CellValue
  coerceOptionalBooleanArgument: (value: StackValue | undefined, fallback: boolean) => boolean | CellValue
  isCellValueError: (value: number | boolean | string | CellValue) => value is CellValue
}

export function stackReferenceOperand(ref: ReferenceOperand, context: EvaluationContext, deps: WorkbookReferenceCallDeps): StackValue {
  const sheetName = deps.referenceSheetName(ref, context) ?? context.sheetName
  if (ref.kind === 'cell') {
    const address = ref.address ?? deps.referenceTopLeftAddress(ref)
    if (!address) {
      return deps.stackScalar(deps.error(ErrorCode.Ref))
    }
    const value = context.resolveCell(sheetName, address)
    return deps.stackScalar(value, value.tag === ValueTag.Empty)
  }

  const start = ref.start
  const end = ref.end
  const refKind = ref.refKind ?? (ref.kind === 'row' ? 'rows' : ref.kind === 'col' ? 'cols' : 'cells')
  if (!start || !end) {
    return deps.stackScalar(deps.error(ErrorCode.Ref))
  }

  const values = context.resolveRange(sheetName, start, end, refKind)
  let rows = values.length
  let cols = 1
  if (refKind === 'cells') {
    try {
      const parsed = parseRangeAddress(`${start}:${end}`, sheetName)
      if (parsed.kind === 'cells') {
        rows = parsed.end.row - parsed.start.row + 1
        cols = parsed.end.col - parsed.start.col + 1
      }
    } catch {
      rows = values.length
      cols = 1
    }
  }
  return {
    kind: 'range',
    values,
    refKind,
    rows,
    cols,
    sheetName,
    start,
    end,
  }
}

export function evaluateIndirectWorkbookSpecialCall(
  rawArgs: StackValue[],
  context: EvaluationContext,
  deps: WorkbookReferenceCallDeps,
): StackValue {
  if (rawArgs.length < 1 || rawArgs.length > 2) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  const refText = deps.coerceScalarTextArgument(rawArgs[0])
  if (deps.isCellValueError(refText)) {
    return deps.stackScalar(refText)
  }
  const a1Mode = deps.coerceOptionalBooleanArgument(rawArgs[1], true)
  if (deps.isCellValueError(a1Mode)) {
    return deps.stackScalar(a1Mode)
  }
  if (!a1Mode) {
    return deps.stackScalar(deps.error(ErrorCode.Value))
  }
  const normalizedRefText = refText.trim()
  if (normalizedRefText === '') {
    return deps.stackScalar(deps.error(ErrorCode.Ref))
  }

  try {
    const cell = parseCellAddress(normalizedRefText, context.sheetName)
    const value = context.resolveCell(cell.sheetName ?? context.sheetName, cell.text)
    return deps.stackScalar(value, value.tag === ValueTag.Empty)
  } catch {
    // Try range, defined-name reference, and scalar name fallbacks below.
  }

  try {
    const range = parseRangeAddress(normalizedRefText, context.sheetName)
    if (range.kind !== 'cells') {
      return deps.stackScalar(deps.error(ErrorCode.Ref))
    }
    const targetSheetName = range.sheetName ?? context.sheetName
    const values = context.resolveRange(targetSheetName, range.start.text, range.end.text, 'cells')
    return {
      kind: 'range',
      values,
      refKind: 'cells',
      rows: range.end.row - range.start.row + 1,
      cols: range.end.col - range.start.col + 1,
      sheetName: targetSheetName,
      start: range.start.text,
      end: range.end.text,
      ...(range.start.row === range.end.row && range.start.col === range.end.col && values[0]?.tag === ValueTag.Empty
        ? { blankReference: true }
        : {}),
    }
  } catch {
    // Try defined-name reference and scalar name fallbacks below.
  }

  const resolvedReference = context.resolveNameReference?.(normalizedRefText)
  if (resolvedReference) {
    return stackReferenceOperand(resolvedReference, context, deps)
  }

  const resolvedName = context.resolveName?.(normalizedRefText)
  return deps.stackScalar(resolvedName ?? deps.error(ErrorCode.Ref))
}
