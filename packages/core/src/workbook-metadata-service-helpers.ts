import { Effect } from 'effect'
import { ErrorCode, type CellRangeRef, type WorkbookDefinedNameValueSnapshot } from '@bilig/protocol'
import { parseCellAddress, parseFormula, renameFormulaSheetReferences, serializeFormula, type FormulaNode } from '@bilig/formula'
import { cloneDataValidationRecord, cloneDefinedNameRecord } from './workbook-metadata-records.js'
import type {
  WorkbookDataValidationRecord,
  WorkbookDefinedNameRecord,
  WorkbookFilterRecord,
  WorkbookMergeRangeRecord,
} from './workbook-metadata-types.js'
import { canonicalWorkbookRangeRef } from './workbook-range-records.js'
import { WorkbookMetadataError } from './workbook-metadata-service-contract.js'

export function renameDataValidationSourceSheet(
  record: WorkbookDataValidationRecord,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDataValidationRecord {
  const cloned = cloneDataValidationRecord(record)
  if (cloned.rule.kind !== 'list' || !cloned.rule.source) {
    return cloned
  }
  switch (cloned.rule.source.kind) {
    case 'cell-ref':
    case 'range-ref':
      if (cloned.rule.source.sheetName === oldSheetName) {
        cloned.rule.source.sheetName = newSheetName
      }
      return cloned
    case 'named-range':
    case 'structured-ref':
      return cloned
    case 'formula':
      cloned.rule.source.formula = renameValidationFormulaSheetReferences(cloned.rule.source.formula, oldSheetName, newSheetName)
      return cloned
  }
  return cloned
}

function renameValidationFormulaSheetReferences(formula: string, oldSheetName: string, newSheetName: string): string {
  const hasFormulaPrefix = formula.startsWith('=')
  const source = hasFormulaPrefix ? formula.slice(1) : formula
  try {
    const renamed = renameFormulaSheetReferences(source, oldSheetName, newSheetName)
    return hasFormulaPrefix ? `=${renamed}` : renamed
  } catch {
    return formula
  }
}

export function rewriteDefinedNameForSheetDeletion(record: WorkbookDefinedNameRecord, deletedSheetName: string): WorkbookDefinedNameRecord {
  const cloned = cloneDefinedNameRecord(record)
  cloned.value = rewriteDefinedNameValueForSheetDeletion(cloned.value, deletedSheetName)
  return cloned
}

function rewriteDefinedNameValueForSheetDeletion(
  value: WorkbookDefinedNameValueSnapshot,
  deletedSheetName: string,
): WorkbookDefinedNameValueSnapshot {
  if (typeof value === 'string') {
    return value.startsWith('=') ? rewriteDefinedNameFormulaForSheetDeletion(value, deletedSheetName) : value
  }
  if (value === null || typeof value !== 'object') {
    return value
  }
  switch (value.kind) {
    case 'cell-ref':
      return value.sheetName === deletedSheetName ? { kind: 'formula', formula: '=#REF!' } : value
    case 'range-ref':
      return value.sheetName === deletedSheetName ? { kind: 'formula', formula: '=#REF!' } : value
    case 'formula':
      return {
        ...value,
        formula: rewriteDefinedNameFormulaForSheetDeletion(value.formula, deletedSheetName),
      }
    case 'scalar':
    case 'structured-ref':
      return value
  }
}

function rewriteDefinedNameFormulaForSheetDeletion(formula: string, deletedSheetName: string): string {
  const hasFormulaPrefix = formula.startsWith('=')
  const source = hasFormulaPrefix ? formula.slice(1) : formula
  try {
    const rewritten = serializeFormula(rewriteFormulaNodeForSheetDeletion(parseFormula(source), deletedSheetName))
    return hasFormulaPrefix ? `=${rewritten}` : rewritten
  } catch {
    return formula
  }
}

function rewriteFormulaNodeForSheetDeletion(node: FormulaNode, deletedSheetName: string): FormulaNode {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'StructuredRef':
      return node
    case 'NameRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
      return node.sheetName === deletedSheetName ? refErrorNode() : node
    case 'RangeRef':
      return node.sheetName === deletedSheetName || node.sheetEndName === deletedSheetName ? refErrorNode() : node
    case 'ArrayConstant':
      return { ...node, rows: node.rows.map((row) => row.map((entry) => rewriteFormulaNodeForSheetDeletion(entry, deletedSheetName))) }
    case 'UnaryExpr':
      return {
        ...node,
        argument: rewriteFormulaNodeForSheetDeletion(node.argument, deletedSheetName),
      }
    case 'BinaryExpr':
      return {
        ...node,
        left: rewriteFormulaNodeForSheetDeletion(node.left, deletedSheetName),
        right: rewriteFormulaNodeForSheetDeletion(node.right, deletedSheetName),
      }
    case 'CallExpr':
      return {
        ...node,
        args: node.args.map((arg) => rewriteFormulaNodeForSheetDeletion(arg, deletedSheetName)),
      }
    case 'InvokeExpr':
      return {
        ...node,
        callee: rewriteFormulaNodeForSheetDeletion(node.callee, deletedSheetName),
        args: node.args.map((arg) => rewriteFormulaNodeForSheetDeletion(arg, deletedSheetName)),
      }
  }
}

function refErrorNode(): FormulaNode {
  return { kind: 'ErrorLiteral', code: ErrorCode.Ref }
}

export function assertMergeRangesDoNotOverlap(ranges: readonly WorkbookMergeRangeRecord[]): void {
  const normalized = ranges.map(normalizeMergeRangeForOverlap).toSorted((left, right) => left.startRow - right.startRow)
  const active: NormalizedMergeRangeRecord[] = []
  for (const range of normalized) {
    for (let index = active.length - 1; index >= 0; index -= 1) {
      if (active[index]!.endRow < range.startRow) {
        active.splice(index, 1)
      }
    }
    if (active.some((entry) => mergeRangesOverlap(entry, range))) {
      throw new Error('Merged ranges cannot overlap')
    }
    active.push(range)
  }
}

export function canonicalWorkbookFilterRange(range: WorkbookFilterRecord['range']): WorkbookFilterRecord['range'] {
  const normalized = canonicalWorkbookRangeRef(range)
  const criteria = range.criteria?.length ? structuredClone(range.criteria) : undefined
  return criteria ? { ...normalized, criteria } : normalized
}

export function canonicalWorkbookRangeOnSheet(sheetName: string, range: CellRangeRef): CellRangeRef {
  return canonicalWorkbookRangeRef({ ...range, sheetName })
}

export function canonicalWorkbookFilterRangeOnSheet(
  sheetName: string,
  range: WorkbookFilterRecord['range'],
): WorkbookFilterRecord['range'] {
  return {
    ...canonicalWorkbookFilterRange({ ...range, sheetName }),
    sheetName,
  }
}

export function metadataEffect<Success>(message: string, run: () => Success): Effect.Effect<Success, WorkbookMetadataError> {
  return Effect.try({
    try: run,
    catch: (cause) =>
      new WorkbookMetadataError({
        message: metadataErrorMessage(message, cause),
        cause,
      }),
  })
}

export function normalizeMetadataKey(key: string): string {
  const trimmed = key.trim()
  if (trimmed.length === 0) {
    throw new Error('Workbook metadata keys must be non-empty')
  }
  return trimmed
}

interface NormalizedMergeRangeRecord {
  readonly record: WorkbookMergeRangeRecord
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

function normalizeMergeRangeForOverlap(record: WorkbookMergeRangeRecord): NormalizedMergeRangeRecord {
  const start = parseCellAddress(record.startAddress, record.sheetName)
  const end = parseCellAddress(record.endAddress, record.sheetName)
  return {
    record,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

function mergeRangesOverlap(left: NormalizedMergeRangeRecord, right: NormalizedMergeRangeRecord): boolean {
  return !(left.endRow < right.startRow || right.endRow < left.startRow || left.endCol < right.startCol || right.endCol < left.startCol)
}

function metadataErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}
