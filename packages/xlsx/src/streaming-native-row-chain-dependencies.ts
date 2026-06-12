import { parseFormula, type FormulaNode } from '@bilig/formula'
import { ValueTag, type CellValue } from '@bilig/protocol'

import { decodeCellAddress } from './address.js'
import type { NativeFormulaCell, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'

export interface CellRangePlan {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
  readonly cellCount: number
}

const maxExpandedFormulaDependencyRowsPerSheet = 20_000
const maxNativeRangeAggregateCellCount = 50_000
const directAggregateOpSum = 1
const directAggregateOpAverage = 2
const directAggregateOpCounta = 6

export function expandStreamingNativeFormulaDependencyRows(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly targetRowsForSheet?: (sheetName: string) => Set<number> | undefined
}): ReadonlySet<string> {
  const changedSheets = new Set<string>()
  for (const scan of args.sheetScans.values()) {
    let expandedRows = 0
    for (const cell of scan.formulaCells) {
      for (const dependency of collectScalarDependencyRows(scan, cell, args.resolveFormulaSource)) {
        const targetRows =
          args.targetRowsForSheet?.(dependency.sheetName) ?? (dependency.sheetName === scan.sheetName ? scan.targetRows : undefined)
        if (!targetRows || targetRows.has(dependency.row)) {
          continue
        }
        targetRows.add(dependency.row)
        changedSheets.add(dependency.sheetName)
        expandedRows += 1
        if (expandedRows >= maxExpandedFormulaDependencyRowsPerSheet) {
          break
        }
      }
      if (expandedRows >= maxExpandedFormulaDependencyRowsPerSheet) {
        break
      }
      for (const dependency of collectVlookupTableDependencyRows(scan, cell, args.resolveFormulaSource)) {
        const targetRows =
          args.targetRowsForSheet?.(dependency.sheetName) ?? (dependency.sheetName === scan.sheetName ? scan.targetRows : undefined)
        if (!targetRows || targetRows.has(dependency.row)) {
          continue
        }
        targetRows.add(dependency.row)
        changedSheets.add(dependency.sheetName)
        expandedRows += 1
        if (expandedRows >= maxExpandedFormulaDependencyRowsPerSheet) {
          break
        }
      }
      if (expandedRows >= maxExpandedFormulaDependencyRowsPerSheet) {
        break
      }
      const plan = tryCompileNativeRangeAggregate(scan, cell, args.resolveFormulaSource)
      if (!plan) {
        continue
      }
      for (let row = plan.range.startRow; row <= plan.range.endRow; row += 1) {
        if (scan.targetRows.has(row)) {
          continue
        }
        scan.targetRows.add(row)
        changedSheets.add(scan.sheetName)
        expandedRows += 1
        if (expandedRows >= maxExpandedFormulaDependencyRowsPerSheet) {
          break
        }
      }
      if (expandedRows >= maxExpandedFormulaDependencyRowsPerSheet) {
        break
      }
    }
  }
  return changedSheets
}

function collectVlookupTableDependencyRows(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
): readonly ScalarDependencyRow[] {
  let node: FormulaNode
  try {
    node = parseFormula(resolveFormulaSource(scan, cell))
  } catch {
    return []
  }
  if (node.kind !== 'CallExpr' || node.callee.toLocaleUpperCase('en-US') !== 'VLOOKUP') {
    return []
  }
  const range = resolveVlookupTableDependencyRange(node.args[1], scan, cell)
  if (!range) {
    return []
  }
  if (range.cellCount <= 0 || range.cellCount > maxNativeRangeAggregateCellCount) {
    return []
  }
  const rows: ScalarDependencyRow[] = []
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    rows.push({ sheetName: range.sheetName, row })
  }
  return rows
}

interface VlookupTableDependencyRange {
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly cellCount: number
}

function resolveVlookupTableDependencyRange(
  node: FormulaNode | undefined,
  scan: SheetScanState,
  cell: NativeFormulaCell,
): VlookupTableDependencyRange | null {
  if (!node) {
    return null
  }
  if (node.kind === 'RangeRef') {
    return decodeVlookupTableDependencyRange(node, scan.sheetName)
  }
  if (node.kind === 'CallExpr' && normalizedFormulaFunctionName(node.callee) === 'INDIRECT') {
    const referenceText = tryEvaluateDependencyText(node.args[0], scan, cell)
    if (referenceText === null) {
      return null
    }
    let range: FormulaNode
    try {
      range = parseFormula(referenceText)
    } catch {
      return null
    }
    return range.kind === 'RangeRef' ? decodeVlookupTableDependencyRange(range, scan.sheetName) : null
  }
  return null
}

function decodeVlookupTableDependencyRange(
  range: Extract<FormulaNode, { readonly kind: 'RangeRef' }>,
  currentSheetName: string,
): VlookupTableDependencyRange | null {
  if (range.refKind !== 'cells' || range.sheetEndName !== undefined) {
    return null
  }
  let start
  let end
  try {
    start = decodeCellAddress(range.start.replaceAll('$', ''))
    end = decodeCellAddress(range.end.replaceAll('$', ''))
  } catch {
    return null
  }
  const startRow = Math.min(start.r, end.r)
  const endRow = Math.max(start.r, end.r)
  const startCol = Math.min(start.c, end.c)
  const endCol = Math.max(start.c, end.c)
  return {
    sheetName: range.sheetName ?? currentSheetName,
    startRow,
    endRow,
    cellCount: (endRow - startRow + 1) * (endCol - startCol + 1),
  }
}

function tryEvaluateDependencyText(node: FormulaNode | undefined, scan: SheetScanState, cell: NativeFormulaCell): string | null {
  if (!node) {
    return null
  }
  switch (node.kind) {
    case 'StringLiteral':
      return node.value
    case 'NumberLiteral':
      return Number.isFinite(node.value) ? String(node.value) : null
    case 'BooleanLiteral':
      return node.value ? 'TRUE' : 'FALSE'
    case 'CellRef':
      return cellText(readDependencyCell(node, scan, cell))
    case 'BinaryExpr':
      if (node.operator !== '&') {
        return null
      }
      {
        const left = tryEvaluateDependencyText(node.left, scan, cell)
        const right = tryEvaluateDependencyText(node.right, scan, cell)
        return left === null || right === null ? null : left + right
      }
    case 'ArrayConstant':
    case 'CallExpr':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'InvokeExpr':
    case 'NameRef':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StructuredRef':
    case 'UnaryExpr':
      return null
  }
}

function readDependencyCell(
  node: Extract<FormulaNode, { readonly kind: 'CellRef' }>,
  scan: SheetScanState,
  cell: NativeFormulaCell,
): CellValue | null {
  if (node.sheetName && node.sheetName !== scan.sheetName) {
    return null
  }
  let address
  try {
    address = decodeCellAddress(node.ref.replaceAll('$', ''))
  } catch {
    return null
  }
  const rowValues = address.r === cell.row ? scan.rows.get(cell.row) : scan.rows.get(address.r)
  return resolvedCellValue(rowValues?.get(address.c))
}

function cellText(value: CellValue | null): string | null {
  if (!value) {
    return null
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return null
    default:
      return null
  }
}

function normalizedFormulaFunctionName(name: string): string {
  return name.toLocaleUpperCase('en-US').replace(/^_XLFN\./u, '')
}

interface ScalarDependencyRow {
  readonly sheetName: string
  readonly row: number
}

function collectScalarDependencyRows(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
): readonly ScalarDependencyRow[] {
  let node: FormulaNode
  try {
    node = parseFormula(resolveFormulaSource(scan, cell))
  } catch {
    return []
  }
  const rows = new Map<string, Set<number>>()
  collectScalarDependencyRowsFromNode(node, scan.sheetName, cell.row, rows)
  return [...rows].flatMap(([sheetName, sheetRows]) => [...sheetRows].map((row) => ({ sheetName, row })))
}

function collectScalarDependencyRowsFromNode(
  node: FormulaNode,
  sheetName: string,
  formulaRow: number,
  rows: Map<string, Set<number>>,
): void {
  switch (node.kind) {
    case 'CellRef': {
      const dependencySheetName = node.sheetName ?? sheetName
      try {
        const address = decodeCellAddress(node.ref.replaceAll('$', ''))
        if (dependencySheetName !== sheetName || address.r !== formulaRow) {
          const sheetRows = rows.get(dependencySheetName) ?? new Set<number>()
          sheetRows.add(address.r)
          rows.set(dependencySheetName, sheetRows)
        }
      } catch {
        return
      }
      return
    }
    case 'UnaryExpr':
      collectScalarDependencyRowsFromNode(node.argument, sheetName, formulaRow, rows)
      return
    case 'BinaryExpr':
      collectScalarDependencyRowsFromNode(node.left, sheetName, formulaRow, rows)
      collectScalarDependencyRowsFromNode(node.right, sheetName, formulaRow, rows)
      return
    case 'CallExpr':
      for (const argument of node.args) {
        collectScalarDependencyRowsFromNode(argument, sheetName, formulaRow, rows)
      }
      return
    case 'InvokeExpr':
      collectScalarDependencyRowsFromNode(node.callee, sheetName, formulaRow, rows)
      for (const argument of node.args) {
        collectScalarDependencyRowsFromNode(argument, sheetName, formulaRow, rows)
      }
      return
    case 'ArrayConstant':
      for (const row of node.rows) {
        for (const entry of row) {
          collectScalarDependencyRowsFromNode(entry, sheetName, formulaRow, rows)
        }
      }
      return
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
      return
    case 'RangeRef':
      collectCellRangeDependencyRows(node, sheetName, formulaRow, rows)
      return
  }
}

function collectCellRangeDependencyRows(
  range: Extract<FormulaNode, { readonly kind: 'RangeRef' }>,
  currentSheetName: string,
  formulaRow: number,
  rows: Map<string, Set<number>>,
): void {
  const dependencyRange = decodeVlookupTableDependencyRange(range, currentSheetName)
  if (!dependencyRange || dependencyRange.cellCount > maxNativeRangeAggregateCellCount) {
    return
  }
  for (let row = dependencyRange.startRow; row <= dependencyRange.endRow; row += 1) {
    if (dependencyRange.sheetName === currentSheetName && row === formulaRow) {
      continue
    }
    const sheetRows = rows.get(dependencyRange.sheetName) ?? new Set<number>()
    sheetRows.add(row)
    rows.set(dependencyRange.sheetName, sheetRows)
  }
}

export function tryCompileNativeRangeAggregate(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
): { readonly aggregateKind: number; readonly range: CellRangePlan } | null {
  let node: FormulaNode
  try {
    node = parseFormula(resolveFormulaSource(scan, cell))
  } catch {
    return null
  }
  if (node.kind !== 'CallExpr' || node.args.length !== 1) {
    return null
  }
  const aggregateKind = directRangeAggregateKind(node.callee)
  if (aggregateKind === null) {
    return null
  }
  const range = node.args[0]
  if (
    !range ||
    range.kind !== 'RangeRef' ||
    range.refKind !== 'cells' ||
    range.sheetEndName !== undefined ||
    (range.sheetName !== undefined && range.sheetName !== scan.sheetName)
  ) {
    return null
  }
  let start
  let end
  try {
    start = decodeCellAddress(range.start.replaceAll('$', ''))
    end = decodeCellAddress(range.end.replaceAll('$', ''))
  } catch {
    return null
  }
  const startRow = Math.min(start.r, end.r)
  const endRow = Math.max(start.r, end.r)
  const startCol = Math.min(start.c, end.c)
  const endCol = Math.max(start.c, end.c)
  const cellCount = (endRow - startRow + 1) * (endCol - startCol + 1)
  return cellCount > 0 && cellCount <= maxNativeRangeAggregateCellCount
    ? {
        aggregateKind,
        range: {
          startRow,
          endRow,
          startCol,
          endCol,
          cellCount,
        },
      }
    : null
}

function directRangeAggregateKind(callee: string): number | null {
  switch (normalizedFormulaFunctionName(callee)) {
    case 'SUM':
      return directAggregateOpSum
    case 'AVERAGE':
      return directAggregateOpAverage
    case 'COUNTA':
      return directAggregateOpCounta
    default:
      return null
  }
}

function resolvedCellValue(value: PendingCellValue | undefined): CellValue {
  if (value === undefined || isSharedStringReference(value)) {
    return { tag: ValueTag.Empty }
  }
  return value
}

function isSharedStringReference(value: PendingCellValue): value is Extract<PendingCellValue, { readonly kind: 'shared-string' }> {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}
