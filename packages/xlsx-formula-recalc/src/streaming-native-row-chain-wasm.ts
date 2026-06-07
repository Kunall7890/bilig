import { parseFormula, type FormulaNode } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createKernelSync } from '@bilig/wasm-kernel'
import { decodeCellAddress, type XlsxSourceLiteralPatch } from '@bilig/xlsx'

import type { NativeFormulaCell, NativeTable, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'

export interface StreamingNativeWasmRowChainResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

export interface StreamingNativeWasmDirectScalarResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

export interface StreamingNativeWasmRangeAggregateResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

export interface StreamingNativeWasmConditionalResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

export interface StreamingNativeWasmFormulaResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

interface RowChainCandidate {
  readonly firstCell: NativeFormulaCell
  readonly secondCell: NativeFormulaCell
  readonly firstOperatorCode: number
  readonly leftValue: number
  readonly rightValue: number
  readonly rowValues: Map<number, PendingCellValue>
}

interface AffineRowChainCandidate extends RowChainCandidate {
  readonly secondKind: 'affine'
  readonly secondScale: number
  readonly secondOffset: number
}

interface DivideRowChainCandidate extends RowChainCandidate {
  readonly secondKind: 'divide'
  readonly denominatorValue: number
}

interface DirectScalarCandidate {
  readonly cell: NativeFormulaCell
  readonly operatorCode: number
  readonly leftValue: number
  readonly rightValue: number
  readonly rowValues: Map<number, PendingCellValue>
}

interface RangeAggregateCandidate {
  readonly cell: NativeFormulaCell
  readonly values: readonly number[]
  readonly rowValues: Map<number, PendingCellValue>
}

interface ConditionalPickCandidate {
  readonly cell: NativeFormulaCell
  readonly conditions: readonly ConditionalPickCondition[]
  readonly branchValues: readonly CellValue[]
  readonly defaultValue: CellValue
  readonly rowValues: Map<number, PendingCellValue>
}

interface ConditionalPickCondition {
  readonly operatorCode: number
  readonly left: CellValue
  readonly right: CellValue
}

interface ConditionalCompileContext {
  readonly sheetName: string
  readonly row: number
  readonly formulaColumn: number
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly rowValues: ReadonlyMap<number, PendingCellValue>
}

interface NumericSource {
  readonly constant?: number
  readonly column?: number
}

interface CellRangePlan {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
  readonly cellCount: number
}

interface FirstFormulaPlan {
  readonly operatorCode: number
  readonly left: NumericSource
  readonly right: NumericSource
}

interface AffineSecondFormulaPlan {
  readonly kind: 'affine'
  readonly firstFormulaColumn: number
  readonly scale: number
  readonly offset: number
}

interface DivideSecondFormulaPlan {
  readonly kind: 'divide'
  readonly firstFormulaColumn: number
  readonly denominator: NumericSource
}

interface InlineDivideSecondFormulaPlan {
  readonly kind: 'inline-divide'
  readonly firstFormula: FirstFormulaPlan
  readonly denominator: NumericSource
}

type SecondFormulaPlan = AffineSecondFormulaPlan | DivideSecondFormulaPlan | InlineDivideSecondFormulaPlan
type CompiledRowChainCandidate = AffineRowChainCandidate | DivideRowChainCandidate

const directAggregateOpSum = 1
const directConditionOpTruthy = 1
const directConditionOpEq = 2
const directConditionOpNeq = 3
const directConditionOpGt = 4
const directConditionOpGte = 5
const directConditionOpLt = 6
const directConditionOpLte = 7
const maxExpandedFormulaDependencyRowsPerSheet = 20_000
const maxNativeSumRangeCellCount = 50_000

export function evaluateStreamingNativeWasmRowChains(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
}): StreamingNativeWasmRowChainResult {
  const candidates = collectRowChainCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const kernel = createKernelSync()
  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  let batchCount = 0
  for (const group of groupCandidates(candidates).values()) {
    if (group.length === 0) {
      continue
    }
    batchCount += 1
    kernel.init(group.length * 2, 1, 1, 1, 1)
    const firstCandidate = group[0]!
    if (firstCandidate.secondKind === 'affine') {
      const affineGroup = group.filter(isAffineRowChainCandidate)
      kernel.evalDenseDirectScalarRowChainStoreTargetBatch(
        Float64Array.from(affineGroup.map((candidate) => candidate.leftValue)),
        Float64Array.from(affineGroup.map((candidate) => candidate.rightValue)),
        Uint32Array.from(affineGroup.map((_candidate, index) => index * 2)),
        Uint32Array.from(affineGroup.map((_candidate, index) => index * 2 + 1)),
        affineGroup.length,
        firstCandidate.firstOperatorCode,
        firstCandidate.secondScale,
        firstCandidate.secondOffset,
      )
    } else {
      const divideGroup = group.filter(isDivideRowChainCandidate)
      kernel.evalDenseDirectScalarRowChainDivideStoreTargetBatch(
        Float64Array.from(divideGroup.map((candidate) => candidate.leftValue)),
        Float64Array.from(divideGroup.map((candidate) => candidate.rightValue)),
        Float64Array.from(divideGroup.map((candidate) => candidate.denominatorValue)),
        Uint32Array.from(divideGroup.map((_candidate, index) => index * 2)),
        Uint32Array.from(divideGroup.map((_candidate, index) => index * 2 + 1)),
        divideGroup.length,
        firstCandidate.firstOperatorCode,
      )
    }
    const numbers = kernel.readNumbers()
    for (let index = 0; index < group.length; index += 1) {
      const candidate = group[index]!
      const firstValue = numbers[index * 2]!
      const secondValue = numbers[index * 2 + 1]!
      if (!Number.isFinite(firstValue) || !Number.isFinite(secondValue)) {
        continue
      }
      candidate.rowValues.set(candidate.firstCell.col, { tag: ValueTag.Number, value: firstValue })
      candidate.rowValues.set(candidate.secondCell.col, { tag: ValueTag.Number, value: secondValue })
      patches.push(formulaPatch(candidate.firstCell, firstValue), formulaPatch(candidate.secondCell, secondValue))
      processedCells.add(cellKey(candidate.firstCell))
      processedCells.add(cellKey(candidate.secondCell))
    }
  }
  return {
    batchCount,
    evaluatedFormulaCellCount: processedCells.size,
    patches,
    processedCells,
  }
}

export function evaluateStreamingNativeWasmFormulas(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
}): StreamingNativeWasmFormulaResult {
  const rowChains = evaluateStreamingNativeWasmRowChains(args)
  const directScalars = evaluateStreamingNativeWasmDirectScalars({
    ...args,
    skippedCells: rowChains.processedCells,
  })
  const rangeAggregates = evaluateStreamingNativeWasmRangeAggregates({
    sheetScans: args.sheetScans,
    resolveFormulaSource: args.resolveFormulaSource,
    skippedCells: new Set([...rowChains.processedCells, ...directScalars.processedCells]),
  })
  const conditionals = evaluateStreamingNativeWasmConditionals({
    ...args,
    skippedCells: new Set([...rowChains.processedCells, ...directScalars.processedCells, ...rangeAggregates.processedCells]),
  })
  return {
    batchCount: rowChains.batchCount + directScalars.batchCount + rangeAggregates.batchCount + conditionals.batchCount,
    evaluatedFormulaCellCount:
      rowChains.evaluatedFormulaCellCount +
      directScalars.evaluatedFormulaCellCount +
      rangeAggregates.evaluatedFormulaCellCount +
      conditionals.evaluatedFormulaCellCount,
    patches: [...rowChains.patches, ...directScalars.patches, ...rangeAggregates.patches, ...conditionals.patches],
    processedCells: new Set([
      ...rowChains.processedCells,
      ...directScalars.processedCells,
      ...rangeAggregates.processedCells,
      ...conditionals.processedCells,
    ]),
  }
}

export function evaluateStreamingNativeWasmDirectScalars(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): StreamingNativeWasmDirectScalarResult {
  const candidates = collectDirectScalarCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const kernel = createKernelSync()
  kernel.init(candidates.length, 1, 1, 1, 1)
  kernel.evalDirectScalarStoreTargetBatch(
    Uint32Array.from(candidates.map((_candidate, index) => index)),
    Uint8Array.from(candidates.map((candidate) => candidate.operatorCode)),
    Uint32Array.from(candidates.map(() => 0xffffffff)),
    Uint8Array.from(candidates.map(() => ValueTag.Number)),
    Float64Array.from(candidates.map((candidate) => candidate.leftValue)),
    Uint16Array.from(candidates.map(() => ErrorCode.None)),
    Uint32Array.from(candidates.map(() => 0xffffffff)),
    Uint8Array.from(candidates.map(() => ValueTag.Number)),
    Float64Array.from(candidates.map((candidate) => candidate.rightValue)),
    Uint16Array.from(candidates.map(() => ErrorCode.None)),
    Float64Array.from(candidates.map(() => 0)),
  )
  const tags = kernel.readTags()
  const numbers = kernel.readNumbers()
  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  for (let index = 0; index < candidates.length; index += 1) {
    const candidate = candidates[index]!
    const value = numbers[index]!
    if (tags[index] !== ValueTag.Number || !Number.isFinite(value)) {
      continue
    }
    candidate.rowValues.set(candidate.cell.col, { tag: ValueTag.Number, value })
    patches.push(formulaPatch(candidate.cell, value))
    processedCells.add(cellKey(candidate.cell))
  }
  return {
    batchCount: patches.length > 0 ? 1 : 0,
    evaluatedFormulaCellCount: processedCells.size,
    patches,
    processedCells,
  }
}

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
      const range = tryCompileNativeSumRange(scan, cell, args.resolveFormulaSource)
      if (!range) {
        continue
      }
      for (let row = range.startRow; row <= range.endRow; row += 1) {
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
  if (range.cellCount <= 0 || range.cellCount > maxNativeSumRangeCellCount) {
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
    case 'RangeRef':
      return
  }
}

export function evaluateStreamingNativeWasmRangeAggregates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): StreamingNativeWasmRangeAggregateResult {
  const candidates = collectRangeAggregateCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const kernel = createKernelSync()
  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  let batchCount = 0
  for (const group of groupRangeAggregateCandidates(candidates).values()) {
    if (group.length === 0) {
      continue
    }
    batchCount += 1
    const valueCount = group[0]!.values.length
    const outNumbers = new Float64Array(group.length)
    kernel.evalDenseNumericRowAggregateBatch(
      directAggregateOpSum,
      Float64Array.from(group.flatMap((candidate) => candidate.values)),
      group.length,
      valueCount,
      0,
      valueCount,
      0,
      outNumbers,
    )
    for (let index = 0; index < group.length; index += 1) {
      const candidate = group[index]!
      const value = outNumbers[index]!
      if (!Number.isFinite(value)) {
        continue
      }
      candidate.rowValues.set(candidate.cell.col, { tag: ValueTag.Number, value })
      patches.push(formulaPatch(candidate.cell, value))
      processedCells.add(cellKey(candidate.cell))
    }
  }
  return {
    batchCount,
    evaluatedFormulaCellCount: processedCells.size,
    patches,
    processedCells,
  }
}

export function evaluateStreamingNativeWasmConditionals(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): StreamingNativeWasmConditionalResult {
  const candidates = collectConditionalPickCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const stringArena = new KernelStringArena()
  const conditionStarts: number[] = []
  const conditionLengths: number[] = []
  const conditionOps: number[] = []
  const leftTags: number[] = []
  const leftNumbers: number[] = []
  const leftStringIds: number[] = []
  const leftErrors: number[] = []
  const rightTags: number[] = []
  const rightNumbers: number[] = []
  const rightStringIds: number[] = []
  const rightErrors: number[] = []
  const branchTags: number[] = []
  const branchNumbers: number[] = []
  const branchStringIds: number[] = []
  const branchErrors: number[] = []
  const defaultTags: number[] = []
  const defaultNumbers: number[] = []
  const defaultStringIds: number[] = []
  const defaultErrors: number[] = []
  for (const candidate of candidates) {
    conditionStarts.push(conditionOps.length)
    conditionLengths.push(candidate.conditions.length)
    for (let index = 0; index < candidate.conditions.length; index += 1) {
      const condition = candidate.conditions[index]!
      conditionOps.push(condition.operatorCode)
      appendKernelValueFields(condition.left, stringArena, 'compare', leftTags, leftNumbers, leftStringIds, leftErrors)
      appendKernelValueFields(condition.right, stringArena, 'compare', rightTags, rightNumbers, rightStringIds, rightErrors)
      appendKernelValueFields(
        candidate.branchValues[index]!,
        stringArena,
        'output',
        branchTags,
        branchNumbers,
        branchStringIds,
        branchErrors,
      )
    }
    appendKernelValueFields(candidate.defaultValue, stringArena, 'output', defaultTags, defaultNumbers, defaultStringIds, defaultErrors)
  }

  const outTags = new Uint8Array(candidates.length)
  const outNumbers = new Float64Array(candidates.length)
  const outStringIds = new Uint32Array(candidates.length)
  const outErrors = new Uint16Array(candidates.length)
  createKernelSync().evalDirectConditionalPickBatch(
    Uint32Array.from(conditionStarts),
    Uint32Array.from(conditionLengths),
    Uint8Array.from(conditionOps),
    Uint8Array.from(leftTags),
    Float64Array.from(leftNumbers),
    Uint32Array.from(leftStringIds),
    Uint16Array.from(leftErrors),
    Uint8Array.from(rightTags),
    Float64Array.from(rightNumbers),
    Uint32Array.from(rightStringIds),
    Uint16Array.from(rightErrors),
    Uint8Array.from(branchTags),
    Float64Array.from(branchNumbers),
    Uint32Array.from(branchStringIds),
    Uint16Array.from(branchErrors),
    Uint8Array.from(defaultTags),
    Float64Array.from(defaultNumbers),
    Uint32Array.from(defaultStringIds),
    Uint16Array.from(defaultErrors),
    outTags,
    outNumbers,
    outStringIds,
    outErrors,
  )

  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  for (let index = 0; index < candidates.length; index += 1) {
    const value = cellValueFromKernelFields(outTags[index]!, outNumbers[index]!, outStringIds[index]!, outErrors[index]!, stringArena)
    const patchValue = value ? cellValueToPatchValue(value) : undefined
    if (!value || patchValue === undefined) {
      continue
    }
    const candidate = candidates[index]!
    candidate.rowValues.set(candidate.cell.col, value)
    patches.push(formulaPatch(candidate.cell, patchValue))
    processedCells.add(cellKey(candidate.cell))
  }
  return {
    batchCount: processedCells.size > 0 ? 1 : 0,
    evaluatedFormulaCellCount: processedCells.size,
    patches,
    processedCells,
  }
}

function collectRowChainCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
}): CompiledRowChainCandidate[] {
  const candidates: CompiledRowChainCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    const formulaCellsByRow = formulaCellsByRowNumber(scan.formulaCells)
    for (const [row, rowFormulaCells] of formulaCellsByRow.entries()) {
      const rowValues = scan.rows.get(row)
      if (!rowValues) {
        continue
      }
      const formulaCellsByColumn = new Map(rowFormulaCells.map((cell) => [cell.col, cell]))
      for (const secondCell of rowFormulaCells) {
        const secondPlan = tryCompileSecondFormula(scan, secondCell, args)
        if (!secondPlan) {
          continue
        }
        const firstCandidate = resolveFirstFormulaCandidate(scan, secondCell, secondPlan, rowFormulaCells, formulaCellsByColumn, args)
        if (!firstCandidate) {
          continue
        }
        const { firstCell, firstPlan } = firstCandidate
        const leftValue = readNumericSource(rowValues, firstPlan.left)
        const rightValue = readNumericSource(rowValues, firstPlan.right)
        if (leftValue === null || rightValue === null || (firstPlan.operatorCode === 5 && rightValue === 0)) {
          continue
        }
        if (secondPlan.kind === 'affine') {
          candidates.push({
            firstCell,
            secondCell,
            firstOperatorCode: firstPlan.operatorCode,
            secondKind: 'affine',
            secondScale: secondPlan.scale,
            secondOffset: secondPlan.offset,
            leftValue,
            rightValue,
            rowValues,
          })
          continue
        }
        const denominatorValue = readNumericSource(rowValues, secondPlan.denominator)
        if (denominatorValue === null || denominatorValue === 0) {
          continue
        }
        candidates.push({
          firstCell,
          secondCell,
          firstOperatorCode: firstPlan.operatorCode,
          secondKind: 'divide',
          denominatorValue,
          leftValue,
          rightValue,
          rowValues,
        })
      }
    }
  }
  return candidates
}

function collectDirectScalarCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): DirectScalarCandidate[] {
  const candidates: DirectScalarCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      if (args.skippedCells.has(cellKey(cell))) {
        continue
      }
      const rowValues = scan.rows.get(cell.row)
      if (!rowValues) {
        continue
      }
      const firstPlan = tryCompileFirstFormula(scan, cell, args)
      const operatorCode = firstPlan ? directScalarOperatorCode(firstPlan.operatorCode) : undefined
      if (!firstPlan || operatorCode === undefined) {
        continue
      }
      const leftValue = readNumericSource(rowValues, firstPlan.left)
      const rightValue = readNumericSource(rowValues, firstPlan.right)
      if (leftValue === null || rightValue === null || (firstPlan.operatorCode === 5 && rightValue === 0)) {
        continue
      }
      candidates.push({
        cell,
        operatorCode,
        leftValue,
        rightValue,
        rowValues,
      })
    }
  }
  return candidates
}

function collectRangeAggregateCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): RangeAggregateCandidate[] {
  const candidates: RangeAggregateCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      if (args.skippedCells.has(cellKey(cell))) {
        continue
      }
      const rowValues = scan.rows.get(cell.row)
      if (!rowValues) {
        continue
      }
      const range = tryCompileNativeSumRange(scan, cell, args.resolveFormulaSource)
      if (!range) {
        continue
      }
      const values = readNumericRangeValues(scan, range)
      if (!values) {
        continue
      }
      candidates.push({
        cell,
        values,
        rowValues,
      })
    }
  }
  return candidates
}

function collectConditionalPickCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): ConditionalPickCandidate[] {
  const candidates: ConditionalPickCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      if (args.skippedCells.has(cellKey(cell))) {
        continue
      }
      const rowValues = scan.rows.get(cell.row)
      if (!rowValues) {
        continue
      }
      const candidate = tryCompileConditionalPickCandidate(scan, cell, rowValues, args)
      if (candidate) {
        candidates.push(candidate)
      }
    }
  }
  return candidates
}

function tryCompileConditionalPickCandidate(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  rowValues: Map<number, PendingCellValue>,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): ConditionalPickCandidate | null {
  let node: FormulaNode
  try {
    node = parseFormula(args.resolveFormulaSource(scan, cell))
  } catch {
    return null
  }
  const context = { sheetName: scan.sheetName, row: cell.row, formulaColumn: cell.col, tablesBySheet: args.tablesBySheet, rowValues }
  if (node.kind === 'CallExpr') {
    const callee = node.callee.toLocaleUpperCase('en-US').replace(/^_XLFN\./u, '')
    if (callee === 'IF') {
      return compileNativeIfCandidate(cell, rowValues, node.args, context)
    }
    if (callee === 'IFS') {
      return compileNativeIfsCandidate(cell, rowValues, node.args, context)
    }
  }
  const value = compileScalarValue(node, context)
  return value ? { cell, rowValues, conditions: [], branchValues: [], defaultValue: value } : null
}

function compileNativeIfCandidate(
  cell: NativeFormulaCell,
  rowValues: Map<number, PendingCellValue>,
  args: readonly FormulaNode[],
  context: ConditionalCompileContext,
): ConditionalPickCandidate | null {
  if (args.length < 2 || args.length > 3) {
    return null
  }
  const condition = compileConditionalCondition(args[0]!, context)
  const trueValue = compileScalarValue(args[1]!, context)
  const falseValue = args[2] ? compileScalarValue(args[2], context) : ({ tag: ValueTag.Boolean, value: false } satisfies CellValue)
  return condition && trueValue && falseValue
    ? {
        cell,
        rowValues,
        conditions: [condition],
        branchValues: [trueValue],
        defaultValue: falseValue,
      }
    : null
}

function compileNativeIfsCandidate(
  cell: NativeFormulaCell,
  rowValues: Map<number, PendingCellValue>,
  args: readonly FormulaNode[],
  context: ConditionalCompileContext,
): ConditionalPickCandidate | null {
  if (args.length < 2 || args.length % 2 !== 0) {
    return null
  }
  const conditions: ConditionalPickCondition[] = []
  const branchValues: CellValue[] = []
  for (let index = 0; index < args.length; index += 2) {
    const condition = compileConditionalCondition(args[index]!, context)
    const value = compileScalarValue(args[index + 1]!, context)
    if (!condition || !value) {
      return null
    }
    conditions.push(condition)
    branchValues.push(value)
  }
  return {
    cell,
    rowValues,
    conditions,
    branchValues,
    defaultValue: { tag: ValueTag.Error, code: ErrorCode.NA },
  }
}

function compileConditionalCondition(node: FormulaNode, context: ConditionalCompileContext): ConditionalPickCondition | null {
  if (node.kind === 'BinaryExpr') {
    const operatorCode = directConditionOperatorCode(node.operator)
    if (operatorCode === undefined) {
      return null
    }
    const left = compileScalarValue(node.left, context)
    const right = compileScalarValue(node.right, context)
    if (!left || !right || !conditionOperandsAreSupported(operatorCode, left, right)) {
      return null
    }
    return { operatorCode, left, right }
  }
  const left = compileScalarValue(node, context)
  return left && conditionOperandsAreSupported(directConditionOpTruthy, left, emptyCellValue())
    ? { operatorCode: directConditionOpTruthy, left, right: emptyCellValue() }
    : null
}

function compileScalarValue(node: FormulaNode, context: ConditionalCompileContext): CellValue | null {
  switch (node.kind) {
    case 'NumberLiteral':
      return Number.isFinite(node.value) ? { tag: ValueTag.Number, value: node.value } : null
    case 'BooleanLiteral':
      return { tag: ValueTag.Boolean, value: node.value }
    case 'StringLiteral':
      return stringCellValue(node.value)
    case 'CellRef':
    case 'StructuredRef':
      return readRowLocalScalarValue(node, context)
    case 'ArrayConstant':
    case 'BinaryExpr':
    case 'CallExpr':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'InvokeExpr':
    case 'NameRef':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'UnaryExpr':
      return null
    default:
      return null
  }
}

function readRowLocalScalarValue(
  node: Extract<FormulaNode, { readonly kind: 'CellRef' | 'StructuredRef' }>,
  context: ConditionalCompileContext,
): CellValue | null {
  if (node.kind === 'CellRef') {
    if (node.sheetName && node.sheetName !== context.sheetName) {
      return null
    }
    let address
    try {
      address = decodeCellAddress(node.ref.replaceAll('$', ''))
    } catch {
      return null
    }
    return address.r === context.row ? resolvedCellValue(context.rowValues.get(address.c)) : null
  }
  const column = compileReferenceColumn(node, context.sheetName, context.row, context.formulaColumn, context.tablesBySheet)
  return column === null ? null : resolvedCellValue(context.rowValues.get(column))
}

function directConditionOperatorCode(operator: string): number | undefined {
  switch (operator) {
    case '=':
      return directConditionOpEq
    case '<>':
      return directConditionOpNeq
    case '>':
      return directConditionOpGt
    case '>=':
      return directConditionOpGte
    case '<':
      return directConditionOpLt
    case '<=':
      return directConditionOpLte
    default:
      return undefined
  }
}

function conditionOperandsAreSupported(operatorCode: number, left: CellValue, right: CellValue): boolean {
  if (left.tag === ValueTag.Error || right.tag === ValueTag.Error) {
    return true
  }
  if (operatorCode === directConditionOpTruthy) {
    return left.tag === ValueTag.Number || left.tag === ValueTag.Boolean || left.tag === ValueTag.Empty
  }
  const leftNumeric = isNumericComparable(left)
  const rightNumeric = isNumericComparable(right)
  if (leftNumeric && rightNumeric) {
    return true
  }
  return (
    (operatorCode === directConditionOpEq || operatorCode === directConditionOpNeq) &&
    (left.tag === ValueTag.String || left.tag === ValueTag.Empty) &&
    (right.tag === ValueTag.String || right.tag === ValueTag.Empty)
  )
}

function tryCompileFirstFormula(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): FirstFormulaPlan | null {
  try {
    return compileFirstFormula(parseFormula(args.resolveFormulaSource(scan, cell)), scan.sheetName, cell.row, cell.col, args.tablesBySheet)
  } catch {
    return null
  }
}

function tryCompileSecondFormula(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): SecondFormulaPlan | null {
  try {
    return compileSecondFormula(parseFormula(args.resolveFormulaSource(scan, cell)), scan.sheetName, cell.row, cell.col, args.tablesBySheet)
  } catch {
    return null
  }
}

function tryCompileNativeSumRange(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string,
): CellRangePlan | null {
  let node: FormulaNode
  try {
    node = parseFormula(resolveFormulaSource(scan, cell))
  } catch {
    return null
  }
  if (node.kind !== 'CallExpr' || node.callee.toLocaleUpperCase('en-US') !== 'SUM' || node.args.length !== 1) {
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
  return cellCount > 0 && cellCount <= maxNativeSumRangeCellCount
    ? {
        startRow,
        endRow,
        startCol,
        endCol,
        cellCount,
      }
    : null
}

function resolveFirstFormulaCandidate(
  scan: SheetScanState,
  secondCell: NativeFormulaCell,
  secondPlan: SecondFormulaPlan,
  rowFormulaCells: readonly NativeFormulaCell[],
  formulaCellsByColumn: ReadonlyMap<number, NativeFormulaCell>,
  args: {
    readonly tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): { readonly firstCell: NativeFormulaCell; readonly firstPlan: FirstFormulaPlan } | null {
  if (secondPlan.kind !== 'inline-divide') {
    const firstCell = formulaCellsByColumn.get(secondPlan.firstFormulaColumn)
    if (!firstCell) {
      return null
    }
    const firstPlan = tryCompileFirstFormula(scan, firstCell, args)
    return firstPlan ? { firstCell, firstPlan } : null
  }
  for (const firstCell of rowFormulaCells) {
    if (firstCell.col === secondCell.col) {
      continue
    }
    const firstPlan = tryCompileFirstFormula(scan, firstCell, args)
    if (firstPlan && firstFormulaPlansMatch(firstPlan, secondPlan.firstFormula)) {
      return { firstCell, firstPlan }
    }
  }
  return null
}

function firstFormulaPlansMatch(left: FirstFormulaPlan, right: FirstFormulaPlan): boolean {
  if (left.operatorCode !== right.operatorCode) {
    return false
  }
  if (numericSourcesEqual(left.left, right.left) && numericSourcesEqual(left.right, right.right)) {
    return true
  }
  return (
    isCommutativeRowPairOperator(left.operatorCode) &&
    numericSourcesEqual(left.left, right.right) &&
    numericSourcesEqual(left.right, right.left)
  )
}

function isCommutativeRowPairOperator(operatorCode: number): boolean {
  return operatorCode === 1 || operatorCode === 4
}

function numericSourcesEqual(left: NumericSource, right: NumericSource): boolean {
  return left.constant === right.constant && left.column === right.column
}

function formulaCellsByRowNumber(formulaCells: readonly NativeFormulaCell[]): Map<number, NativeFormulaCell[]> {
  const output = new Map<number, NativeFormulaCell[]>()
  for (const cell of formulaCells) {
    const rowCells = output.get(cell.row) ?? []
    rowCells.push(cell)
    output.set(cell.row, rowCells)
  }
  return output
}

function compileFirstFormula(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): FirstFormulaPlan | null {
  if (node.kind !== 'BinaryExpr') {
    return null
  }
  const operatorCode = rowPairOperatorCode(node.operator)
  if (operatorCode === undefined) {
    return null
  }
  const left = compileNumericSource(node.left, sheetName, row, formulaColumn, tablesBySheet)
  const right = compileNumericSource(node.right, sheetName, row, formulaColumn, tablesBySheet)
  return left && right ? { operatorCode, left, right } : null
}

function rowPairOperatorCode(operator: string): number | undefined {
  switch (operator) {
    case '+':
      return 1
    case '-':
      return 2
    case '*':
      return 4
    case '/':
      return 5
    default:
      return undefined
  }
}

function directScalarOperatorCode(rowPairCode: number): number | undefined {
  switch (rowPairCode) {
    case 1:
      return 1
    case 2:
      return 2
    case 4:
      return 3
    case 5:
      return 4
    default:
      return undefined
  }
}

function compileSecondFormula(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): SecondFormulaPlan | null {
  if (node.kind !== 'BinaryExpr') {
    return null
  }
  const leftReference = compileReferenceColumn(node.left, sheetName, row, formulaColumn, tablesBySheet)
  const rightReference = compileReferenceColumn(node.right, sheetName, row, formulaColumn, tablesBySheet)
  if (node.operator === '*') {
    const leftNumber = numberLiteralValue(node.left)
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: leftReference, scale: rightNumber, offset: 0 }
    }
    if (rightReference !== null && leftNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: rightReference, scale: leftNumber, offset: 0 }
    }
  }
  if (node.operator === '+') {
    const leftNumber = numberLiteralValue(node.left)
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: leftReference, scale: 1, offset: rightNumber }
    }
    if (rightReference !== null && leftNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: rightReference, scale: 1, offset: leftNumber }
    }
  }
  if (node.operator === '-') {
    const rightNumber = numberLiteralValue(node.right)
    if (leftReference !== null && rightNumber !== null) {
      return { kind: 'affine', firstFormulaColumn: leftReference, scale: 1, offset: -rightNumber }
    }
  }
  if (node.operator === '/' && leftReference !== null) {
    const denominator = compileNumericSource(node.right, sheetName, row, formulaColumn, tablesBySheet)
    if (denominator) {
      return { kind: 'divide', firstFormulaColumn: leftReference, denominator }
    }
  }
  if (node.operator === '/' && leftReference === null) {
    const firstFormula = compileFirstFormula(node.left, sheetName, row, formulaColumn, tablesBySheet)
    const denominator = compileNumericSource(node.right, sheetName, row, formulaColumn, tablesBySheet)
    if (firstFormula && denominator) {
      return { kind: 'inline-divide', firstFormula, denominator }
    }
  }
  return null
}

function compileNumericSource(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): NumericSource | null {
  const constant = numberLiteralValue(node)
  if (constant !== null) {
    return { constant }
  }
  const column = compileReferenceColumn(node, sheetName, row, formulaColumn, tablesBySheet)
  return column === null ? null : { column }
}

function numberLiteralValue(node: FormulaNode): number | null {
  return node.kind === 'NumberLiteral' && Number.isFinite(node.value) ? node.value : null
}

function compileReferenceColumn(
  node: FormulaNode,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): number | null {
  if (node.kind === 'CellRef') {
    if (node.sheetName && node.sheetName !== sheetName) {
      return null
    }
    const address = decodeCellAddress(node.ref.replaceAll('$', ''))
    return address.r === row ? address.c : null
  }
  if (node.kind !== 'StructuredRef' || node.endColumnName || (node.section && node.section !== 'this-row')) {
    return null
  }
  const table = findCurrentRowTable(node.tableName, sheetName, row, formulaColumn, tablesBySheet)
  if (!table) {
    return null
  }
  const columnIndex = table.columns.findIndex(
    (column) => column.toLocaleLowerCase('en-US') === normalizeStructuredReferenceColumnName(node.columnName).toLocaleLowerCase('en-US'),
  )
  return columnIndex >= 0 ? table.range.s.c + columnIndex : null
}

function findCurrentRowTable(
  tableName: string,
  sheetName: string,
  row: number,
  formulaColumn: number,
  tablesBySheet: ReadonlyMap<string, readonly NativeTable[]>,
): NativeTable | null {
  const tables = tablesBySheet.get(sheetName) ?? []
  const matching = tables.filter((table) => {
    const nameMatches =
      tableName.length === 0 ||
      table.name.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US') ||
      table.displayName.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US')
    return nameMatches && rowIsInTableDataBody(table, row)
  })
  return matching.find((table) => formulaColumn >= table.range.s.c && formulaColumn <= table.range.e.c) ?? matching[0] ?? null
}

function rowIsInTableDataBody(table: NativeTable, row: number): boolean {
  const start = table.range.s.r + table.headerRowCount
  const end = table.range.e.r - table.totalsRowCount
  return row >= start && row <= end
}

function normalizeStructuredReferenceColumnName(value: string): string {
  return decodeExcelEscapedText(value).replace(/\r\n?/gu, '\n').trim()
}

function decodeExcelEscapedText(value: string): string {
  const escapedUnderscore = '\uE000'
  return value
    .replace(/_x005F_/giu, escapedUnderscore)
    .replace(/_x([0-9a-fA-F]{4})_/gu, (_match, code: string) => {
      const codePoint = Number.parseInt(code, 16)
      return Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : ''
    })
    .replaceAll(escapedUnderscore, '_')
}

function readNumericSource(rowValues: ReadonlyMap<number, PendingCellValue>, source: NumericSource): number | null {
  if (source.constant !== undefined) {
    return source.constant
  }
  if (source.column === undefined) {
    return null
  }
  const value = resolvedCellValue(rowValues.get(source.column))
  return value.tag === ValueTag.Number && Number.isFinite(value.value) ? value.value : null
}

function readNumericRangeValues(scan: SheetScanState, range: CellRangePlan): readonly number[] | null {
  const values: number[] = []
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const rowValues = scan.rows.get(row)
    if (!rowValues) {
      return null
    }
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const value = resolvedCellValue(rowValues.get(col))
      if (value.tag === ValueTag.Empty) {
        values.push(0)
        continue
      }
      if (value.tag !== ValueTag.Number || !Number.isFinite(value.value)) {
        return null
      }
      values.push(value.value)
    }
  }
  return values.length === range.cellCount ? values : null
}

function resolvedCellValue(value: PendingCellValue | undefined): CellValue {
  if (value === undefined || isSharedStringReference(value)) {
    return { tag: ValueTag.Empty }
  }
  return value
}

function emptyCellValue(): CellValue {
  return { tag: ValueTag.Empty }
}

function stringCellValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function isNumericComparable(value: CellValue): boolean {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean || value.tag === ValueTag.Empty
}

type KernelStringMode = 'compare' | 'output'

class KernelStringArena {
  private readonly ids = new Map<string, number>()
  private readonly outputValues = new Map<number, string>()

  idFor(value: string, mode: KernelStringMode): number {
    const key = mode === 'compare' ? value.toLocaleLowerCase('en-US') : value
    if (mode === 'compare' && key.length === 0) {
      return 0
    }
    const existing = this.ids.get(key)
    if (existing !== undefined) {
      if (mode === 'output') {
        this.outputValues.set(existing, value)
      }
      return existing
    }
    const id = this.ids.size + 1
    this.ids.set(key, id)
    if (mode === 'output') {
      this.outputValues.set(id, value)
    }
    return id
  }

  outputValue(id: number): string | undefined {
    return this.outputValues.get(id)
  }
}

function appendKernelValueFields(
  value: CellValue,
  stringArena: KernelStringArena,
  mode: KernelStringMode,
  tags: number[],
  numbers: number[],
  stringIds: number[],
  errors: number[],
): void {
  tags.push(value.tag)
  switch (value.tag) {
    case ValueTag.Number:
      numbers.push(value.value)
      stringIds.push(0)
      errors.push(ErrorCode.None)
      break
    case ValueTag.Boolean:
      numbers.push(value.value ? 1 : 0)
      stringIds.push(0)
      errors.push(ErrorCode.None)
      break
    case ValueTag.String:
      numbers.push(0)
      stringIds.push(stringArena.idFor(value.value, mode))
      errors.push(ErrorCode.None)
      break
    case ValueTag.Error:
      numbers.push(0)
      stringIds.push(0)
      errors.push(value.code)
      break
    case ValueTag.Empty:
      numbers.push(0)
      stringIds.push(0)
      errors.push(ErrorCode.None)
      break
  }
}

function cellValueFromKernelFields(
  tag: number,
  number: number,
  stringId: number,
  error: number,
  stringArena: KernelStringArena,
): CellValue | null {
  if (!isValueTag(tag)) {
    return null
  }
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return Number.isFinite(number) ? { tag: ValueTag.Number, value: number } : null
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: number !== 0 }
    case ValueTag.String: {
      const value = stringArena.outputValue(stringId)
      return value === undefined ? null : stringCellValue(value)
    }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: error as ErrorCode }
    default:
      return null
  }
}

function isValueTag(value: number): value is ValueTag {
  return Number.isInteger(value) && value >= 0 && value <= 4
}

function cellValueToPatchValue(value: CellValue): XlsxSourceLiteralPatch['value'] | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return Number.isFinite(value.value) ? value.value : undefined
    case ValueTag.Boolean:
      return value.value
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return undefined
  }
}

function isSharedStringReference(value: PendingCellValue): value is Extract<PendingCellValue, { readonly kind: 'shared-string' }> {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}

function groupCandidates(candidates: readonly CompiledRowChainCandidate[]): Map<string, CompiledRowChainCandidate[]> {
  const output = new Map<string, CompiledRowChainCandidate[]>()
  for (const candidate of candidates) {
    const key =
      candidate.secondKind === 'affine'
        ? `affine:${String(candidate.firstOperatorCode)}:${String(candidate.secondScale)}:${String(candidate.secondOffset)}`
        : `divide:${String(candidate.firstOperatorCode)}`
    const group = output.get(key) ?? []
    group.push(candidate)
    output.set(key, group)
  }
  return output
}

function groupRangeAggregateCandidates(candidates: readonly RangeAggregateCandidate[]): Map<number, RangeAggregateCandidate[]> {
  const output = new Map<number, RangeAggregateCandidate[]>()
  for (const candidate of candidates) {
    const key = candidate.values.length
    const group = output.get(key) ?? []
    group.push(candidate)
    output.set(key, group)
  }
  return output
}

function isAffineRowChainCandidate(candidate: CompiledRowChainCandidate): candidate is AffineRowChainCandidate {
  return candidate.secondKind === 'affine'
}

function isDivideRowChainCandidate(candidate: CompiledRowChainCandidate): candidate is DivideRowChainCandidate {
  return candidate.secondKind === 'divide'
}

function formulaPatch(cell: NativeFormulaCell, value: XlsxSourceLiteralPatch['value']): XlsxSourceLiteralPatch {
  return {
    sheetName: cell.sheetName,
    address: cell.address,
    value,
    preserveFormula: true,
  }
}

function cellKey(cell: NativeFormulaCell): string {
  return `${cell.sheetName}!${cell.address}`
}

function emptyResult(): StreamingNativeWasmRowChainResult {
  return {
    batchCount: 0,
    evaluatedFormulaCellCount: 0,
    patches: [],
    processedCells: new Set(),
  }
}
