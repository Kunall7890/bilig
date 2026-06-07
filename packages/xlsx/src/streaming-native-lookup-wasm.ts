import { parseFormula, type FormulaNode } from '@bilig/formula'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { createKernelSync } from '@bilig/wasm-kernel'

import { decodeCellAddress } from './address.js'
import type { XlsxSourceLiteralPatch } from './source-preserving-literal-patches.js'
import type { NativeFormulaCell, PendingCellValue, SheetScanState } from './streaming-native-recalc.js'

export interface StreamingNativeWasmLookupResult {
  readonly batchCount: number
  readonly evaluatedFormulaCellCount: number
  readonly patches: readonly XlsxSourceLiteralPatch[]
  readonly processedCells: ReadonlySet<string>
}

interface ExactUniformNumericVlookupCandidate {
  readonly cell: NativeFormulaCell
  readonly lookupValue: number
  readonly tableSheetName: string
  readonly tableStartRow: number
  readonly resultCol: number
  readonly start: number
  readonly step: number
  readonly length: number
  readonly rowValues: Map<number, PendingCellValue>
}

interface VlookupTableRange {
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly width: number
  readonly cellCount: number
}

const directLookupKindExactUniformNumeric = 1
const directLookupMatchModeAscending = 1
const maxNativeLookupTableCellCount = 50_000

export function evaluateStreamingNativeWasmLookups(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): StreamingNativeWasmLookupResult {
  const candidates = collectExactUniformNumericVlookupCandidates(args)
  if (candidates.length === 0) {
    return emptyResult()
  }
  const outTags = new Uint8Array(candidates.length)
  const outNumbers = new Float64Array(candidates.length)
  const outErrors = new Uint16Array(candidates.length)
  createKernelSync().evalUniformNumericLookupBatch(
    Uint8Array.from(candidates.map(() => directLookupKindExactUniformNumeric)),
    Uint8Array.from(candidates.map(() => directLookupMatchModeAscending)),
    Float64Array.from(candidates.map((candidate) => candidate.start)),
    Float64Array.from(candidates.map((candidate) => candidate.step)),
    Uint32Array.from(candidates.map((candidate) => candidate.length)),
    Uint32Array.from(candidates.map(() => 0)),
    Uint8Array.from(candidates.map(() => ValueTag.Number)),
    Float64Array.from(candidates.map((candidate) => candidate.lookupValue)),
    outTags,
    outNumbers,
    outErrors,
  )

  const patches: XlsxSourceLiteralPatch[] = []
  const processedCells = new Set<string>()
  for (let index = 0; index < candidates.length; index += 1) {
    if (outTags[index] !== ValueTag.Number) {
      continue
    }
    const position = outNumbers[index]!
    if (!Number.isInteger(position) || position < 1) {
      continue
    }
    const candidate = candidates[index]!
    const row = candidate.tableStartRow + position - 1
    const value = readScannedCellValue(args.sheetScans, candidate.tableSheetName, row, candidate.resultCol)
    const patchValue = cellValueToPatchValue(value)
    if (patchValue === undefined) {
      continue
    }
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

function collectExactUniformNumericVlookupCandidates(args: {
  readonly sheetScans: ReadonlyMap<string, SheetScanState>
  readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  readonly skippedCells: ReadonlySet<string>
}): ExactUniformNumericVlookupCandidate[] {
  const candidates: ExactUniformNumericVlookupCandidate[] = []
  for (const scan of args.sheetScans.values()) {
    for (const cell of scan.formulaCells) {
      if (args.skippedCells.has(cellKey(cell))) {
        continue
      }
      const rowValues = scan.rows.get(cell.row)
      if (!rowValues) {
        continue
      }
      const candidate = tryCompileExactUniformNumericVlookupCandidate(scan, cell, rowValues, args)
      if (candidate) {
        candidates.push(candidate)
      }
    }
  }
  return candidates
}

function tryCompileExactUniformNumericVlookupCandidate(
  scan: SheetScanState,
  cell: NativeFormulaCell,
  rowValues: Map<number, PendingCellValue>,
  args: {
    readonly sheetScans: ReadonlyMap<string, SheetScanState>
    readonly resolveFormulaSource: (scan: SheetScanState, cell: NativeFormulaCell) => string
  },
): ExactUniformNumericVlookupCandidate | null {
  let node: FormulaNode
  try {
    node = parseFormula(args.resolveFormulaSource(scan, cell))
  } catch {
    return null
  }
  if (
    node.kind !== 'CallExpr' ||
    normalizedFormulaFunctionName(node.callee) !== 'VLOOKUP' ||
    node.args.length < 3 ||
    node.args.length > 4 ||
    !isExactVlookupMatchModeLiteral(node.args[3])
  ) {
    return null
  }
  const tableRange = node.args[1]?.kind === 'RangeRef' ? decodeVlookupTableRange(node.args[1], scan.sheetName) : null
  if (!tableRange || tableRange.cellCount > maxNativeLookupTableCellCount) {
    return null
  }
  const columnIndex = numberLiteralValue(node.args[2]!)
  if (columnIndex === null || !Number.isInteger(columnIndex) || columnIndex < 1 || columnIndex > tableRange.width) {
    return null
  }
  const lookupValue = compileNumericLookupValue(node.args[0]!, scan, cell, rowValues, args.sheetScans)
  if (lookupValue === null) {
    return null
  }
  const lookupShape = exactUniformNumericLookupShape(args.sheetScans, tableRange)
  if (!lookupShape) {
    return null
  }
  return {
    cell,
    lookupValue,
    tableSheetName: tableRange.sheetName,
    tableStartRow: tableRange.startRow,
    resultCol: tableRange.startCol + columnIndex - 1,
    start: lookupShape.start,
    step: lookupShape.step,
    length: lookupShape.length,
    rowValues,
  }
}

function isExactVlookupMatchModeLiteral(node: FormulaNode | undefined): boolean {
  return (node?.kind === 'BooleanLiteral' && !node.value) || (node?.kind === 'NumberLiteral' && node.value === 0)
}

function numberLiteralValue(node: FormulaNode): number | null {
  return node.kind === 'NumberLiteral' && Number.isFinite(node.value) ? node.value : null
}

function compileNumericLookupValue(
  node: FormulaNode,
  scan: SheetScanState,
  cell: NativeFormulaCell,
  rowValues: Map<number, PendingCellValue>,
  sheetScans: ReadonlyMap<string, SheetScanState>,
): number | null {
  if (node.kind === 'NumberLiteral') {
    return Number.isFinite(node.value) ? node.value : null
  }
  if (node.kind !== 'CellRef') {
    return null
  }
  let address
  try {
    address = decodeCellAddress(node.ref.replaceAll('$', ''))
  } catch {
    return null
  }
  const sheetName = node.sheetName ?? scan.sheetName
  const value =
    sheetName === scan.sheetName && address.r === cell.row
      ? resolvedCellValue(rowValues.get(address.c))
      : readScannedCellValue(sheetScans, sheetName, address.r, address.c)
  return value.tag === ValueTag.Number && Number.isFinite(value.value) ? value.value : null
}

function exactUniformNumericLookupShape(
  sheetScans: ReadonlyMap<string, SheetScanState>,
  range: VlookupTableRange,
): { readonly start: number; readonly step: number; readonly length: number } | null {
  const length = range.endRow - range.startRow + 1
  if (length < 2) {
    return null
  }
  const first = readScannedCellValue(sheetScans, range.sheetName, range.startRow, range.startCol)
  const second = readScannedCellValue(sheetScans, range.sheetName, range.startRow + 1, range.startCol)
  if (first.tag !== ValueTag.Number || second.tag !== ValueTag.Number || !Number.isFinite(first.value) || !Number.isFinite(second.value)) {
    return null
  }
  const step = second.value - first.value
  if (step === 0 || !Number.isFinite(step)) {
    return null
  }
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const value = readScannedCellValue(sheetScans, range.sheetName, row, range.startCol)
    if (value.tag !== ValueTag.Number || !Number.isFinite(value.value)) {
      return null
    }
    const offset = row - range.startRow
    if (value.value !== first.value + step * offset) {
      return null
    }
  }
  return { start: first.value, step, length }
}

function decodeVlookupTableRange(
  range: Extract<FormulaNode, { readonly kind: 'RangeRef' }>,
  currentSheetName: string,
): VlookupTableRange | null {
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
    startCol,
    width: endCol - startCol + 1,
    cellCount: (endRow - startRow + 1) * (endCol - startCol + 1),
  }
}

function readScannedCellValue(sheetScans: ReadonlyMap<string, SheetScanState>, sheetName: string, row: number, col: number): CellValue {
  return resolvedCellValue(sheetScans.get(sheetName)?.rows.get(row)?.get(col))
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

function normalizedFormulaFunctionName(name: string): string {
  return name.toLocaleUpperCase('en-US').replace(/^_XLFN\./u, '')
}

function emptyResult(): StreamingNativeWasmLookupResult {
  return {
    batchCount: 0,
    evaluatedFormulaCellCount: 0,
    patches: [],
    processedCells: new Set(),
  }
}
