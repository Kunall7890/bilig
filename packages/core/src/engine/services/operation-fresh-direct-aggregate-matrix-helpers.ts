import { ErrorCode } from '@bilig/protocol'
import { columnToIndex, type CompiledFormula } from '@bilig/formula'
import type { SheetRecord } from '../../workbook-store.js'
import type { RuntimeDirectAggregateDescriptor } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import { translateSimpleDirectAggregateFormula } from '../../formula/simple-direct-aggregate-compile.js'

const SIMPLE_ROW_DIRECT_AGGREGATE_SOURCE_RE =
  /^=?(SUM|AVERAGE|AVG|COUNT|MIN|MAX)\s*\(\s*([A-Za-z]+)([1-9]\d*):([A-Za-z]+)([1-9]\d*)\s*\)(?:\s*\+\s*([+-]?(?:\d+|\d*\.\d+)))?\s*$/i

export interface FreshMatrixDirectAggregateTemplate {
  readonly aggregateKind: RuntimeDirectAggregateDescriptor['aggregateKind']
  readonly compiled: CompiledFormula
  readonly formulaCol: number
  readonly rangeColEnd: number
  readonly rangeColStart: number
  readonly resultOffset: number | undefined
  readonly row: number
  readonly templateId: number
}

type FreshFormulaCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

export function createFreshMatrixDirectAggregateTemplate(input: {
  readonly aggregate: NonNullable<CompiledFormula['directAggregateCandidate']>
  readonly compiled: CompiledFormula
  readonly formulaCol: number
  readonly range: NonNullable<CompiledFormula['parsedSymbolicRanges']>[number]
  readonly row: number
  readonly templateId: number
}): FreshMatrixDirectAggregateTemplate | undefined {
  if (
    input.range.refKind !== 'cells' ||
    input.range.sheetName !== undefined ||
    input.range.startRow !== input.row ||
    input.range.endRow !== input.row ||
    input.range.startCol > input.range.endCol ||
    input.range.endCol >= input.formulaCol
  ) {
    return undefined
  }
  return {
    aggregateKind: input.aggregate.aggregateKind,
    compiled: input.compiled,
    formulaCol: input.formulaCol,
    rangeColEnd: input.range.endCol,
    rangeColStart: input.range.startCol,
    resultOffset: normalizeFreshMatrixDirectAggregateOffset(input.aggregate.resultOffset),
    row: input.row,
    templateId: input.templateId,
  }
}

export function tryTranslateFreshMatrixDirectAggregateTemplate(
  template: FreshMatrixDirectAggregateTemplate,
  source: string,
  row: number,
  col: number,
): CompiledFormula | undefined {
  if (col !== template.formulaCol || !freshMatrixDirectAggregateSourceMatchesTemplate(template, source, row)) {
    return undefined
  }
  return translateSimpleDirectAggregateFormula(template.compiled, row - template.row, 0, source)
}

export function normalizeFreshMatrixDirectAggregateOffset(offset: number | undefined): number | undefined {
  return offset === undefined || offset === 0 ? undefined : offset
}

export function createFreshFormulaCellAttacher(sheet: SheetRecord): FreshFormulaCellAttacher {
  const attachFreshVisibleCellIdentity = sheet.logical.setFreshVisibleCellIdentityWithAxisIdsDeferred.bind(sheet.logical)
  sheet.logical.deferVisibleCellPageRebuild()
  const setGridCell = sheet.grid.createRowMajorSetter()
  return (row, col, cellIndex, rowId, colId) => {
    attachFreshVisibleCellIdentity(cellIndex, rowId, colId)
    setGridCell(row, col, cellIndex)
  }
}

export function materializeFreshMatrixAxisIds(count: number, start: number, ensureAxisId: (index: number) => string): string[] {
  const axisIds: string[] = []
  for (let offset = 0; offset < count; offset += 1) {
    axisIds[offset] = ensureAxisId(start + offset)
  }
  return axisIds
}

export function attachFreshDenseDirectAggregateMatrixCells(
  sheet: SheetRecord,
  firstCellIndex: number,
  rowStart: number,
  colStart: number,
  rowIds: readonly string[],
  colIds: readonly string[],
): void {
  sheet.logical.deferVisibleCellPageRebuild()
  sheet.logical.setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred(firstCellIndex, rowIds, colIds)
  sheet.grid.setDenseRowMajor(rowStart, colStart, rowIds.length, colIds.length, firstCellIndex)
}

export function evaluateFreshDirectAggregateMatrixRow(input: {
  readonly aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
  readonly colEnd: number
  readonly colStart: number
  readonly inputColCount: number
  readonly matrixColStart: number
  readonly resultOffset: number | undefined
  readonly rowOffset: number
  readonly values: Float64Array
}): DirectScalarCurrentOperand {
  let sum = 0
  let count = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  const rowBase = input.rowOffset * input.inputColCount
  for (let col = input.colStart; col <= input.colEnd; col += 1) {
    const value = input.values[rowBase + col - input.matrixColStart]!
    sum += value
    count += 1
    minimum = Math.min(minimum, value)
    maximum = Math.max(maximum, value)
  }
  const result =
    input.aggregateKind === 'sum'
      ? sum
      : input.aggregateKind === 'count'
        ? count
        : input.aggregateKind === 'average'
          ? count === 0
            ? undefined
            : sum / count
          : input.aggregateKind === 'min'
            ? minimum === Number.POSITIVE_INFINITY
              ? 0
              : minimum
            : maximum === Number.NEGATIVE_INFINITY
              ? 0
              : maximum
  if (result === undefined) {
    return { kind: 'error', code: ErrorCode.Div0 }
  }
  return { kind: 'number', value: result + (input.resultOffset ?? 0) }
}

function freshMatrixDirectAggregateSourceMatchesTemplate(
  template: FreshMatrixDirectAggregateTemplate,
  source: string,
  row: number,
): boolean {
  const match = SIMPLE_ROW_DIRECT_AGGREGATE_SOURCE_RE.exec(source.trim())
  if (!match) {
    return false
  }
  const aggregateKind = directAggregateKindFromSourceCallee(match[1]!)
  if (aggregateKind !== template.aggregateKind) {
    return false
  }
  const startCol = columnToIndex(match[2]!.toUpperCase())
  const startRow = parseFreshMatrixA1RowIndex(match[3]!)
  const endCol = columnToIndex(match[4]!.toUpperCase())
  const endRow = parseFreshMatrixA1RowIndex(match[5]!)
  if (startCol !== template.rangeColStart || endCol !== template.rangeColEnd || startRow !== row || endRow !== row || endCol < startCol) {
    return false
  }
  const resultOffset = match[6] === undefined ? undefined : normalizeFreshMatrixDirectAggregateOffset(Number(match[6]))
  return (resultOffset ?? 0) === (template.resultOffset ?? 0)
}

function directAggregateKindFromSourceCallee(callee: string): RuntimeDirectAggregateDescriptor['aggregateKind'] | undefined {
  switch (callee.toUpperCase()) {
    case 'SUM':
      return 'sum'
    case 'AVERAGE':
    case 'AVG':
      return 'average'
    case 'COUNT':
      return 'count'
    case 'MIN':
      return 'min'
    case 'MAX':
      return 'max'
    default:
      return undefined
  }
}

function parseFreshMatrixA1RowIndex(source: string): number {
  if (!/^[1-9]\d*$/.test(source)) {
    return -1
  }
  const rowNumber = Number(source)
  return Number.isSafeInteger(rowNumber) ? rowNumber - 1 : -1
}
