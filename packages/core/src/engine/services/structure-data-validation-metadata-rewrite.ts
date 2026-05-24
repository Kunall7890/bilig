import { MAX_COLS, MAX_ROWS, type CellRangeRef, type WorkbookDataValidationSnapshot } from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'

type WorkbookDataValidationListSource = NonNullable<Extract<WorkbookDataValidationSnapshot['rule'], { kind: 'list' }>['source']>

type DataValidationSourceRewrite =
  | { readonly kind: 'unchanged' }
  | { readonly kind: 'invalid' }
  | { readonly kind: 'rewritten'; readonly validation: WorkbookDataValidationSnapshot }

const DATA_VALIDATION_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i

export function rewriteDataValidationSourceForStructuralTransform(
  validation: WorkbookDataValidationSnapshot,
  sheetName: string,
  transform: StructuralAxisTransform,
): DataValidationSourceRewrite {
  if (validation.rule.kind !== 'list' || !validation.rule.source) {
    return { kind: 'unchanged' }
  }
  const nextSource = rewriteDataValidationListSourceForStructuralTransform(
    validation.rule.source,
    validation.range.sheetName,
    sheetName,
    transform,
  )
  if (!nextSource) {
    return { kind: 'invalid' }
  }
  if (nextSource === validation.rule.source) {
    return { kind: 'unchanged' }
  }
  return {
    kind: 'rewritten',
    validation: {
      ...structuredClone(validation),
      rule: {
        ...validation.rule,
        source: nextSource,
      },
    },
  }
}

function rewriteDataValidationListSourceForStructuralTransform(
  source: WorkbookDataValidationListSource,
  ownerSheetName: string,
  sheetName: string,
  transform: StructuralAxisTransform,
): WorkbookDataValidationListSource | undefined {
  switch (source.kind) {
    case 'cell-ref': {
      if (source.sheetName !== sheetName) {
        return source
      }
      const nextAddress = rewriteDataValidationAddressForStructuralTransform(source.address, transform)
      return nextAddress ? { ...source, address: nextAddress } : undefined
    }
    case 'range-ref': {
      if (source.sheetName !== sheetName) {
        return source
      }
      return rewriteDataValidationRangeForStructuralTransform(source, transform)
    }
    case 'named-range':
    case 'structured-ref':
      return source
    case 'formula': {
      const nextFormula = rewriteDataValidationFormulaSourceForStructuralTransform(source.formula, ownerSheetName, sheetName, transform)
      return nextFormula === undefined ? source : { kind: 'formula', formula: nextFormula }
    }
  }
}

function rewriteDataValidationFormulaSourceForStructuralTransform(
  formula: string,
  ownerSheetName: string,
  targetSheetName: string,
  transform: StructuralAxisTransform,
): string | undefined {
  const hasFormulaPrefix = formula.startsWith('=')
  const source = hasFormulaPrefix ? formula.slice(1) : formula
  try {
    const rewritten = rewriteFormulaForStructuralTransform(source, ownerSheetName, targetSheetName, transform)
    const nextFormula = hasFormulaPrefix ? `=${rewritten}` : rewritten
    return nextFormula === formula ? undefined : nextFormula
  } catch {
    return undefined
  }
}

function rewriteDataValidationRangeForStructuralTransform<T extends CellRangeRef>(
  range: T,
  transform: StructuralAxisTransform,
): T | undefined {
  const rewritten = rewriteRangeForStructuralTransform(range.startAddress, range.endAddress, transform)
  if (!rewritten) {
    return undefined
  }
  const clipped = clipDataValidationRangeToSheetGrid(range.sheetName, rewritten.startAddress, rewritten.endAddress)
  return clipped ? { ...range, startAddress: clipped.startAddress, endAddress: clipped.endAddress } : undefined
}

function rewriteDataValidationAddressForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseDataValidationCellAddress(rewritten)
  if (!parsed) {
    throw new Error('Invalid data validation reference')
  }
  if (parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function clipDataValidationRangeToSheetGrid(sheetName: string, startAddress: string, endAddress: string): CellRangeRef | undefined {
  const start = parseDataValidationCellAddress(startAddress)
  const end = parseDataValidationCellAddress(endAddress)
  if (!start || !end) {
    throw new Error('Invalid data validation reference')
  }
  const startRow = Math.min(start[0], end[0])
  const endRow = Math.min(MAX_ROWS - 1, Math.max(start[0], end[0]))
  const startCol = Math.min(start[1], end[1])
  const endCol = Math.min(MAX_COLS - 1, Math.max(start[1], end[1]))
  if (startRow > endRow || startCol > endCol) {
    return undefined
  }
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  }
}

function parseDataValidationCellAddress(address: string): [number, number] | undefined {
  const match = DATA_VALIDATION_CELL_REF_RE.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
