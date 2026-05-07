import { formatAddress, parseCellAddress, parseRangeAddress } from '@bilig/formula'
import type { CellSnapshot } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'

export interface StructuralFormulaUndoRecord {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly formula: string
}

export type StructuralDeletedCellUndoRecord =
  | {
      readonly kind: 'formula'
      readonly sheetName: string
      readonly row: number
      readonly col: number
      readonly formula: string
      readonly explicitFormat?: string
    }
  | {
      readonly kind: 'value'
      readonly sheetName: string
      readonly row: number
      readonly col: number
      readonly value: number | boolean
      readonly explicitFormat?: string
    }
  | {
      readonly kind: 'snapshot'
      readonly sheetName: string
      readonly row: number
      readonly col: number
      readonly snapshot: CellSnapshot
    }
  | {
      readonly kind: 'blank'
      readonly sheetName: string
      readonly row: number
      readonly col: number
      readonly restoreExplicitBlank: boolean
      readonly explicitFormat?: string
    }

export function dependencyTouchesStructuralDeleteSpan(
  dependency: string,
  ownerSheetName: string,
  targetSheetName: string,
  axis: 'row' | 'column',
  start: number,
): boolean {
  if (dependency.includes(':')) {
    const parsed = parseRangeAddress(dependency, ownerSheetName)
    const dependencySheetName = parsed.sheetName ?? ownerSheetName
    if (dependencySheetName !== targetSheetName) {
      return false
    }
    if (parsed.kind === 'cells') {
      return (axis === 'row' ? parsed.end.row : parsed.end.col) >= start
    }
    if (parsed.kind === 'rows') {
      return axis === 'row' && parsed.end.row >= start
    }
    return axis === 'column' && parsed.end.col >= start
  }
  const parsed = parseCellAddress(dependency, ownerSheetName)
  const dependencySheetName = parsed.sheetName ?? ownerSheetName
  if (dependencySheetName !== targetSheetName) {
    return false
  }
  return (axis === 'row' ? parsed.row : parsed.col) >= start
}

export function structuralFormulaUndoRecordToOp(record: StructuralFormulaUndoRecord): EngineOp {
  return {
    kind: 'setCellFormula',
    sheetName: record.sheetName,
    address: formatAddress(record.row, record.col),
    formula: record.formula,
  }
}

export function structuralDeletedCellUndoRecordToOps(
  record: StructuralDeletedCellUndoRecord,
  toCellStateOps: (sheetName: string, address: string, snapshot: CellSnapshot) => EngineOp[],
): EngineOp[] {
  const address = formatAddress(record.row, record.col)
  switch (record.kind) {
    case 'formula':
      return [
        { kind: 'setCellFormula', sheetName: record.sheetName, address, formula: record.formula },
        ...(record.explicitFormat === undefined
          ? []
          : [{ kind: 'setCellFormat' as const, sheetName: record.sheetName, address, format: record.explicitFormat }]),
      ]
    case 'value':
      return [
        { kind: 'setCellValue', sheetName: record.sheetName, address, value: record.value },
        ...(record.explicitFormat === undefined
          ? []
          : [{ kind: 'setCellFormat' as const, sheetName: record.sheetName, address, format: record.explicitFormat }]),
      ]
    case 'snapshot':
      return toCellStateOps(record.sheetName, address, record.snapshot)
    case 'blank':
      return [
        record.restoreExplicitBlank
          ? { kind: 'setCellValue', sheetName: record.sheetName, address, value: null }
          : { kind: 'clearCell', sheetName: record.sheetName, address },
        ...(record.explicitFormat === undefined
          ? []
          : [{ kind: 'setCellFormat' as const, sheetName: record.sheetName, address, format: record.explicitFormat }]),
      ]
  }
}
