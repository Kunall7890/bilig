import type { LiteralInput } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'

export const CapturedCellMutationKind = {
  Clear: 0,
  NumberValue: 1,
  BooleanValue: 2,
  LiteralValue: 3,
  NullValue: 4,
  Formula: 5,
} as const

export interface CapturedCellMutationRestores {
  readonly sheetIds: Uint32Array
  readonly cellIndexPlusOnes: Uint32Array
  readonly rows: Uint32Array
  readonly cols: Uint32Array
  readonly kinds: Uint8Array
  readonly numbers: Float64Array
  readonly values?: Array<LiteralInput | undefined>
  readonly formulas?: Array<string | undefined>
  readonly potentialNewCells: number
}

export function materializeCapturedCellMutationRestores(captured: CapturedCellMutationRestores): EngineCellMutationRef[] {
  const refs = Array<EngineCellMutationRef>(captured.sheetIds.length)
  for (let index = 0; index < refs.length; index += 1) {
    const sheetId = captured.sheetIds[index]!
    const cellIndexPlusOne = captured.cellIndexPlusOnes[index]!
    const cellIndex = cellIndexPlusOne === 0 ? undefined : cellIndexPlusOne - 1
    const row = captured.rows[index]!
    const col = captured.cols[index]!
    switch (captured.kinds[index]!) {
      case CapturedCellMutationKind.NumberValue:
        refs[index] = {
          sheetId,
          ...(cellIndex === undefined ? {} : { cellIndex }),
          mutation: { kind: 'setCellValue', row, col, value: captured.numbers[index] ?? 0 },
        }
        break
      case CapturedCellMutationKind.BooleanValue:
        refs[index] = {
          sheetId,
          ...(cellIndex === undefined ? {} : { cellIndex }),
          mutation: { kind: 'setCellValue', row, col, value: (captured.numbers[index] ?? 0) !== 0 },
        }
        break
      case CapturedCellMutationKind.NullValue:
        refs[index] = {
          sheetId,
          ...(cellIndex === undefined ? {} : { cellIndex }),
          mutation: { kind: 'setCellValue', row, col, value: null },
        }
        break
      case CapturedCellMutationKind.LiteralValue:
        refs[index] = {
          sheetId,
          ...(cellIndex === undefined ? {} : { cellIndex }),
          mutation: { kind: 'setCellValue', row, col, value: captured.values?.[index] ?? null },
        }
        break
      case CapturedCellMutationKind.Formula:
        refs[index] = {
          sheetId,
          ...(cellIndex === undefined ? {} : { cellIndex }),
          mutation: { kind: 'setCellFormula', row, col, formula: captured.formulas?.[index] ?? '' },
        }
        break
      default:
        refs[index] = {
          sheetId,
          ...(cellIndex === undefined ? {} : { cellIndex }),
          mutation: { kind: 'clearCell', row, col },
        }
        break
    }
  }
  return refs
}
