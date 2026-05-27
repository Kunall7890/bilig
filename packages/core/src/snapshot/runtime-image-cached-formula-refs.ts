import type { CellValue } from '@bilig/protocol'

export interface CachedRuntimeFormulaRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly value: CellValue
  readonly cellIndex: number
}

export function pushRuntimeImageCachedFormulaRef(
  refs: CachedRuntimeFormulaRef[],
  sheetId: number,
  row: number,
  col: number,
  cellIndex: number,
  source: string,
  value: CellValue,
): void {
  refs.push({ sheetId, row, col, cellIndex, source, value })
}
