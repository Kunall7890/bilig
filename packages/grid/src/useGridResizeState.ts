import type { HeaderSelection } from './gridPointer.js'

export function resolveResizeGuideColumn(input: {
  readonly activeResizeColumn: number | null
  readonly cursor: string
  readonly header: HeaderSelection | null
}): number | null {
  void input.cursor
  void input.header
  return input.activeResizeColumn
}

export function resolveResizeGuideRow(input: {
  readonly activeResizeRow: number | null
  readonly cursor: string
  readonly header: HeaderSelection | null
}): number | null {
  void input.cursor
  void input.header
  return input.activeResizeRow
}
