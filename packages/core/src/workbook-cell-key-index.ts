import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'

const SHEET_STRIDE = MAX_ROWS * MAX_COLS

type CellKeyResolver = (sheetId: number, row: number, col: number) => number | undefined

export function makeCellKey(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_STRIDE + row * MAX_COLS + col
}

export function makeLogicalCellKey(sheetId: number, rowId: string, colId: string): string {
  return `${sheetId}\t${rowId}\t${colId}`
}

export function createCellKeyIndexMap(resolve: CellKeyResolver): Map<number, number> {
  return new LogicalCellKeyIndexMap(resolve)
}

function decodeCellKey(key: number): { sheetId: number; row: number; col: number } | undefined {
  if (!Number.isInteger(key) || key < 0) {
    return undefined
  }
  const sheetId = Math.floor(key / SHEET_STRIDE)
  const offset = key - sheetId * SHEET_STRIDE
  const row = Math.floor(offset / MAX_COLS)
  const col = offset - row * MAX_COLS
  if (sheetId <= 0 || row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) {
    return undefined
  }
  return { sheetId, row, col }
}

class LogicalCellKeyIndexMap extends Map<number, number> {
  constructor(private readonly resolve: CellKeyResolver) {
    super()
  }

  override get(key: number): number | undefined {
    const decoded = decodeCellKey(key)
    if (decoded) {
      return this.resolve(decoded.sheetId, decoded.row, decoded.col)
    }
    return super.get(key)
  }

  override has(key: number): boolean {
    return this.get(key) !== undefined
  }
}
