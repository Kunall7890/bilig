export interface A1CellRef {
  readonly r: number
  readonly c: number
}

export interface A1RangeRef {
  readonly s: A1CellRef
  readonly e: A1CellRef
}

const cellRefPattern = /^\$?([A-Z]+)\$?(\d+)$/iu

export function decodeA1CellRef(address: string): A1CellRef {
  const normalized = address.trim().replaceAll('$', '')
  const match = cellRefPattern.exec(normalized)
  if (!match) {
    throw new Error(`Invalid A1 cell address: ${address}`)
  }
  const row = Number(match[2])
  if (!Number.isSafeInteger(row) || row <= 0) {
    throw new Error(`Invalid A1 cell address: ${address}`)
  }
  return { r: row - 1, c: decodeA1ColumnRef(match[1]!) }
}

export function encodeA1CellRef(cell: A1CellRef): string {
  if (!Number.isSafeInteger(cell.r) || !Number.isSafeInteger(cell.c) || cell.r < 0 || cell.c < 0) {
    throw new Error(`Invalid A1 cell coordinate: ${String(cell.r)},${String(cell.c)}`)
  }
  return `${encodeA1ColumnRef(cell.c)}${String(cell.r + 1)}`
}

export function decodeA1RangeRef(ref: string): A1RangeRef {
  const [start, end] = ref.replaceAll('$', '').split(':')
  if (!start) {
    throw new Error(`Invalid A1 range: ${ref}`)
  }
  const startCell = decodeA1CellRef(start)
  const endCell = decodeA1CellRef(end ?? start)
  return {
    s: {
      r: Math.min(startCell.r, endCell.r),
      c: Math.min(startCell.c, endCell.c),
    },
    e: {
      r: Math.max(startCell.r, endCell.r),
      c: Math.max(startCell.c, endCell.c),
    },
  }
}

export function encodeA1RangeRef(range: A1RangeRef): string {
  const start = encodeA1CellRef(range.s)
  const end = encodeA1CellRef(range.e)
  return start === end ? start : `${start}:${end}`
}

export function decodeA1ColumnRef(column: string): number {
  let value = 0
  for (const char of column.toUpperCase()) {
    const code = char.charCodeAt(0)
    if (code < 65 || code > 90) {
      throw new Error(`Invalid A1 column: ${column}`)
    }
    value = value * 26 + code - 64
  }
  if (value <= 0) {
    throw new Error(`Invalid A1 column: ${column}`)
  }
  return value - 1
}

export function encodeA1ColumnRef(column: number): string {
  if (!Number.isSafeInteger(column) || column < 0) {
    throw new Error(`Invalid A1 column index: ${String(column)}`)
  }
  let value = column + 1
  let output = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    value = Math.floor((value - 1) / 26)
  }
  return output
}
