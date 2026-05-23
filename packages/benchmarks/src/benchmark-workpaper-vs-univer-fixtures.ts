export function numberColumnSheet(rowCount: number): Array<Array<number | string | null>> {
  return Array.from({ length: rowCount }, (_, row) => [row + 1])
}

export function numberColumnValues(rowCount: number): number[][] {
  return Array.from({ length: rowCount }, (_, row) => [row + 1])
}

export function twoDimensionalValues(rowCount: number, colCount: number): number[][] {
  return Array.from({ length: rowCount }, (__, row) => Array.from({ length: colCount }, (_, col) => (row + 1) * (col + 1)))
}

export function formulaChainFormulas(rowCount: number): string[][] {
  return Array.from({ length: rowCount }, (_, row) => [`=${row === 0 ? 'A1' : `B${String(row)}`}+1`])
}

export function scalarFanoutFormulas(rowCount: number): string[][] {
  return Array.from({ length: rowCount }, (_, row) => [`=A1+${String(row + 1)}`])
}

export function deepChainFormulas(chainLength: number): string[][] {
  return [Array.from({ length: chainLength }, (_, index) => `=${columnName(index)}1+1`)]
}

export function overlappingAggregateFormulas(rowCount: number): string[][] {
  return Array.from({ length: rowCount }, (_, row) => [`=SUM(A1:A${String(row + 1)})`])
}

export function lookupTableValues(rowCount: number): number[][] {
  return Array.from({ length: rowCount }, (_, row) => [row + 1, (row + 1) * 10])
}

export function evenLookupTableValues(rowCount: number): number[][] {
  return Array.from({ length: rowCount }, (_, row) => [(row + 1) * 2, (row + 1) * 10])
}

export function duplicateApproximateLookupTableValues(rowCount: number): number[][] {
  return Array.from({ length: rowCount }, (_, row) => {
    const rowNumber = row + 1
    return [Math.ceil(rowNumber / 2), rowNumber]
  })
}

export function hlookupTableValues(colCount: number): number[][] {
  return [Array.from({ length: colCount }, (_, col) => col + 1), Array.from({ length: colCount }, (_, col) => (col + 1) * 5)]
}

export function rangeStatsValues(rowCount: number): number[][] {
  return Array.from({ length: rowCount }, (_, row) => {
    const rowNumber = row + 1
    return [rowNumber, rowNumber, rowNumber]
  })
}

export function textLookupTableValues(rowCount: number): Array<[string, number]> {
  return Array.from({ length: rowCount }, (_, row) => {
    const rowNumber = row + 1
    return [textLookupKey(rowNumber), rowNumber * 10]
  })
}

export function textLookupKey(rowNumber: number): string {
  return `KEY-${String(rowNumber).padStart(5, '0')}`
}

export function normalizeBenchmarkValue(value: unknown): boolean | number | string | null | { error: string } {
  if (typeof value === 'number') {
    return Number(value.toPrecision(12))
  }
  if (value === null || typeof value === 'boolean' || typeof value === 'string') {
    return value
  }
  if (isErrorRecord(value)) {
    return { error: String(value.error) }
  }
  return { error: 'UNKNOWN_VALUE' }
}

export function columnName(col: number): string {
  let value = col + 1
  let name = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    value = Math.floor((value - 1) / 26)
  }
  return name
}

function isErrorRecord(value: unknown): value is { error: unknown } {
  return typeof value === 'object' && value !== null && 'error' in value
}
