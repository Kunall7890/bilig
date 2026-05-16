export function parseA1RowNumber(rowText: string): number | undefined {
  if (!/^(?:[1-9]\d*)$/u.test(rowText)) {
    return undefined
  }
  const row = Number(rowText)
  return Number.isSafeInteger(row) ? row : undefined
}

export function parseA1RowIndex(rowText: string): number | undefined {
  const row = parseA1RowNumber(rowText)
  return row === undefined ? undefined : row - 1
}
