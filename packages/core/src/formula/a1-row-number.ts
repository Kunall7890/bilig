export function parseA1RowNumber(rowText: string): number | undefined {
  if (rowText.length === 0) {
    return undefined
  }
  const firstCode = rowText.charCodeAt(0)
  if (firstCode < 49 || firstCode > 57) {
    return undefined
  }
  let row = firstCode - 48
  for (let index = 1; index < rowText.length; index += 1) {
    const code = rowText.charCodeAt(index)
    if (code < 48 || code > 57) {
      return undefined
    }
    row = row * 10 + (code - 48)
    if (!Number.isSafeInteger(row)) {
      return undefined
    }
  }
  return row
}

export function parseA1RowIndex(rowText: string): number | undefined {
  const row = parseA1RowNumber(rowText)
  return row === undefined ? undefined : row - 1
}
