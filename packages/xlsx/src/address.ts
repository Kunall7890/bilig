export interface XlsxCellAddress {
  readonly r: number
  readonly c: number
}

export interface XlsxCellRange {
  readonly s: XlsxCellAddress
  readonly e: XlsxCellAddress
}

const cellAddressPattern = /^\$?([A-Za-z]{1,4})\$?([1-9][0-9]*)$/u

export function decodeCellAddress(address: string): XlsxCellAddress {
  const match = cellAddressPattern.exec(address)
  if (!match) {
    throw new Error(`Invalid XLSX cell address: ${address}`)
  }
  const rowNumber = Number(match[2])
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 1) {
    throw new Error(`Invalid XLSX row number: ${address}`)
  }
  let column = 0
  for (const character of match[1]!.toUpperCase()) {
    const code = character.charCodeAt(0)
    if (code < 65 || code > 90) {
      throw new Error(`Invalid XLSX column address: ${address}`)
    }
    column = column * 26 + code - 64
  }
  return { r: rowNumber - 1, c: column - 1 }
}

export function encodeColumnAddress(columnIndex: number): string {
  if (!Number.isSafeInteger(columnIndex) || columnIndex < 0) {
    throw new Error(`Invalid XLSX column index: ${String(columnIndex)}`)
  }
  let column = columnIndex + 1
  let output = ''
  while (column > 0) {
    const remainder = (column - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    column = Math.floor((column - 1) / 26)
  }
  return output
}

export function decodeColumnAddress(columnAddress: string): number {
  const normalized = columnAddress.trim().replaceAll('$', '').toUpperCase()
  if (!/^[A-Z]{1,4}$/u.test(normalized)) {
    throw new Error(`Invalid XLSX column address: ${columnAddress}`)
  }
  let column = 0
  for (const character of normalized) {
    const code = character.charCodeAt(0)
    if (code < 65 || code > 90) {
      throw new Error(`Invalid XLSX column address: ${columnAddress}`)
    }
    column = column * 26 + code - 64
  }
  return column - 1
}

export function encodeCellAddress(address: XlsxCellAddress): string {
  if (!Number.isSafeInteger(address.r) || !Number.isSafeInteger(address.c) || address.r < 0 || address.c < 0) {
    throw new Error(`Invalid XLSX cell coordinates: ${JSON.stringify(address)}`)
  }
  return `${encodeColumnAddress(address.c)}${String(address.r + 1)}`
}

export function normalizeCellAddress(address: string): string {
  return encodeCellAddress(decodeCellAddress(address.replaceAll('$', '')))
}

export function decodeCellRange(ref: string): XlsxCellRange {
  const parts = ref.split(':')
  const [startRef, endRef] = parts
  if (!startRef || parts.length > 2) {
    throw new Error(`Invalid XLSX cell range: ${ref}`)
  }
  const start = decodeCellAddress(startRef)
  const end = decodeCellAddress(endRef ?? startRef)
  return {
    s: {
      r: Math.min(start.r, end.r),
      c: Math.min(start.c, end.c),
    },
    e: {
      r: Math.max(start.r, end.r),
      c: Math.max(start.c, end.c),
    },
  }
}

export function encodeCellRange(range: XlsxCellRange): string {
  const start = encodeCellAddress(range.s)
  const end = encodeCellAddress(range.e)
  return start === end ? start : `${start}:${end}`
}
