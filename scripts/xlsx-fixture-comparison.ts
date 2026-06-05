import { readXlsxZipEntries } from '@bilig/xlsx'

export function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  return Buffer.from(left).equals(Buffer.from(right))
}

export function xlsxZipEntryContentsEqual(leftBytes: Uint8Array, rightBytes: Uint8Array): boolean {
  let left: Record<string, Uint8Array>
  let right: Record<string, Uint8Array>
  try {
    left = readXlsxZipEntries(leftBytes)
    right = readXlsxZipEntries(rightBytes)
  } catch {
    return false
  }

  const leftPaths = Object.keys(left).toSorted()
  const rightPaths = Object.keys(right).toSorted()
  if (leftPaths.length !== rightPaths.length) {
    return false
  }

  for (let index = 0; index < leftPaths.length; index += 1) {
    const path = leftPaths[index]
    if (path !== rightPaths[index]) {
      return false
    }
    if (!bytesEqual(left[path], right[path])) {
      return false
    }
  }
  return true
}
