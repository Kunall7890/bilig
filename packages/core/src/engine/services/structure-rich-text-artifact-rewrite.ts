import { MAX_COLS, MAX_ROWS, type WorkbookSheetRichTextArtifactsSnapshot } from '@bilig/protocol'
import { columnToIndex, formatAddress, rewriteAddressForStructuralTransform, type StructuralAxisTransform } from '@bilig/formula'

const cellReferencePattern = /^\$?([A-Z]+)\$?([1-9]\d*)$/iu

export function rewriteRichTextArtifactsForStructuralTransform(
  richTextArtifacts: WorkbookSheetRichTextArtifactsSnapshot | undefined,
  transform: StructuralAxisTransform,
): WorkbookSheetRichTextArtifactsSnapshot | undefined {
  if (!richTextArtifacts) {
    return undefined
  }

  let changed = false
  const cells = richTextArtifacts.cells.flatMap((cell) => {
    const rewrittenAddress = rewriteRichTextAddress(cell.address, transform)
    if (!rewrittenAddress) {
      changed = true
      return []
    }
    if (rewrittenAddress !== cell.address) {
      changed = true
      return [{ ...cell, address: rewrittenAddress }]
    }
    return [cell]
  })

  if (cells.length === 0) {
    return undefined
  }
  return changed ? { cells } : richTextArtifacts
}

function rewriteRichTextAddress(address: string, transform: StructuralAxisTransform): string | undefined {
  if (!cellReferencePattern.test(address)) {
    return address
  }
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  return rewritten ? clipAddressToSheetGrid(rewritten) : undefined
}

function clipAddressToSheetGrid(address: string): string | undefined {
  const parsed = parseCellAddress(address)
  if (!parsed) {
    return undefined
  }
  if (parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function parseCellAddress(address: string): [number, number] | undefined {
  const match = cellReferencePattern.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
