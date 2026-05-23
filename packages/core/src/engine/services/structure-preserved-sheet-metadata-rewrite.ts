import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { columnToIndex, formatAddress, rewriteAddressForStructuralTransform, type StructuralAxisTransform } from '@bilig/formula'
import type { WorkbookPreservedSheetMetadataRecord } from '../../workbook-metadata-types.js'
import { hasPreservedSheetMetadata } from '../../workbook-preserved-metadata.js'

const cellReferencePattern = /^\$?([A-Z]+)\$?([1-9]\d*)$/iu

export function rewritePreservedSheetMetadataForStructuralTransform(
  metadata: WorkbookPreservedSheetMetadataRecord | undefined,
  transform: StructuralAxisTransform,
): WorkbookPreservedSheetMetadataRecord | undefined {
  if (!metadata) {
    return undefined
  }
  const next: WorkbookPreservedSheetMetadataRecord = { ...metadata }
  if (metadata.styleArtifacts) {
    const styleArtifacts = rewriteStyleArtifactsForStructuralTransform(metadata.styleArtifacts, transform)
    if (styleArtifacts) {
      next.styleArtifacts = styleArtifacts
    } else {
      delete next.styleArtifacts
    }
  }
  return hasPreservedSheetMetadata(next) ? next : undefined
}

export function preservedSheetMetadataTouchesStructuralDelete(
  metadata: WorkbookPreservedSheetMetadataRecord | undefined,
  axis: 'row' | 'column',
  start: number,
): boolean {
  if (!metadata) {
    return false
  }
  const styleArtifacts = metadata.styleArtifacts
  if (
    styleArtifacts?.cellStyleIndexes.some((entry) => cellReferenceTouchesAxisDelete(entry.address, axis, start)) ||
    styleArtifacts?.blankCellAddresses?.some((address) => cellReferenceTouchesAxisDelete(address, axis, start))
  ) {
    return true
  }
  return false
}

function rewriteStyleArtifactsForStructuralTransform(
  styleArtifacts: NonNullable<WorkbookPreservedSheetMetadataRecord['styleArtifacts']>,
  transform: StructuralAxisTransform,
): WorkbookPreservedSheetMetadataRecord['styleArtifacts'] | undefined {
  const cellStyleIndexes = styleArtifacts.cellStyleIndexes.flatMap((entry) => {
    const address = rewriteCellReferenceForStructuralTransform(entry.address, transform)
    return address ? [{ ...entry, address }] : []
  })
  const blankCellAddresses = (styleArtifacts.blankCellAddresses ?? []).flatMap((address) => {
    const nextAddress = rewriteCellReferenceForStructuralTransform(address, transform)
    return nextAddress ? [nextAddress] : []
  })
  if (cellStyleIndexes.length === 0 && blankCellAddresses.length === 0) {
    return undefined
  }
  return {
    cellStyleIndexes,
    ...(blankCellAddresses.length > 0 ? { blankCellAddresses } : {}),
  }
}

function rewriteCellReferenceForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  if (!isCellReference(address)) {
    return address
  }
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  return rewritten ? clipAddressToSheetGrid(rewritten) : undefined
}

function cellReferenceTouchesAxisDelete(address: string, axis: 'row' | 'column', start: number): boolean {
  if (!isCellReference(address)) {
    return false
  }
  const parsed = parseCellAddress(address)
  if (!parsed) {
    return false
  }
  return (axis === 'row' ? parsed[0] : parsed[1]) >= start
}

function isCellReference(address: string): boolean {
  return cellReferencePattern.test(address)
}

function clipAddressToSheetGrid(address: string): string | undefined {
  const parsed = parseCellAddress(address)
  if (!parsed || parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
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
