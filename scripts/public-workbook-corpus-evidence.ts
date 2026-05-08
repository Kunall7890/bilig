import type { PublicWorkbookCorpusCase } from './public-workbook-corpus-types.ts'

export function publicWorkbookCorpusCaseNeedsEvidenceRefresh(entry: PublicWorkbookCorpusCase): boolean {
  return !hasPublicWorkbookCorpusUsedRangeEvidence(entry)
}

export function hasPublicWorkbookCorpusUsedRangeEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.workbookMetadata.dimensions.every((dimension) => {
    if (!Object.hasOwn(dimension, 'usedRange')) {
      return false
    }
    const range = dimension.usedRange
    if (dimension.nonEmptyCellCount === 0) {
      return range === null
    }
    return (
      range !== null &&
      range !== undefined &&
      range.startRow >= 0 &&
      range.startColumn >= 0 &&
      range.endRow >= range.startRow &&
      range.endColumn >= range.startColumn &&
      dimension.rowCount === range.endRow + 1 &&
      dimension.columnCount === range.endColumn + 1
    )
  })
}
