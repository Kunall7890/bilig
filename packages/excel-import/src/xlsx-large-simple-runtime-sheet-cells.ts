import type { WorkbookSnapshot } from '@bilig/protocol'
import type { ParsedWorksheet } from './xlsx-large-simple-import-types.js'

export function buildLargeSimpleRuntimeSheetCells(
  sheetStats: readonly ParsedWorksheet['stats'][],
  sheets: WorkbookSnapshot['sheets'],
): readonly {
  readonly sheetName: string
  readonly coords: readonly []
  readonly coordinateOrder: 'dense-row-major'
  readonly dimensions: {
    readonly width: number
    readonly height: number
  }
  readonly cellCount: number
}[] {
  return sheetStats.flatMap((entry, index) => {
    const sheet = sheets[index]
    const usedRange = entry.dimension.usedRange
    if (
      !sheet ||
      usedRange === null ||
      usedRange.startRow !== 0 ||
      usedRange.startColumn !== 0 ||
      entry.cellCount !== sheet.cells.length ||
      entry.cellCount !== entry.dimension.rowCount * entry.dimension.columnCount
    ) {
      return []
    }
    return [
      {
        sheetName: sheet.name,
        coords: [],
        coordinateOrder: 'dense-row-major' as const,
        dimensions: {
          width: entry.dimension.columnCount,
          height: entry.dimension.rowCount,
        },
        cellCount: entry.cellCount,
      },
    ]
  })
}
