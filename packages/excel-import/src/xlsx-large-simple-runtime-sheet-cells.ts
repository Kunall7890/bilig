import type { WorkbookSnapshot } from '@bilig/protocol'

export interface LargeSimpleRuntimeSheetCellStat {
  readonly cellCount: number
  readonly dimension: {
    readonly rowCount: number
    readonly columnCount: number
    readonly usedRange: {
      readonly startRow: number
      readonly startColumn: number
    } | null
  }
}

export function buildLargeSimpleRuntimeSheetCells(
  sheetStats: readonly LargeSimpleRuntimeSheetCellStat[],
  sheets: WorkbookSnapshot['sheets'],
): Array<{
  readonly sheetName: string
  readonly coords: []
  readonly coordinateOrder: 'dense-row-major'
  readonly dimensions: { readonly width: number; readonly height: number }
  readonly cellCount: number
}> {
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
        coordinateOrder: 'dense-row-major',
        dimensions: {
          width: entry.dimension.columnCount,
          height: entry.dimension.rowCount,
        },
        cellCount: entry.cellCount,
      },
    ]
  })
}
