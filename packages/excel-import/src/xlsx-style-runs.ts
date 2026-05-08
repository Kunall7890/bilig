import * as XLSX from 'xlsx'

import type { SheetStyleRangeSnapshot } from '@bilig/protocol'

export interface HorizontalStyleRun {
  styleId: string
  startColumn: number
  endColumn: number
}

export interface RectangularStyleRun extends HorizontalStyleRun {
  startRow: number
  endRow: number
}

function styleRunKey(run: Pick<RectangularStyleRun, 'styleId' | 'startColumn' | 'endColumn'>): string {
  return `${run.styleId}:${String(run.startColumn)}:${String(run.endColumn)}`
}

export function mergeStyleRuns(
  rowIndex: number,
  rowRuns: readonly HorizontalStyleRun[],
  openRunsByKey: ReadonlyMap<string, RectangularStyleRun>,
  styleRuns: RectangularStyleRun[],
): Map<string, RectangularStyleRun> {
  const nextOpenRunsByKey = new Map<string, RectangularStyleRun>()
  for (const rowRun of rowRuns) {
    const key = styleRunKey(rowRun)
    const openRun = openRunsByKey.get(key)
    if (openRun && openRun.endRow === rowIndex - 1) {
      openRun.endRow = rowIndex
      nextOpenRunsByKey.set(key, openRun)
      continue
    }
    const nextRun: RectangularStyleRun = {
      ...rowRun,
      startRow: rowIndex,
      endRow: rowIndex,
    }
    styleRuns.push(nextRun)
    nextOpenRunsByKey.set(key, nextRun)
  }
  return nextOpenRunsByKey
}

export function styleRunsToRanges(sheetName: string, styleRuns: readonly RectangularStyleRun[]): SheetStyleRangeSnapshot[] {
  return styleRuns.map((run) => ({
    range: {
      sheetName,
      startAddress: XLSX.utils.encode_cell({ r: run.startRow, c: run.startColumn }),
      endAddress: XLSX.utils.encode_cell({ r: run.endRow, c: run.endColumn }),
    },
    styleId: run.styleId,
  }))
}
