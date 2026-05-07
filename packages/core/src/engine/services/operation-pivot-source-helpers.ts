import { parseCellAddress } from '@bilig/formula'
import type { WorkbookPivotRecord } from '../../workbook-store.js'

export interface OperationPivotSourceWorkbook {
  readonly hasPivots: () => boolean
  readonly listPivots: () => readonly Pick<WorkbookPivotRecord, 'source'>[]
}

export function cellTouchesOperationPivotSource(input: {
  readonly workbook: OperationPivotSourceWorkbook
  readonly sheetName: string
  readonly row: number
  readonly col: number
}): boolean {
  if (!input.workbook.hasPivots()) {
    return false
  }
  return input.workbook.listPivots().some((pivot) => {
    if (pivot.source.sheetName !== input.sheetName) {
      return false
    }
    const start = parseCellAddress(pivot.source.startAddress, pivot.source.sheetName)
    const end = parseCellAddress(pivot.source.endAddress, pivot.source.sheetName)
    return input.row >= start.row && input.row <= end.row && input.col >= start.col && input.col <= end.col
  })
}
