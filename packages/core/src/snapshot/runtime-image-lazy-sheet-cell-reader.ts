import type { WorkbookSnapshot } from '@bilig/protocol'

type WorkbookSnapshotCell = WorkbookSnapshot['sheets'][number]['cells'][number]
type WorkbookSnapshotCells = WorkbookSnapshot['sheets'][number]['cells']

const lazySheetCellReaderSymbol = Symbol.for('bilig.lazyImportedXlsxCells.readCellWithCoordinates')

export interface LazySheetCellReaderResult {
  readonly cell: WorkbookSnapshotCell
  readonly row: number
  readonly col: number
}

export type LazySheetCellReader = (index: number) => LazySheetCellReaderResult | undefined

export function readLazySheetCellReader(cells: WorkbookSnapshotCells): LazySheetCellReader | undefined {
  const reader = Reflect.get(cells, lazySheetCellReaderSymbol)
  return isLazySheetCellReader(reader) ? reader : undefined
}

function isLazySheetCellReader(value: unknown): value is LazySheetCellReader {
  return typeof value === 'function'
}
