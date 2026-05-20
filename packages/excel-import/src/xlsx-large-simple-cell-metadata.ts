import type { WorkbookCellMetadataReferenceSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import {
  buildImportedCellMetadataReferenceSnapshots,
  buildImportedCellMetadataReferenceSnapshotsFromCellMap,
} from './xlsx-cell-metadata.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import type { LargeSimpleWorksheetCellMetadataRef } from './xlsx-large-simple-worksheet-metadata.js'

type WorkbookSheetCells = WorkbookSnapshot['sheets'][number]['cells']

export function buildLargeSimpleCellMetadataReferenceSnapshots(
  refs: readonly LargeSimpleWorksheetCellMetadataRef[] | undefined,
  cells: WorkbookSheetCells,
  cellScan: Pick<ImportedWorksheetCellScan, 'arena' | 'sheetIndex'>,
  useLazyCells: boolean,
): WorkbookCellMetadataReferenceSnapshot[] | undefined {
  if (!refs || refs.length === 0) {
    return undefined
  }
  if (!useLazyCells) {
    return buildImportedCellMetadataReferenceSnapshots(refs, cells)
  }
  return buildImportedCellMetadataReferenceSnapshotsFromCellMap(
    refs,
    cellScan.arena.materializeSheetCellsByAddress(cellScan.sheetIndex, new Set(refs.map((ref) => ref.address))),
  )
}
