import type { CellStyleRecord, SheetMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { createSheetPreview } from './workbook-import-helpers.js'
import {
  buildLargeSimpleCellMetadataReferenceSnapshots,
  buildLargeSimpleLazyCellMetadataReferenceSnapshots,
} from './xlsx-large-simple-cell-metadata.js'
import { internLargeSimpleSheetMetadataInput } from './xlsx-large-simple-metadata-interning.js'
import {
  normalizeLargeSimpleConditionalFormatIds,
  readLargeSimpleConditionalFormattingBlockCount,
} from './xlsx-large-simple-conditional-format-helpers.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import { buildLargeSimpleStyleRanges } from './xlsx-large-simple-style-ranges.js'
import { applyLargeSimpleNumberFormatsToCells } from './xlsx-large-simple-number-formats.js'
import { applyImportedAutoFilterVisibility } from './xlsx-filter-visibility.js'
import { releaseProjectedCellScanStorage } from './xlsx-large-simple-materialization-helpers.js'
import {
  readLargeSimpleColumnMetadata,
  readLargeSimpleMergeRanges,
  readLargeSimpleRowMetadata,
  readLargeSimpleSheetFormatPr,
  type LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'
import type { LargeSimpleSheetMetadataInput, ParsedWorksheet } from './xlsx-large-simple-import-types.js'

export const lazySheetCellMaterializationThreshold = 65_536
export const lazySheetCellMaterializationNumberFormatThreshold = 100_000

export function buildParsedWorksheet(
  sheetName: string,
  order: number,
  cellScan: ImportedWorksheetCellScan,
  worksheetXml: string | undefined,
  metadataScan: LargeSimpleWorksheetScannedMetadata | undefined,
  input: LargeSimpleSheetMetadataInput = {},
  options: {
    readonly materializeCells: boolean
    readonly releaseArenaAfterMaterialization?: boolean
    readonly numberFormatsByStyleIndex?: ReadonlyMap<number, string>
    readonly styleCatalog?: Map<string, CellStyleRecord>
    readonly stylesByIndex?: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>
    readonly stringPool?: ImportedWorkbookStringPool
  } = { materializeCells: true },
): ParsedWorksheet {
  const internedInput = internLargeSimpleSheetMetadataInput(input, options.stringPool)
  const merges =
    metadataScan?.merges?.map((range) => ({ sheetName, ...range })) ??
    (worksheetXml ? readLargeSimpleMergeRanges(sheetName, worksheetXml) : [])
  const mergeCount = worksheetXml ? merges.length : (cellScan.mergeCount ?? 0)
  const columns = metadataScan?.columns ?? (worksheetXml ? readLargeSimpleColumnMetadata(worksheetXml) : { entries: [], metadata: [] })
  const rows = metadataScan?.rows ?? (worksheetXml ? readLargeSimpleRowMetadata(worksheetXml) : { entries: [], metadata: [] })
  const sheetFormatPr = metadataScan?.sheetFormatPr ?? (worksheetXml ? readLargeSimpleSheetFormatPr(worksheetXml) : undefined)
  const conditionalFormatCount =
    internedInput.conditionalFormats?.length ??
    (worksheetXml ? readLargeSimpleConditionalFormattingBlockCount(worksheetXml) : (cellScan.conditionalFormatCount ?? 0))
  const conditionalFormats = normalizeLargeSimpleConditionalFormatIds(sheetName, internedInput.conditionalFormats)
  const dataValidationCount = internedInput.validations?.length ?? cellScan.dataValidationCount ?? 0
  const styleRanges =
    options.materializeCells && options.styleCatalog && options.stylesByIndex
      ? buildLargeSimpleStyleRanges(sheetName, cellScan, options.stylesByIndex, options.styleCatalog)
      : []
  const useLazyCells =
    options.materializeCells &&
    cellScan.cellCount >
      (options.numberFormatsByStyleIndex && options.numberFormatsByStyleIndex.size > 0
        ? lazySheetCellMaterializationNumberFormatThreshold
        : lazySheetCellMaterializationThreshold)
  const detachLazyCells = useLazyCells && options.releaseArenaAfterMaterialization !== false
  const preview = createSheetPreview({
    name: sheetName,
    rowCount: cellScan.rowCount,
    columnCount: cellScan.columnCount,
    nonEmptyCellCount: cellScan.cellCount,
    readCellText: (row, column) => cellScan.arena.readPreviewText(row, column),
  })
  const detachedLazyCellMetadataRefs = detachLazyCells
    ? buildLargeSimpleLazyCellMetadataReferenceSnapshots(metadataScan?.cellMetadataRefs, cellScan)
    : undefined
  const cells = options.materializeCells
    ? useLazyCells
      ? detachLazyCells
        ? cellScan.arena.createDetachedLazySheetCells(cellScan.sheetIndex)
        : cellScan.arena.createLazySheetCells(cellScan.sheetIndex)
      : cellScan.arena.materializeSheetCells(cellScan.sheetIndex)
    : []
  const arenaReleasedAfterCellProjection =
    options.materializeCells && options.releaseArenaAfterMaterialization === true && (!useLazyCells || detachLazyCells)
  if (arenaReleasedAfterCellProjection) {
    cellScan.arena.release()
  }
  if (!useLazyCells && options.numberFormatsByStyleIndex && options.numberFormatsByStyleIndex.size > 0) {
    applyLargeSimpleNumberFormatsToCells(cells, cellScan, options.numberFormatsByStyleIndex)
  }
  const cellMetadataRefs =
    detachedLazyCellMetadataRefs ??
    buildLargeSimpleCellMetadataReferenceSnapshots(metadataScan?.cellMetadataRefs, cells, cellScan, useLazyCells)
  releaseProjectedCellScanStorage(cellScan, {
    releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization,
    arenaReleased: arenaReleasedAfterCellProjection,
    detachLazyCells,
    useLazyCells,
  })
  const visibleRows = applyImportedAutoFilterVisibility(sheetName, cells, rows.entries, internedInput.filters)
  const metadata: SheetMetadataSnapshot = {
    ...(columns.entries.length > 0 ? { columns: columns.entries } : {}),
    ...(visibleRows && visibleRows.length > 0 ? { rows: visibleRows } : {}),
    ...(columns.metadata.length > 0 ? { columnMetadata: columns.metadata } : {}),
    ...(rows.metadata.length > 0 ? { rowMetadata: rows.metadata } : {}),
    ...(sheetFormatPr ? { sheetFormatPr } : {}),
    ...(styleRanges.length > 0 ? { styleRanges } : {}),
    ...(merges.length > 0 ? { merges } : {}),
    ...(internedInput.drawingArtifacts ? { drawingArtifacts: internedInput.drawingArtifacts } : {}),
    ...(internedInput.controlArtifacts ? { controlArtifacts: internedInput.controlArtifacts } : {}),
    ...(internedInput.pivotArtifacts ? { pivotArtifacts: internedInput.pivotArtifacts } : {}),
    ...(internedInput.legacyCommentVml ? { legacyCommentVml: internedInput.legacyCommentVml } : {}),
    ...(internedInput.sheetProtection ? { sheetProtection: internedInput.sheetProtection } : {}),
    ...(internedInput.filters ? { filters: internedInput.filters } : {}),
    ...(internedInput.hyperlinks ? { hyperlinks: internedInput.hyperlinks } : {}),
    ...(internedInput.validations ? { validations: internedInput.validations } : {}),
    ...(conditionalFormats ? { conditionalFormats } : {}),
    ...(internedInput.conditionalFormatArtifacts ? { conditionalFormatArtifacts: internedInput.conditionalFormatArtifacts } : {}),
    ...(cellMetadataRefs ? { cellMetadataRefs } : {}),
    ...(internedInput.printerSettings ? { printerSettings: internedInput.printerSettings } : {}),
    ...(internedInput.printPageSetup ? { printPageSetup: internedInput.printPageSetup } : {}),
    ...(cellScan.richTextCells.length > 0 ? { richTextArtifacts: { cells: cellScan.richTextCells } } : {}),
  }
  const sheet: WorkbookSnapshot['sheets'][number] = {
    id: order + 1,
    name: sheetName,
    order,
    ...(Object.keys(metadata).length > 0 ? { metadata } : {}),
    cells,
  }
  return {
    sheet,
    preview,
    stats: {
      cellCount: cellScan.cellCount,
      formulaCellCount: cellScan.formulaCellCount,
      valueCellCount: cellScan.valueCellCount,
      tableCount: cellScan.tableCount ?? 0,
      mergeCount,
      conditionalFormatCount,
      dataValidationCount,
      dimension: {
        sheetName,
        rowCount: cellScan.rowCount,
        columnCount: cellScan.columnCount,
        nonEmptyCellCount: cellScan.cellCount,
        usedRange: cellScan.usedRange,
      },
    },
  }
}
