import type { CellStyleRecord, SheetMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { createSheetPreview, toDisplayText } from './workbook-import-helpers.js'
import { buildLargeSimpleCellMetadataReferenceSnapshots } from './xlsx-large-simple-cell-metadata.js'
import {
  normalizeLargeSimpleConditionalFormatIds,
  readLargeSimpleConditionalFormattingBlockCount,
} from './xlsx-large-simple-conditional-format-helpers.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import { buildLargeSimpleStyleRanges } from './xlsx-large-simple-style-ranges.js'
import { applyLargeSimpleNumberFormatsToCells } from './xlsx-large-simple-number-formats.js'
import { releaseProjectedCellScanStorage } from './xlsx-large-simple-materialization-helpers.js'
import {
  readLargeSimpleColumnMetadata,
  readLargeSimpleMergeRanges,
  readLargeSimpleRowMetadata,
  readLargeSimpleSheetFormatPr,
  type LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'
import type { LargeSimpleSheetMetadataInput, ParsedWorksheet } from './xlsx-large-simple-import-types.js'
import { decodeCellAddress } from './xlsx-large-simple-xml-byte-utils.js'

export const lazySheetCellMaterializationThreshold = 100_000

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
  } = { materializeCells: true },
): ParsedWorksheet {
  const merges =
    metadataScan?.merges?.map((range) => ({ sheetName, ...range })) ??
    (worksheetXml ? readLargeSimpleMergeRanges(sheetName, worksheetXml) : [])
  const mergeCount = worksheetXml ? merges.length : (cellScan.mergeCount ?? 0)
  const columns = metadataScan?.columns ?? (worksheetXml ? readLargeSimpleColumnMetadata(worksheetXml) : { entries: [], metadata: [] })
  const rows = metadataScan?.rows ?? (worksheetXml ? readLargeSimpleRowMetadata(worksheetXml) : { entries: [], metadata: [] })
  const sheetFormatPr = metadataScan?.sheetFormatPr ?? (worksheetXml ? readLargeSimpleSheetFormatPr(worksheetXml) : undefined)
  const conditionalFormatCount =
    input.conditionalFormats?.length ??
    (worksheetXml ? readLargeSimpleConditionalFormattingBlockCount(worksheetXml) : (cellScan.conditionalFormatCount ?? 0))
  const conditionalFormats = normalizeLargeSimpleConditionalFormatIds(sheetName, input.conditionalFormats)
  const dataValidationCount = input.validations?.length ?? cellScan.dataValidationCount ?? 0
  const styleRanges =
    options.materializeCells && options.styleCatalog && options.stylesByIndex
      ? buildLargeSimpleStyleRanges(sheetName, cellScan, options.stylesByIndex, options.styleCatalog)
      : []
  const useLazyCells = options.materializeCells && cellScan.cellCount > lazySheetCellMaterializationThreshold
  const cells = options.materializeCells
    ? useLazyCells
      ? cellScan.arena.createLazySheetCells(cellScan.sheetIndex)
      : cellScan.arena.materializeSheetCells(cellScan.sheetIndex)
    : []
  if (!useLazyCells && options.numberFormatsByStyleIndex && options.numberFormatsByStyleIndex.size > 0) {
    applyLargeSimpleNumberFormatsToCells(cells, cellScan, options.numberFormatsByStyleIndex)
  }
  const cellMetadataRefs = buildLargeSimpleCellMetadataReferenceSnapshots(metadataScan?.cellMetadataRefs, cells, cellScan, useLazyCells)
  const preview =
    options.materializeCells && !useLazyCells
      ? createPreviewFromMaterializedCells(sheetName, cellScan, cells)
      : createSheetPreview({
          name: sheetName,
          rowCount: cellScan.rowCount,
          columnCount: cellScan.columnCount,
          nonEmptyCellCount: cellScan.cellCount,
          readCellText: (row, column) => cellScan.arena.readPreviewText(row, column),
        })
  releaseProjectedCellScanStorage(cellScan, {
    releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization,
    useLazyCells,
  })
  const metadata: SheetMetadataSnapshot = {
    ...(columns.entries.length > 0 ? { columns: columns.entries } : {}),
    ...(rows.entries.length > 0 ? { rows: rows.entries } : {}),
    ...(columns.metadata.length > 0 ? { columnMetadata: columns.metadata } : {}),
    ...(rows.metadata.length > 0 ? { rowMetadata: rows.metadata } : {}),
    ...(sheetFormatPr ? { sheetFormatPr } : {}),
    ...(styleRanges.length > 0 ? { styleRanges } : {}),
    ...(merges.length > 0 ? { merges } : {}),
    ...(input.drawingArtifacts ? { drawingArtifacts: input.drawingArtifacts } : {}),
    ...(input.controlArtifacts ? { controlArtifacts: input.controlArtifacts } : {}),
    ...(input.pivotArtifacts ? { pivotArtifacts: input.pivotArtifacts } : {}),
    ...(input.legacyCommentVml ? { legacyCommentVml: input.legacyCommentVml } : {}),
    ...(input.sheetProtection ? { sheetProtection: input.sheetProtection } : {}),
    ...(input.filters ? { filters: input.filters } : {}),
    ...(input.hyperlinks ? { hyperlinks: input.hyperlinks } : {}),
    ...(input.validations ? { validations: input.validations } : {}),
    ...(conditionalFormats ? { conditionalFormats } : {}),
    ...(input.conditionalFormatArtifacts ? { conditionalFormatArtifacts: input.conditionalFormatArtifacts } : {}),
    ...(cellMetadataRefs ? { cellMetadataRefs } : {}),
    ...(input.printerSettings ? { printerSettings: input.printerSettings } : {}),
    ...(input.printPageSetup ? { printPageSetup: input.printPageSetup } : {}),
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

function createPreviewFromMaterializedCells(
  sheetName: string,
  cellScan: ImportedWorksheetCellScan,
  cells: WorkbookSnapshot['sheets'][number]['cells'],
): ParsedWorksheet['preview'] {
  const rowLimit = Math.min(cellScan.rowCount, 8)
  const columnLimit = Math.min(cellScan.columnCount, 6)
  const previewValues = Array.from({ length: rowLimit }, () => Array.from({ length: columnLimit }, () => ''))
  for (const cell of cells) {
    const decoded = decodeCellAddress(cell.address)
    if (!decoded || decoded.row >= rowLimit || decoded.column >= columnLimit) {
      continue
    }
    previewValues[decoded.row]![decoded.column] = 'value' in cell ? toDisplayText(cell.value) : 'formula' in cell ? `=${cell.formula}` : ''
  }
  return {
    name: sheetName,
    rowCount: cellScan.rowCount,
    columnCount: cellScan.columnCount,
    nonEmptyCellCount: cellScan.cellCount,
    previewRows: previewValues,
  }
}
