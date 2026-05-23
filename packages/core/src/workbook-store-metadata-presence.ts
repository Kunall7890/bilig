import type { WorkbookMetadataRecord } from './workbook-metadata-types.js'
import { hasPreservedSheetMetadata, hasPreservedWorkbookMetadata } from './workbook-preserved-metadata.js'
import type { SheetRecord } from './workbook-sheet-record.js'

export function hasStructuralMetadataForSheetRecord(
  metadata: WorkbookMetadataRecord,
  sheetName: string,
  sheet: SheetRecord | undefined,
): boolean {
  return (
    metadata.definedNames.size > 0 ||
    metadata.tables.size > 0 ||
    metadata.spills.size > 0 ||
    metadata.pivots.size > 0 ||
    metadata.charts.size > 0 ||
    metadata.images.size > 0 ||
    metadata.shapes.size > 0 ||
    metadata.drawingArtifacts !== undefined ||
    metadata.controlArtifacts !== undefined ||
    metadata.externalLinkArtifacts !== undefined ||
    hasPreservedWorkbookMetadata(metadata.preservedWorkbookMetadata) ||
    metadata.sheetDrawingArtifacts.has(sheetName) ||
    metadata.threadedCommentArtifacts !== undefined ||
    metadata.sheetThreadedCommentArtifacts.has(sheetName) ||
    metadata.sheetLegacyCommentVml.has(sheetName) ||
    hasPreservedSheetMetadata(metadata.preservedSheetMetadata.get(sheetName)) ||
    metadata.freezePanes.size > 0 ||
    metadata.merges.size > 0 ||
    metadata.filters.size > 0 ||
    metadata.sorts.size > 0 ||
    metadata.dataValidations.size > 0 ||
    metadata.conditionalFormats.size > 0 ||
    metadata.conditionalFormatArtifacts.has(sheetName) ||
    metadata.rangeProtections.size > 0 ||
    metadata.commentThreads.size > 0 ||
    metadata.notes.size > 0 ||
    metadata.hyperlinks.size > 0 ||
    (sheet?.arrayFormulas?.formulas.length ?? 0) > 0 ||
    (sheet?.dataTableFormulas?.formulas.length ?? 0) > 0 ||
    sheet?.sheetPr !== undefined ||
    sheet?.ignoredErrors !== undefined ||
    sheet?.printPageSetup !== undefined ||
    sheet?.sparklines !== undefined ||
    sheet?.controlArtifacts !== undefined ||
    sheet?.richTextArtifacts !== undefined ||
    sheet?.cellMetadataRefs !== undefined ||
    (sheet?.styleRanges.length ?? 0) > 0 ||
    (sheet?.formatRanges.length ?? 0) > 0
  )
}
