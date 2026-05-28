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

export function hasWorkbookMetadataForSheetRename(metadata: WorkbookMetadataRecord): boolean {
  const recordCount =
    metadata.macroPayloads.size +
    metadata.definedNames.size +
    metadata.tables.size +
    metadata.spills.size +
    metadata.pivots.size +
    metadata.charts.size +
    metadata.images.size +
    metadata.shapes.size +
    metadata.preservedSheetMetadata.size +
    metadata.sheetThreadedCommentArtifacts.size +
    metadata.sheetLegacyCommentVml.size +
    metadata.sheetDrawingArtifacts.size +
    metadata.rowMetadata.size +
    metadata.columnMetadata.size +
    metadata.freezePanes.size +
    metadata.sheetTabColors.size +
    metadata.merges.size +
    metadata.sheetProtections.size +
    metadata.filters.size +
    metadata.sorts.size +
    metadata.dataValidations.size +
    metadata.conditionalFormats.size +
    metadata.conditionalFormatArtifacts.size +
    metadata.rangeProtections.size +
    metadata.commentThreads.size +
    metadata.notes.size +
    metadata.hyperlinks.size
  return (
    recordCount !== 0 ||
    metadata.drawingArtifacts !== undefined ||
    metadata.controlArtifacts !== undefined ||
    metadata.externalLinkArtifacts !== undefined ||
    metadata.threadedCommentArtifacts !== undefined ||
    metadata.cellMetadata !== undefined ||
    hasPreservedWorkbookMetadataForSheetRename(metadata.preservedWorkbookMetadata)
  )
}

function hasPreservedWorkbookMetadataForSheetRename(metadata: WorkbookMetadataRecord['preservedWorkbookMetadata']): boolean {
  return (
    metadata.documentPropertyArtifacts !== undefined ||
    metadata.externalWorkbookReferences !== undefined ||
    metadata.unsupportedFormulaDependencies !== undefined ||
    metadata.unsupportedPivots !== undefined ||
    metadata.formulaAudit !== undefined ||
    metadata.externalConnections !== undefined ||
    metadata.pivotArtifacts !== undefined ||
    metadata.chartArtifacts !== undefined ||
    metadata.chartSheetArtifacts !== undefined ||
    metadata.dataModelArtifacts !== undefined ||
    metadata.slicerConnectionArtifacts !== undefined ||
    metadata.viewState !== undefined ||
    metadata.styleArtifacts !== undefined
  )
}
