import type { SheetMetadataSnapshot, WorkbookConditionalFormatSnapshot } from '@bilig/protocol'

export type LargeSimpleSheetMetadataInput = Pick<
  SheetMetadataSnapshot,
  | 'conditionalFormatArtifacts'
  | 'conditionalFormats'
  | 'controlArtifacts'
  | 'validations'
  | 'drawingArtifacts'
  | 'filters'
  | 'hyperlinks'
  | 'legacyCommentVml'
  | 'pivotArtifacts'
  | 'printerSettings'
  | 'printPageSetup'
  | 'sheetProtection'
>

export function appendConditionalFormats(
  input: LargeSimpleSheetMetadataInput,
  conditionalFormats: readonly WorkbookConditionalFormatSnapshot[] | undefined,
): LargeSimpleSheetMetadataInput {
  if (!conditionalFormats || conditionalFormats.length === 0) {
    return input
  }
  return {
    ...input,
    conditionalFormats: [...(input.conditionalFormats ?? []), ...conditionalFormats],
  }
}

export function normalizeConditionalFormatIds(
  sheetName: string,
  conditionalFormats: readonly WorkbookConditionalFormatSnapshot[] | undefined,
): SheetMetadataSnapshot['conditionalFormats'] | undefined {
  if (!conditionalFormats || conditionalFormats.length === 0) {
    return undefined
  }
  return conditionalFormats.map((format, index) => ({
    ...format,
    id: `xlsx-cf:${sheetName}:${format.range.startAddress}:${format.range.endAddress}:${String(index + 1)}`,
  }))
}

export function readConditionalFormattingBlockCount(worksheetXml: string): number {
  return [...worksheetXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b/gu)].length
}
