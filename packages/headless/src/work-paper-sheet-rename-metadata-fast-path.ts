import type { MetadataRenameEngine } from './work-paper-engine-types.js'

export function tryRenameSheetMetadataOnlyPrevalidated(
  engine: MetadataRenameEngine,
  sheetId: number,
  oldName: string,
  newName: string,
  hasWorkbookRenameMetadata: boolean,
): boolean {
  return (
    engine.renameSheetMetadataOnlyByIdTrustedPrevalidated?.(sheetId, oldName, newName, false, !hasWorkbookRenameMetadata) ??
    engine.renameSheetMetadataOnlyByIdPrevalidated?.(sheetId, oldName, newName) ??
    engine.renameSheetMetadataOnlyById?.(sheetId, newName) ??
    engine.renameSheetMetadataOnly(oldName, newName)
  )
}
