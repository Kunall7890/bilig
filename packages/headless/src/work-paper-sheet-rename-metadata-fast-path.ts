export type TrustedMetadataSheetRename = (
  sheetId: number,
  oldName: string,
  newName: string,
  assumeLocalRuntimeReady?: boolean,
  assumeNoWorkbookRenameMetadata?: boolean,
) => boolean

export function tryRenameSheetMetadataOnlyFast(
  sheetId: number,
  oldName: string,
  newName: string,
  batchDepth: number,
  batchUsesTrackedFastPath: boolean,
  emitterHasAnyListeners: boolean,
  engineEventsHasPendingLazyChanges: boolean,
  engineEventsHasTrackedEvents: boolean,
  evaluationSuspended: boolean,
  hasImportedXlsxSource: boolean,
  hasImportedXlsxSourceCellPatches: boolean,
  hasPreservedImportedSnapshot: boolean,
  hasVisibilityCache: boolean,
  hasNamedExpressions: boolean,
  renameSheetMetadataOnlyByIdTrustedPrevalidated: TrustedMetadataSheetRename | undefined,
  renameSheetMetadataOnlyByIdTrustedPrevalidatedThis: object,
): boolean {
  if (
    hasPreservedImportedSnapshot ||
    hasImportedXlsxSource ||
    hasImportedXlsxSourceCellPatches ||
    batchDepth !== 0 ||
    evaluationSuspended ||
    hasVisibilityCache ||
    emitterHasAnyListeners ||
    engineEventsHasPendingLazyChanges ||
    batchUsesTrackedFastPath ||
    engineEventsHasTrackedEvents
  ) {
    return false
  }
  return (
    renameSheetMetadataOnlyByIdTrustedPrevalidated?.call(
      renameSheetMetadataOnlyByIdTrustedPrevalidatedThis,
      sheetId,
      oldName,
      newName,
      true,
      !hasNamedExpressions,
    ) ?? false
  )
}
