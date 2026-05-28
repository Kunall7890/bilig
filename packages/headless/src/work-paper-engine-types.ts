import type { SpreadsheetEngine } from '@bilig/core/headless-runtime'

export type MetadataRenameEngine = SpreadsheetEngine & {
  readonly hasWorkbookMetadataForSheetRenameFastPath?: () => boolean
  readonly renameSheetMetadataOnlyById?: (sheetId: number, newName: string) => boolean
  readonly renameSheetMetadataOnlyByIdPrevalidated?: (sheetId: number, oldName: string, newName: string) => boolean
  readonly renameSheetSimpleMetadataOnlyByIdTrustedPrevalidated: (
    sheetId: number,
    oldName: string,
    newName: string,
    assumeLocalRuntimeReady?: boolean,
    assumeNoWorkbookRenameMetadata?: boolean,
  ) => boolean
  readonly renameSheetMetadataOnlyByIdTrustedPrevalidated?: (
    sheetId: number,
    oldName: string,
    newName: string,
    assumeLocalRuntimeReady?: boolean,
    assumeNoWorkbookRenameMetadata?: boolean,
  ) => boolean
}

export type WorkPaperStructuralInsertEngine = SpreadsheetEngine & {
  insertRows(
    sheetName: string,
    start: number,
    count: number,
    options?: { readonly emitTracked?: boolean; readonly recordHistory?: boolean },
  ): void
  insertColumns(
    sheetName: string,
    start: number,
    count: number,
    options?: { readonly emitTracked?: boolean; readonly recordHistory?: boolean },
  ): void
}
