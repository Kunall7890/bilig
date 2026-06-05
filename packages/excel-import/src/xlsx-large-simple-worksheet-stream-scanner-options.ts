import type { ImportedWorkbookArenaDedupeMode } from './xlsx-large-simple-arena.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'

export interface LargeSimpleWorksheetStreamScannerOptions {
  readonly hasSharedStrings: boolean
  readonly retainCells?: boolean
  readonly sharedStrings?: LargeSimpleSharedStrings
  readonly deferSharedStrings?: boolean
  readonly retainMetadataXml?: boolean
  readonly sheetName?: string
  readonly stringPool?: ImportedWorkbookStringPool
  readonly deduplicateStrings?: ImportedWorkbookArenaDedupeMode
  readonly deduplicateFormulas?: ImportedWorkbookArenaDedupeMode
  readonly dedupeMaxEntries?: number
  readonly allowUnsupportedFormulaText?: boolean
  readonly allowUnsupportedCellMetadata?: boolean
  readonly preserveBlankStyleCells?: boolean
  readonly retainStyleIndexes?: boolean
  readonly retainStyleCoordinates?: boolean
  readonly useWasmScanStorage?: boolean
  readonly maxDimensionCellPreallocation?: number
  readonly onRetainedBufferLength?: (length: number) => void
}

export interface ResolvedLargeSimpleWorksheetStreamScannerOptions {
  readonly hasSharedStrings: boolean
  readonly retainCells: boolean
  readonly sharedStrings: LargeSimpleSharedStrings
  readonly deferSharedStrings: boolean
  readonly retainMetadataXml: boolean
  readonly sheetName: string | undefined
  readonly stringPool: ImportedWorkbookStringPool | undefined
  readonly deduplicateStrings: ImportedWorkbookArenaDedupeMode | undefined
  readonly deduplicateFormulas: ImportedWorkbookArenaDedupeMode | undefined
  readonly dedupeMaxEntries: number | undefined
  readonly allowUnsupportedFormulaText: boolean | undefined
  readonly allowUnsupportedCellMetadata: boolean | undefined
  readonly preserveBlankStyleCells: boolean
  readonly retainStyleIndexes: boolean
  readonly retainStyleCoordinates: boolean
  readonly useWasmScanStorage: boolean | undefined
  readonly maxDimensionCellPreallocation: number | undefined
  readonly onRetainedBufferLength: ((length: number) => void) | undefined
}
