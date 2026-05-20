import type { WorkbookSnapshot } from '@bilig/protocol'
import type { ImportedWorkbookPreview } from './workbook-import-preview.js'
import type { LargeSimpleXlsxImportStats } from './xlsx-large-simple-import.js'

export interface ImportedWorkbook {
  snapshot: WorkbookSnapshot
  workbookName: string
  sheetNames: string[]
  warnings: string[]
  preview: ImportedWorkbookPreview
  stats?: LargeSimpleXlsxImportStats
}
