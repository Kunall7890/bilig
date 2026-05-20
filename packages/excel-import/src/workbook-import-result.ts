import type { WorkbookSnapshot } from '@bilig/protocol'
import type { ImportedWorkbookPreview } from './workbook-import-preview.js'

export interface ImportedWorkbook {
  snapshot: WorkbookSnapshot
  workbookName: string
  sheetNames: string[]
  warnings: string[]
  preview: ImportedWorkbookPreview
}
