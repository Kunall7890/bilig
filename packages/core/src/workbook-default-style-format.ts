import type { WorkbookCellNumberFormatRecord, WorkbookCellStyleRecord } from './workbook-metadata-types.js'
import { cellStyleKey } from './workbook-store-records.js'

export const WORKBOOK_DEFAULT_STYLE_ID = 'style-0'
export const WORKBOOK_DEFAULT_FORMAT_ID = 'format-0'

export interface WorkbookDefaultStyleFormatCatalog {
  readonly cellStyles: Map<string, WorkbookCellStyleRecord>
  readonly styleKeys: Map<string, string>
  readonly cellNumberFormats: Map<string, WorkbookCellNumberFormatRecord>
  readonly numberFormatKeys: Map<string, string>
}

export function ensureWorkbookDefaultStyleFormat(catalog: WorkbookDefaultStyleFormatCatalog): void {
  ensureWorkbookDefaultStyle(catalog)
  ensureWorkbookDefaultNumberFormat(catalog)
}

export function ensureWorkbookDefaultStyle(catalog: Pick<WorkbookDefaultStyleFormatCatalog, 'cellStyles' | 'styleKeys'>): void {
  const defaultStyle: WorkbookCellStyleRecord = { id: WORKBOOK_DEFAULT_STYLE_ID }
  catalog.cellStyles.set(defaultStyle.id, defaultStyle)
  catalog.styleKeys.set(cellStyleKey(defaultStyle), defaultStyle.id)
}

export function ensureWorkbookDefaultNumberFormat(
  catalog: Pick<WorkbookDefaultStyleFormatCatalog, 'cellNumberFormats' | 'numberFormatKeys'>,
): void {
  const defaultFormat: WorkbookCellNumberFormatRecord = {
    id: WORKBOOK_DEFAULT_FORMAT_ID,
    code: 'general',
    kind: 'general',
  }
  catalog.cellNumberFormats.set(defaultFormat.id, defaultFormat)
  catalog.numberFormatKeys.set(defaultFormat.code, defaultFormat.id)
}
