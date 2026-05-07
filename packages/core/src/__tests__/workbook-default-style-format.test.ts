import { describe, expect, it } from 'vitest'
import {
  WORKBOOK_DEFAULT_FORMAT_ID,
  WORKBOOK_DEFAULT_STYLE_ID,
  ensureWorkbookDefaultStyleFormat,
} from '../workbook-default-style-format.js'
import type { WorkbookCellNumberFormatRecord, WorkbookCellStyleRecord } from '../workbook-metadata-types.js'

describe('workbook default style and format catalog', () => {
  it('seeds the default style and number format into empty catalogs', () => {
    const catalog = {
      cellStyles: new Map<string, WorkbookCellStyleRecord>(),
      styleKeys: new Map<string, string>(),
      cellNumberFormats: new Map<string, WorkbookCellNumberFormatRecord>(),
      numberFormatKeys: new Map<string, string>(),
    }

    ensureWorkbookDefaultStyleFormat(catalog)

    expect(catalog.cellStyles.get(WORKBOOK_DEFAULT_STYLE_ID)).toEqual({ id: WORKBOOK_DEFAULT_STYLE_ID })
    expect([...catalog.styleKeys.values()]).toEqual([WORKBOOK_DEFAULT_STYLE_ID])
    expect(catalog.cellNumberFormats.get(WORKBOOK_DEFAULT_FORMAT_ID)).toEqual({
      id: WORKBOOK_DEFAULT_FORMAT_ID,
      code: 'general',
      kind: 'general',
    })
    expect(catalog.numberFormatKeys.get('general')).toBe(WORKBOOK_DEFAULT_FORMAT_ID)
  })
})
