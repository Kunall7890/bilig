import { describe, expect, test } from 'vitest'
import { WORKBOOK_FONT_SANS } from '../workbookTheme.js'

describe('workbookTheme', () => {
  test('uses an Arial-first spreadsheet font stack for rendered grid text', () => {
    expect(WORKBOOK_FONT_SANS.startsWith('Arial,')).toBe(true)
    expect(WORKBOOK_FONT_SANS).not.toContain('Aptos')
    expect(WORKBOOK_FONT_SANS).not.toContain('Calibri')
  })
})
