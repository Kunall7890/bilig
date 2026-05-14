import { describe, expect, test } from 'vitest'
import { WORKBOOK_FONT_SANS } from '../workbookTheme.js'

describe('workbookTheme', () => {
  test('uses the product font stack for rendered grid text', () => {
    expect(WORKBOOK_FONT_SANS.startsWith('"IBM Plex Sans", Inter')).toBe(true)
    expect(WORKBOOK_FONT_SANS).toContain('"SF Pro Text"')
    expect(WORKBOOK_FONT_SANS).toContain('Arial')
    expect(WORKBOOK_FONT_SANS).not.toContain('Aptos')
    expect(WORKBOOK_FONT_SANS).not.toContain('Calibri')
  })
})
