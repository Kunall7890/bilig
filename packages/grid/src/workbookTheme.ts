export const WORKBOOK_FONT_SANS = 'Arial, "Helvetica Neue", Helvetica, "Segoe UI", sans-serif'
export const WORKBOOK_DEFAULT_FONT_SIZE = 10
export const WORKBOOK_HEADER_FONT_POINT_SIZE = WORKBOOK_DEFAULT_FONT_SIZE
export const WORKBOOK_FONT_POINT_TO_CSS_PX = 4 / 3

export function workbookFontPointSizeToCssPx(pointSize: number): number {
  return Math.max(1, Number((Math.max(1, pointSize) * WORKBOOK_FONT_POINT_TO_CSS_PX).toFixed(4)))
}

export function workbookHeaderFontPointSizeToCssPx(pointSize = WORKBOOK_HEADER_FONT_POINT_SIZE): number {
  return workbookFontPointSizeToCssPx(pointSize)
}

export function workbookDisplayFontCssPx(fontSizeCssPx: number): number {
  return Math.max(1, Math.round(fontSizeCssPx))
}

export function workbookDisplayFontPointSizeToCssPx(pointSize: number): number {
  return workbookDisplayFontCssPx(workbookFontPointSizeToCssPx(pointSize))
}

export function workbookDisplayLineHeightCssPx(fontSizeCssPx: number): number {
  return Math.max(1, Math.round(workbookDisplayFontCssPx(fontSizeCssPx) * 1.2))
}

export const workbookThemeColors = {
  accent: '#21563a',
  accentDark: '#163f29',
  accentSoft: 'rgba(33, 86, 58, 0.2)',
  selectionAccent: '#1a73e8',
  selectionFill: 'rgba(26, 115, 232, 0.16)',
  selectionHeaderFill: 'rgba(26, 115, 232, 0.14)',
  border: '#ddd8cc',
  borderSubtle: '#e7e2d6',
  gridBorder: '#ddd8cc',
  hoverFill: 'rgba(82, 96, 109, 0.08)',
  hoverOutline: 'rgba(82, 96, 109, 0.42)',
  muted: '#f0ece3',
  mutedStrong: '#e4ddd0',
  surface: '#ffffff',
  surfaceSubtle: '#f3f2ee',
  text: '#1f2933',
  textMuted: '#52606d',
  textSubtle: '#7b8794',
} as const
