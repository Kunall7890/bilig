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

export function workbookSnapCssPixel(value: number, dpr = 1): number {
  const resolvedDpr = Number.isFinite(dpr) ? Math.max(1, dpr) : 1
  return Number((Math.round(value * resolvedDpr) / resolvedDpr).toFixed(4))
}

export function workbookSnapTextCssPixel(value: number): number {
  return Math.max(1, Math.round(Number.isFinite(value) ? value : workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)))
}

export function workbookDisplayFontCssPx(fontSizeCssPx: number, _dpr = 1): number {
  const resolvedFontSize = Number.isFinite(fontSizeCssPx)
    ? Math.max(1, fontSizeCssPx)
    : workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)
  return workbookSnapTextCssPixel(resolvedFontSize)
}

export function workbookDisplayFontPointSizeToCssPx(pointSize: number, dpr = 1): number {
  return workbookDisplayFontCssPx(workbookFontPointSizeToCssPx(pointSize), dpr)
}

export function workbookDisplayLineHeightCssPx(fontSizeCssPx: number, dpr = 1): number {
  return workbookSnapTextCssPixel(workbookDisplayFontCssPx(fontSizeCssPx, dpr) * 1.2)
}

export const workbookThemeColors = {
  accent: '#21563a',
  accentDark: '#163f29',
  accentSoft: 'rgba(33, 86, 58, 0.2)',
  selectionAccent: '#217346',
  selectionFill: 'rgba(33, 115, 70, 0.22)',
  selectionHeaderFill: 'rgba(33, 115, 70, 0.16)',
  selectionHeaderSeamFill: '#d1ddd2',
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
