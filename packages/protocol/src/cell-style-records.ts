import type {
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellStyleRecord,
  CellVerticalAlignment,
} from './cell-format-types.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined
}

function asBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

function asFiniteNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined
}

function asCellHorizontalAlignment(value: unknown): CellHorizontalAlignment | undefined {
  switch (value) {
    case 'general':
    case 'left':
    case 'center':
    case 'right':
    case 'fill':
    case 'justify':
    case 'centerContinuous':
    case 'distributed':
      return value
    default:
      return undefined
  }
}

function asCellVerticalAlignment(value: unknown): CellVerticalAlignment | undefined {
  switch (value) {
    case 'top':
    case 'middle':
    case 'bottom':
    case 'justify':
    case 'distributed':
      return value
    default:
      return undefined
  }
}

function asCellBorderStyle(value: unknown): CellBorderStyle | undefined {
  switch (value) {
    case 'solid':
    case 'dashed':
    case 'dotted':
    case 'double':
      return value
    default:
      return undefined
  }
}

function asCellBorderWeight(value: unknown): CellBorderWeight | undefined {
  switch (value) {
    case 'thin':
    case 'medium':
    case 'thick':
      return value
    default:
      return undefined
  }
}

function sanitizeStyleFill(value: unknown): CellStyleRecord['fill'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const backgroundColor = asString(value['backgroundColor'])
  return backgroundColor ? { backgroundColor } : undefined
}

function sanitizeStyleFont(value: unknown): CellStyleRecord['font'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const font: NonNullable<CellStyleRecord['font']> = {}
  const family = asString(value['family'])
  const size = asFiniteNumber(value['size'])
  const bold = asBoolean(value['bold'])
  const italic = asBoolean(value['italic'])
  const underline = asBoolean(value['underline'])
  const color = asString(value['color'])
  if (family) {
    font.family = family
  }
  if (size !== undefined) {
    font.size = size
  }
  if (bold !== undefined) {
    font.bold = bold
  }
  if (italic !== undefined) {
    font.italic = italic
  }
  if (underline !== undefined) {
    font.underline = underline
  }
  if (color) {
    font.color = color
  }
  return Object.keys(font).length > 0 ? font : undefined
}

function sanitizeStyleAlignment(value: unknown): CellStyleRecord['alignment'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const alignment: NonNullable<CellStyleRecord['alignment']> = {}
  const horizontal = asCellHorizontalAlignment(value['horizontal'])
  const vertical = asCellVerticalAlignment(value['vertical'])
  const wrap = asBoolean(value['wrap'])
  const indent = asFiniteNumber(value['indent'])
  const shrinkToFit = asBoolean(value['shrinkToFit'])
  const readingOrder = asFiniteNumber(value['readingOrder'])
  const textRotation = asFiniteNumber(value['textRotation'])
  const justifyLastLine = asBoolean(value['justifyLastLine'])
  if (horizontal) {
    alignment.horizontal = horizontal
  }
  if (vertical) {
    alignment.vertical = vertical
  }
  if (wrap !== undefined) {
    alignment.wrap = wrap
  }
  if (indent !== undefined) {
    alignment.indent = indent
  }
  if (shrinkToFit !== undefined) {
    alignment.shrinkToFit = shrinkToFit
  }
  if (readingOrder !== undefined) {
    alignment.readingOrder = readingOrder
  }
  if (textRotation !== undefined) {
    alignment.textRotation = textRotation
  }
  if (justifyLastLine !== undefined) {
    alignment.justifyLastLine = justifyLastLine
  }
  return Object.keys(alignment).length > 0 ? alignment : undefined
}

function sanitizeBorderSide(value: unknown): CellBorderSideSnapshot | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const style = asCellBorderStyle(value['style'])
  const weight = asCellBorderWeight(value['weight'])
  const color = asString(value['color'])
  if (!style || !weight || !color) {
    return undefined
  }
  return { style, weight, color }
}

function sanitizeStyleBorders(value: unknown): CellStyleRecord['borders'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const borders: NonNullable<CellStyleRecord['borders']> = {}
  const top = sanitizeBorderSide(value['top'])
  const right = sanitizeBorderSide(value['right'])
  const bottom = sanitizeBorderSide(value['bottom'])
  const left = sanitizeBorderSide(value['left'])
  if (top) {
    borders.top = top
  }
  if (right) {
    borders.right = right
  }
  if (bottom) {
    borders.bottom = bottom
  }
  if (left) {
    borders.left = left
  }
  return Object.keys(borders).length > 0 ? borders : undefined
}

function sanitizeStyleProtection(value: unknown): CellStyleRecord['protection'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const protection: NonNullable<CellStyleRecord['protection']> = {}
  const locked = asBoolean(value['locked'])
  const hidden = asBoolean(value['hidden'])
  if (locked !== undefined) {
    protection.locked = locked
  }
  if (hidden !== undefined) {
    protection.hidden = hidden
  }
  return Object.keys(protection).length > 0 ? protection : undefined
}

export function sanitizeCellStyleRecord(id: string, value: unknown): CellStyleRecord | null {
  if (!id || !isRecord(value)) {
    return null
  }
  const style: CellStyleRecord = { id }
  const fill = sanitizeStyleFill(value['fill'])
  const font = sanitizeStyleFont(value['font'])
  const alignment = sanitizeStyleAlignment(value['alignment'])
  const borders = sanitizeStyleBorders(value['borders'])
  const protection = sanitizeStyleProtection(value['protection'])
  if (fill) {
    style.fill = fill
  }
  if (font) {
    style.font = font
  }
  if (alignment) {
    style.alignment = alignment
  }
  if (borders) {
    style.borders = borders
  }
  if (protection) {
    style.protection = protection
  }
  return style
}
