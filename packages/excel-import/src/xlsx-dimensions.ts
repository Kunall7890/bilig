import { unzipSync, zipSync } from 'fflate'

import type {
  SheetMetadataSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookSheetFormatPrSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { escapeXmlAttribute, getZipText, setXmlAttribute, setZipText } from './xlsx-export-xml.js'

interface ExportRowMetadata {
  readonly rowNumber: number
  readonly size?: number
  readonly hidden?: boolean
  readonly xlsxHeight?: number
  readonly customHeight?: boolean
  readonly outlineLevel?: number
  readonly collapsed?: boolean
  readonly thickTop?: boolean
  readonly thickBottom?: boolean
  readonly exact: boolean
}

interface ExportColumnMetadata {
  start: number
  count: number
  xlsxWidth?: number
  customWidth?: boolean
  bestFit?: boolean
  hidden?: boolean
  outlineLevel?: number
  collapsed?: boolean
}

function finitePositiveNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : undefined
}

function finiteNonNegativeInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 ? value : undefined
}

function optionalBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

function formatXmlNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(12)))
}

function formatXmlBoolean(value: boolean): string {
  return value ? '1' : '0'
}

function xmlAttribute(name: string, value: string | undefined): string {
  return value === undefined ? '' : ` ${name}="${escapeXmlAttribute(value)}"`
}

function removeXmlAttribute(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${name}="[^"]*"`, 'u'), '')
}

function readXmlAttribute(tag: string, name: string): string | null {
  const doubleQuoted = new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(tag)
  if (doubleQuoted) {
    return doubleQuoted[1] ?? null
  }
  const singleQuoted = new RegExp(`\\b${name}='([^']*)'`, 'u').exec(tag)
  return singleQuoted?.[1] ?? null
}

function readXmlNumberAttribute(tag: string, name: string): number | null {
  const raw = readXmlAttribute(tag, name)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const number = Number(raw)
  return Number.isFinite(number) ? number : null
}

function readXmlPositiveIntegerAttribute(tag: string, name: string): number | null {
  const number = readXmlNumberAttribute(tag, name)
  return number !== null && Number.isSafeInteger(number) && number > 0 ? number : null
}

function readXmlOptionalBooleanAttribute(tag: string, name: string): boolean | undefined {
  const raw = readXmlAttribute(tag, name)
  if (raw === null) {
    return undefined
  }
  return raw === '1' || raw.toLowerCase() === 'true'
}

function hasExactRowGeometry(row: WorkbookAxisMetadataSnapshot): boolean {
  return (
    row.xlsxHeight !== undefined ||
    row.customHeight !== undefined ||
    row.outlineLevel !== undefined ||
    row.collapsed !== undefined ||
    row.thickTop !== undefined ||
    row.thickBottom !== undefined
  )
}

function hasExactColumnGeometry(column: WorkbookAxisMetadataSnapshot): boolean {
  return (
    column.xlsxWidth !== undefined ||
    column.customWidth !== undefined ||
    column.bestFit !== undefined ||
    column.outlineLevel !== undefined ||
    column.collapsed !== undefined
  )
}

function normalizeExportRowMetadata(
  rows: readonly WorkbookAxisEntrySnapshot[] | undefined,
  rowMetadata: readonly WorkbookAxisMetadataSnapshot[] | undefined,
): ExportRowMetadata[] {
  const exactRows =
    rowMetadata
      ?.flatMap((row) => {
        if (!Number.isSafeInteger(row.start) || row.start < 0 || row.count !== 1 || !hasExactRowGeometry(row)) {
          return []
        }
        const xlsxHeight = finitePositiveNumber(row.xlsxHeight ?? undefined)
        const size = finitePositiveNumber(row.size ?? undefined)
        const hidden = optionalBoolean(row.hidden)
        const customHeight = optionalBoolean(row.customHeight)
        const outlineLevel = finiteNonNegativeInteger(row.outlineLevel ?? undefined)
        const collapsed = optionalBoolean(row.collapsed)
        const thickTop = optionalBoolean(row.thickTop)
        const thickBottom = optionalBoolean(row.thickBottom)
        if (xlsxHeight === undefined && size === undefined && row.hidden !== true && row.hidden !== false && !hasExactRowGeometry(row)) {
          return []
        }
        const normalized: ExportRowMetadata = {
          rowNumber: row.start + 1,
          ...(size !== undefined ? { size } : {}),
          ...(hidden !== undefined ? { hidden } : {}),
          ...(xlsxHeight !== undefined ? { xlsxHeight } : {}),
          ...(customHeight !== undefined ? { customHeight } : {}),
          ...(outlineLevel !== undefined ? { outlineLevel } : {}),
          ...(collapsed !== undefined ? { collapsed } : {}),
          ...(thickTop !== undefined ? { thickTop } : {}),
          ...(thickBottom !== undefined ? { thickBottom } : {}),
          exact: true,
        }
        return [normalized]
      })
      .toSorted((left, right) => left.rowNumber - right.rowNumber) ?? []
  if (!rows || rows.length === 0) {
    return exactRows
  }
  const exactRowNumbers = new Set(exactRows.map((row) => row.rowNumber))
  const fallbackRows = rows.flatMap((row) => {
    if (!Number.isSafeInteger(row.index) || row.index < 0) {
      return []
    }
    const rowNumber = row.index + 1
    if (exactRowNumbers.has(rowNumber)) {
      return []
    }
    const size = finitePositiveNumber(row.size ?? undefined)
    if (size === undefined && row.hidden !== true) {
      return []
    }
    return [
      {
        rowNumber,
        ...(size !== undefined ? { size } : {}),
        ...(row.hidden === true ? { hidden: true } : {}),
        exact: false,
      },
    ]
  })
  return [...exactRows, ...fallbackRows].toSorted((left, right) => left.rowNumber - right.rowNumber)
}

function normalizeExportColumnMetadata(columnMetadata: readonly WorkbookAxisMetadataSnapshot[] | undefined): ExportColumnMetadata[] {
  if (!columnMetadata || columnMetadata.length === 0) {
    return []
  }
  return columnMetadata
    .flatMap((column) => {
      if (
        !Number.isSafeInteger(column.start) ||
        column.start < 0 ||
        !Number.isSafeInteger(column.count) ||
        column.count <= 0 ||
        !hasExactColumnGeometry(column)
      ) {
        return []
      }
      const xlsxWidth = finitePositiveNumber(column.xlsxWidth ?? undefined)
      if (xlsxWidth === undefined) {
        return []
      }
      const customWidth = optionalBoolean(column.customWidth)
      const bestFit = optionalBoolean(column.bestFit)
      const hidden = optionalBoolean(column.hidden)
      const outlineLevel = finiteNonNegativeInteger(column.outlineLevel ?? undefined)
      const collapsed = optionalBoolean(column.collapsed)
      const normalized: ExportColumnMetadata = {
        start: column.start,
        count: column.count,
        xlsxWidth,
        ...(customWidth !== undefined ? { customWidth } : {}),
        ...(bestFit !== undefined ? { bestFit } : {}),
        ...(hidden !== undefined ? { hidden } : {}),
        ...(outlineLevel !== undefined ? { outlineLevel } : {}),
        ...(collapsed !== undefined ? { collapsed } : {}),
      }
      return [normalized]
    })
    .toSorted((left, right) => left.start - right.start || left.count - right.count)
}

function parseExistingColumnMetadata(sheetXml: string): ExportColumnMetadata[] {
  return [...sheetXml.matchAll(/<col\b[^>]*\/?>/gu)].flatMap((match) => {
    const columnTag = match[0]
    const min = readXmlPositiveIntegerAttribute(columnTag, 'min')
    const max = readXmlPositiveIntegerAttribute(columnTag, 'max') ?? min
    if (min === null || max === null || max < min) {
      return []
    }
    const xlsxWidth = finitePositiveNumber(readXmlNumberAttribute(columnTag, 'width') ?? undefined)
    const customWidth = readXmlOptionalBooleanAttribute(columnTag, 'customWidth')
    const bestFit = readXmlOptionalBooleanAttribute(columnTag, 'bestFit')
    const hidden = readXmlOptionalBooleanAttribute(columnTag, 'hidden')
    const outlineLevel = finiteNonNegativeInteger(readXmlNumberAttribute(columnTag, 'outlineLevel') ?? undefined)
    const collapsed = readXmlOptionalBooleanAttribute(columnTag, 'collapsed')
    if (
      xlsxWidth === undefined &&
      customWidth === undefined &&
      bestFit === undefined &&
      hidden === undefined &&
      outlineLevel === undefined &&
      collapsed === undefined
    ) {
      return []
    }
    const column: ExportColumnMetadata = {
      start: min - 1,
      count: max - min + 1,
      ...(xlsxWidth !== undefined ? { xlsxWidth } : {}),
      ...(customWidth !== undefined ? { customWidth } : {}),
      ...(bestFit !== undefined ? { bestFit } : {}),
      ...(hidden !== undefined ? { hidden } : {}),
      ...(outlineLevel !== undefined ? { outlineLevel } : {}),
      ...(collapsed !== undefined ? { collapsed } : {}),
    }
    return [column]
  })
}

function columnRangeEnd(column: Pick<ExportColumnMetadata, 'start' | 'count'>): number {
  return column.start + column.count - 1
}

function subtractColumnRanges(column: ExportColumnMetadata, exactColumns: readonly ExportColumnMetadata[]): ExportColumnMetadata[] {
  let segments: Array<{ start: number; count: number }> = [{ start: column.start, count: column.count }]
  for (const exactColumn of exactColumns) {
    const exactStart = exactColumn.start
    const exactEnd = columnRangeEnd(exactColumn)
    segments = segments.flatMap((segment) => {
      const segmentEnd = columnRangeEnd(segment)
      if (segmentEnd < exactStart || segment.start > exactEnd) {
        return [segment]
      }
      const output: Array<{ start: number; count: number }> = []
      if (segment.start < exactStart) {
        output.push({ start: segment.start, count: exactStart - segment.start })
      }
      if (segmentEnd > exactEnd) {
        output.push({ start: exactEnd + 1, count: segmentEnd - exactEnd })
      }
      return output
    })
  }
  return segments.map((segment) => {
    const output: ExportColumnMetadata = {
      start: segment.start,
      count: segment.count,
    }
    if (column.xlsxWidth !== undefined) {
      output.xlsxWidth = column.xlsxWidth
    }
    if (column.customWidth !== undefined) {
      output.customWidth = column.customWidth
    }
    if (column.bestFit !== undefined) {
      output.bestFit = column.bestFit
    }
    if (column.hidden !== undefined) {
      output.hidden = column.hidden
    }
    if (column.outlineLevel !== undefined) {
      output.outlineLevel = column.outlineLevel
    }
    if (column.collapsed !== undefined) {
      output.collapsed = column.collapsed
    }
    return output
  })
}

function buildSheetFormatPrXml(sheetFormatPr: WorkbookSheetFormatPrSnapshot | undefined): string | null {
  if (!sheetFormatPr) {
    return null
  }
  const baseColWidth = finiteNonNegativeInteger(sheetFormatPr.baseColWidth ?? undefined)
  const defaultColWidth = finitePositiveNumber(sheetFormatPr.defaultColWidth ?? undefined)
  const defaultRowHeight = finitePositiveNumber(sheetFormatPr.defaultRowHeight ?? undefined)
  const customHeight = optionalBoolean(sheetFormatPr.customHeight)
  const outlineLevelRow = finiteNonNegativeInteger(sheetFormatPr.outlineLevelRow ?? undefined)
  const outlineLevelCol = finiteNonNegativeInteger(sheetFormatPr.outlineLevelCol ?? undefined)
  const thickTop = optionalBoolean(sheetFormatPr.thickTop)
  const thickBottom = optionalBoolean(sheetFormatPr.thickBottom)
  const attributes = [
    xmlAttribute('baseColWidth', baseColWidth !== undefined ? formatXmlNumber(baseColWidth) : undefined),
    xmlAttribute('defaultColWidth', defaultColWidth !== undefined ? formatXmlNumber(defaultColWidth) : undefined),
    xmlAttribute('defaultRowHeight', defaultRowHeight !== undefined ? formatXmlNumber(defaultRowHeight) : undefined),
    xmlAttribute('customHeight', customHeight !== undefined ? formatXmlBoolean(customHeight) : undefined),
    xmlAttribute('outlineLevelRow', outlineLevelRow !== undefined ? formatXmlNumber(outlineLevelRow) : undefined),
    xmlAttribute('outlineLevelCol', outlineLevelCol !== undefined ? formatXmlNumber(outlineLevelCol) : undefined),
    xmlAttribute('thickTop', thickTop !== undefined ? formatXmlBoolean(thickTop) : undefined),
    xmlAttribute('thickBottom', thickBottom !== undefined ? formatXmlBoolean(thickBottom) : undefined),
  ].join('')
  return attributes.length > 0 ? `<sheetFormatPr${attributes}/>` : null
}

function applySheetFormatPr(sheetXml: string, sheetFormatPr: WorkbookSheetFormatPrSnapshot | undefined): string {
  const sheetFormatPrXml = buildSheetFormatPrXml(sheetFormatPr)
  if (!sheetFormatPrXml) {
    return sheetXml
  }
  const existingPattern = /<sheetFormatPr\b[^>]*(?:\/>|>[\s\S]*?<\/sheetFormatPr>)/u
  if (existingPattern.test(sheetXml)) {
    return sheetXml.replace(existingPattern, sheetFormatPrXml)
  }
  const insertPattern = /<cols\b|<sheetData\b|<\/worksheet>/u
  const match = insertPattern.exec(sheetXml)
  return match ? `${sheetXml.slice(0, match.index)}${sheetFormatPrXml}${sheetXml.slice(match.index)}` : sheetXml
}

function buildColumnsXml(columns: readonly ExportColumnMetadata[]): string {
  const columnXml = columns
    .map((column) => {
      const min = column.start + 1
      const max = column.start + column.count
      return [
        '<col',
        xmlAttribute('min', String(min)),
        xmlAttribute('max', String(max)),
        xmlAttribute('width', column.xlsxWidth !== undefined ? formatXmlNumber(column.xlsxWidth) : undefined),
        xmlAttribute('customWidth', column.customWidth !== undefined ? formatXmlBoolean(column.customWidth) : undefined),
        xmlAttribute('bestFit', column.bestFit !== undefined ? formatXmlBoolean(column.bestFit) : undefined),
        xmlAttribute('hidden', column.hidden !== undefined ? formatXmlBoolean(column.hidden) : undefined),
        xmlAttribute('outlineLevel', column.outlineLevel !== undefined ? formatXmlNumber(column.outlineLevel) : undefined),
        xmlAttribute('collapsed', column.collapsed !== undefined ? formatXmlBoolean(column.collapsed) : undefined),
        '/>',
      ].join('')
    })
    .join('')
  return `<cols>${columnXml}</cols>`
}

function applyColumnMetadata(sheetXml: string, columns: readonly ExportColumnMetadata[]): string {
  if (columns.length === 0) {
    return sheetXml
  }
  const existingColumns = parseExistingColumnMetadata(sheetXml).flatMap((column) => subtractColumnRanges(column, columns))
  const columnsXml = buildColumnsXml(
    [...columns, ...existingColumns].toSorted((left, right) => left.start - right.start || left.count - right.count),
  )
  const existingPattern = /<cols\b[^>]*(?:\/>|>[\s\S]*?<\/cols>)/u
  if (existingPattern.test(sheetXml)) {
    return sheetXml.replace(existingPattern, columnsXml)
  }
  const match = /<sheetData\b|<\/worksheet>/u.exec(sheetXml)
  return match ? `${sheetXml.slice(0, match.index)}${columnsXml}${sheetXml.slice(match.index)}` : sheetXml
}

function rowOpeningTagPattern(rowNumber: number): RegExp {
  return new RegExp(`<row\\b(?=[^>]*\\br="${String(rowNumber)}"(?:\\s|/|>))[^>]*>`, 'u')
}

function rowElementPattern(rowNumber: number): RegExp {
  return new RegExp(`<row\\b(?=[^>]*\\br="${String(rowNumber)}"(?:\\s|/|>))[^>]*>([\\s\\S]*?)<\\/row>`, 'u')
}

function readRowNumber(rowTag: string): number | null {
  const match = /\br="([0-9]+)"/u.exec(rowTag)
  if (!match) {
    return null
  }
  const rowNumber = Number(match[1])
  return Number.isSafeInteger(rowNumber) && rowNumber > 0 ? rowNumber : null
}

function clearManagedRowAttributes(rowTag: string): string {
  return ['ht', 'customHeight', 'hidden', 'outlineLevel', 'collapsed', 'thickTop', 'thickBot'].reduce(
    (tag, attribute) => removeXmlAttribute(tag, attribute),
    rowTag,
  )
}

function applyRowMetadata(rowTag: string, row: ExportRowMetadata): string {
  let nextTag = clearManagedRowAttributes(rowTag)
  const height = row.xlsxHeight ?? row.size
  if (height !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'ht', formatXmlNumber(height))
    if (!row.exact || row.customHeight !== undefined) {
      nextTag = setXmlAttribute(nextTag, 'customHeight', formatXmlBoolean(row.customHeight ?? true))
    }
  } else if (row.customHeight !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'customHeight', formatXmlBoolean(row.customHeight))
  }
  if (row.hidden !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'hidden', formatXmlBoolean(row.hidden))
  }
  if (row.outlineLevel !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'outlineLevel', formatXmlNumber(row.outlineLevel))
  }
  if (row.collapsed !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'collapsed', formatXmlBoolean(row.collapsed))
  }
  if (row.thickTop !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'thickTop', formatXmlBoolean(row.thickTop))
  }
  if (row.thickBottom !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'thickBot', formatXmlBoolean(row.thickBottom))
  }
  return nextTag
}

function buildEmptyRowXml(row: ExportRowMetadata): string {
  let rowTag = `<row r="${escapeXmlAttribute(String(row.rowNumber))}"/>`
  rowTag = applyRowMetadata(rowTag, row)
  return rowTag
}

function insertRowIntoSheetData(sheetDataXml: string, rowXml: string, rowNumber: number): string {
  const closeIndex = sheetDataXml.lastIndexOf('</sheetData>')
  if (closeIndex < 0) {
    return sheetDataXml
  }
  let insertIndex = closeIndex
  for (const match of sheetDataXml.matchAll(/<row\b[^>]*>/gu)) {
    const existingRowNumber = readRowNumber(match[0])
    if (existingRowNumber !== null && existingRowNumber > rowNumber) {
      insertIndex = match.index
      break
    }
  }
  return `${sheetDataXml.slice(0, insertIndex)}${rowXml}${sheetDataXml.slice(insertIndex)}`
}

function upsertWorksheetRowMetadata(sheetXml: string, rows: readonly ExportRowMetadata[]): string {
  let nextXml = sheetXml
  for (const row of rows) {
    const existingRowElementPattern = rowElementPattern(row.rowNumber)
    const existingRowElementMatch = existingRowElementPattern.exec(nextXml)
    if (existingRowElementMatch) {
      const rowBody = existingRowElementMatch[1] ?? ''
      if (rowBody.trim().length === 0) {
        nextXml = nextXml.replace(existingRowElementPattern, () => buildEmptyRowXml(row))
        continue
      }
      nextXml = nextXml.replace(rowOpeningTagPattern(row.rowNumber), (rowTag) => applyRowMetadata(rowTag, row))
      continue
    }

    const existingRowPattern = rowOpeningTagPattern(row.rowNumber)
    if (existingRowPattern.test(nextXml)) {
      nextXml = nextXml.replace(existingRowPattern, (rowTag) => applyRowMetadata(rowTag, row))
      continue
    }
    const rowXml = buildEmptyRowXml(row)
    if (/<sheetData\b[^>]*\/>/u.test(nextXml)) {
      nextXml = nextXml.replace(/<sheetData\b([^>]*)\/>/u, (_match, attributes: string) => `<sheetData${attributes}>${rowXml}</sheetData>`)
      continue
    }
    const sheetDataMatch = /<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/u.exec(nextXml)
    if (!sheetDataMatch) {
      continue
    }
    const sheetDataXml = sheetDataMatch[0]
    const updatedSheetDataXml = insertRowIntoSheetData(sheetDataXml, rowXml, row.rowNumber)
    nextXml = `${nextXml.slice(0, sheetDataMatch.index)}${updatedSheetDataXml}${nextXml.slice(sheetDataMatch.index + sheetDataXml.length)}`
  }
  return nextXml
}

export function hasExportWorksheetDimensions(snapshot: WorkbookSnapshot): boolean {
  return snapshot.sheets.some((sheet) => {
    const metadata = sheet.metadata
    return (
      metadata?.sheetFormatPr !== undefined ||
      normalizeExportColumnMetadata(metadata?.columnMetadata).length > 0 ||
      normalizeExportRowMetadata(metadata?.rows, metadata?.rowMetadata).length > 0
    )
  })
}

export function applyExportWorksheetDimensionsToWorksheetXml(sheetXml: string, metadata: SheetMetadataSnapshot | undefined): string {
  const sheetFormatXml = applySheetFormatPr(sheetXml, metadata?.sheetFormatPr)
  const columnXml = applyColumnMetadata(sheetFormatXml, normalizeExportColumnMetadata(metadata?.columnMetadata))
  const normalizedRows = normalizeExportRowMetadata(metadata?.rows, metadata?.rowMetadata)
  return normalizedRows.length > 0 ? upsertWorksheetRowMetadata(columnXml, normalizedRows) : columnXml
}

export function addExportWorksheetDimensionsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!hasExportWorksheetDimensions(snapshot)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const updatedSheetXml = applyExportWorksheetDimensionsToWorksheetXml(sheetXml, sheet.metadata)
      if (updatedSheetXml === sheetXml) {
        return
      }
      setZipText(zip, sheetPath, updatedSheetXml)
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}
