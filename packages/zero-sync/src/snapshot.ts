import { formatAddress } from '@bilig/formula'
import type {
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellNumberFormatRecord,
  CellStyleRecord,
  CellVerticalAlignment,
  CompatibilityMode,
  SheetFormatRangeSnapshot,
  SheetMetadataSnapshot,
  SheetStyleRangeSnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookDefinedNameSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookPropertySnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { isWorkbookSnapshot } from '@bilig/protocol'
import { isLiteralInput } from './mutators.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined
}

function asNonNegativeNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : undefined
}

function asFiniteNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined
}

function asSafeNonNegativeInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 ? value : undefined
}

function asSafePositiveInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0 ? value : undefined
}

function asBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

function asArray(value: unknown): unknown[] {
  return Array.isArray(value) ? value : []
}

function isCellNumberFormatKind(value: unknown): value is CellNumberFormatRecord['kind'] {
  return (
    value === 'general' ||
    value === 'number' ||
    value === 'currency' ||
    value === 'accounting' ||
    value === 'percent' ||
    value === 'date' ||
    value === 'time' ||
    value === 'datetime' ||
    value === 'text'
  )
}

function isCompatibilityMode(value: unknown): value is CompatibilityMode {
  return value === 'excel-modern' || value === 'odf-1.4'
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

export function createEmptyWorkbookSnapshot(documentId: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: documentId,
    },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        cells: [],
      },
    ],
  }
}

function parseAxisMetadata(entries: unknown[]): WorkbookAxisMetadataSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const start = asSafeNonNegativeInteger(entry['startIndex'])
      const count = asSafePositiveInteger(entry['count'])
      if (start === undefined || count === undefined) {
        return null
      }
      const next: WorkbookAxisMetadataSnapshot = {
        start,
        count,
      }
      const size = asNonNegativeNumber(entry['size'])
      const hiddenFlag = asBoolean(entry['hidden'])
      if (size !== undefined) {
        next.size = size
      }
      if (hiddenFlag !== undefined) {
        next.hidden = hiddenFlag
      }
      return next
    })
    .filter((entry): entry is WorkbookAxisMetadataSnapshot => entry !== null)
}

function mergeAxisMetadataEntries(
  primary: readonly WorkbookAxisMetadataSnapshot[],
  fallback: readonly WorkbookAxisMetadataSnapshot[] | undefined,
): WorkbookAxisMetadataSnapshot[] {
  if (!fallback || fallback.length === 0) {
    return [...primary]
  }
  const fallbackByKey = new Map(fallback.map((entry) => [`${String(entry.start)}:${String(entry.count)}`, entry]))
  return primary.map((entry) => {
    const preserved = fallbackByKey.get(`${String(entry.start)}:${String(entry.count)}`)
    if (!preserved) {
      return entry
    }
    return {
      ...preserved,
      ...entry,
    }
  })
}

function parseWorkbookProperties(entries: unknown[]): WorkbookPropertySnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const key = asString(entry['key'])
      const value = entry['value']
      if (!key || !isLiteralInput(value)) {
        return null
      }
      return { key, value }
    })
    .filter((entry): entry is WorkbookPropertySnapshot => entry !== null)
}

function isWorkbookDefinedNameValueSnapshot(value: unknown): value is WorkbookDefinedNameValueSnapshot {
  if (isLiteralInput(value)) {
    return true
  }
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'scalar':
      return isLiteralInput(value['value'])
    case 'cell-ref':
      return typeof value['sheetName'] === 'string' && typeof value['address'] === 'string'
    case 'range-ref':
      return typeof value['sheetName'] === 'string' && typeof value['startAddress'] === 'string' && typeof value['endAddress'] === 'string'
    case 'structured-ref':
      return typeof value['tableName'] === 'string' && typeof value['columnName'] === 'string'
    case 'formula':
      return typeof value['formula'] === 'string'
    default:
      return false
  }
}

function parseDefinedNames(entries: unknown[]): WorkbookDefinedNameSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const name = asString(entry['name'])
      const value = entry['value']
      if (!name || !isWorkbookDefinedNameValueSnapshot(value)) {
        return null
      }
      return { name, value }
    })
    .filter((entry): entry is WorkbookDefinedNameSnapshot => entry !== null)
}

function parseStyleFill(value: unknown): CellStyleRecord['fill'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const backgroundColor = asString(value['backgroundColor'])
  return backgroundColor ? { backgroundColor } : undefined
}

function parseStyleFont(value: unknown): CellStyleRecord['font'] | undefined {
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

function parseStyleAlignment(value: unknown): CellStyleRecord['alignment'] | undefined {
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

function parseBorderSide(value: unknown): CellBorderSideSnapshot | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const style = asCellBorderStyle(value['style'])
  const weight = asCellBorderWeight(value['weight'])
  const color = asString(value['color'])
  if (!style || !weight || !color) {
    return undefined
  }
  return {
    style,
    weight,
    color,
  }
}

function parseStyleBorders(value: unknown): CellStyleRecord['borders'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const borders: NonNullable<CellStyleRecord['borders']> = {}
  const top = parseBorderSide(value['top'])
  const right = parseBorderSide(value['right'])
  const bottom = parseBorderSide(value['bottom'])
  const left = parseBorderSide(value['left'])
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

function parseStyleProtection(value: unknown): CellStyleRecord['protection'] | undefined {
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

function parseStyleRecord(id: string, value: unknown): CellStyleRecord | null {
  if (!isRecord(value)) {
    return null
  }
  const style: CellStyleRecord = { id }
  const fill = parseStyleFill(value['fill'])
  const font = parseStyleFont(value['font'])
  const alignment = parseStyleAlignment(value['alignment'])
  const borders = parseStyleBorders(value['borders'])
  const protection = parseStyleProtection(value['protection'])
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

function parseStyleRecords(entries: unknown[]): CellStyleRecord[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const id = asString(entry['id'])
      const recordJSON = entry['recordJSON']
      if (!id || !isRecord(recordJSON)) {
        return null
      }
      return parseStyleRecord(id, recordJSON)
    })
    .filter((entry): entry is CellStyleRecord => entry !== null)
}

function parseNumberFormats(entries: unknown[]): CellNumberFormatRecord[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const id = asString(entry['id'])
      const code = asString(entry['code'])
      const kind = asString(entry['kind'])
      if (!id || !code || !isCellNumberFormatKind(kind)) {
        return null
      }
      return {
        id,
        code,
        kind,
      }
    })
    .filter((entry): entry is CellNumberFormatRecord => entry !== null)
}

function parseFreezePane(
  freezeRows: unknown,
  freezeCols: unknown,
  fallback?: WorkbookFreezePaneSnapshot,
): WorkbookFreezePaneSnapshot | undefined {
  const rows = asSafeNonNegativeInteger(freezeRows)
  const cols = asSafeNonNegativeInteger(freezeCols)
  if ((rows ?? 0) > 0 || (cols ?? 0) > 0) {
    return {
      rows: rows ?? 0,
      cols: cols ?? 0,
    }
  }
  return fallback
}

function parseStyleRanges(entries: unknown[]): SheetStyleRangeSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const startRow = asSafeNonNegativeInteger(entry['startRow'])
      const endRow = asSafeNonNegativeInteger(entry['endRow'])
      const startCol = asSafeNonNegativeInteger(entry['startCol'])
      const endCol = asSafeNonNegativeInteger(entry['endCol'])
      const styleId = asString(entry['styleId'])
      if (
        startRow === undefined ||
        endRow === undefined ||
        startCol === undefined ||
        endCol === undefined ||
        endRow < startRow ||
        endCol < startCol ||
        !styleId
      ) {
        return null
      }
      return {
        range: {
          sheetName: '',
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        styleId,
      }
    })
    .filter((entry): entry is SheetStyleRangeSnapshot => entry !== null)
}

function parseFormatRanges(entries: unknown[]): SheetFormatRangeSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const startRow = asSafeNonNegativeInteger(entry['startRow'])
      const endRow = asSafeNonNegativeInteger(entry['endRow'])
      const startCol = asSafeNonNegativeInteger(entry['startCol'])
      const endCol = asSafeNonNegativeInteger(entry['endCol'])
      const formatId = asString(entry['formatId'])
      if (
        startRow === undefined ||
        endRow === undefined ||
        startCol === undefined ||
        endCol === undefined ||
        endRow < startRow ||
        endCol < startCol ||
        !formatId
      ) {
        return null
      }
      return {
        range: {
          sheetName: '',
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        formatId,
      }
    })
    .filter((entry): entry is SheetFormatRangeSnapshot => entry !== null)
}

function withSheetMetadataFallback(
  sheetName: string,
  rowEntries: WorkbookAxisMetadataSnapshot[],
  columnEntries: WorkbookAxisMetadataSnapshot[],
  styleRanges: SheetStyleRangeSnapshot[],
  formatRanges: SheetFormatRangeSnapshot[],
  freezePane: WorkbookFreezePaneSnapshot | undefined,
  fallback?: SheetMetadataSnapshot,
) {
  const next: SheetMetadataSnapshot = {}
  if (fallback?.rows) {
    next.rows = fallback.rows
  }
  if (fallback?.columns) {
    next.columns = fallback.columns
  }
  if (fallback?.filters) {
    next.filters = fallback.filters
  }
  if (fallback?.sorts) {
    next.sorts = fallback.sorts
  }
  if (fallback?.ignoredErrors) {
    next.ignoredErrors = fallback.ignoredErrors
  }
  if (fallback?.sparklines) {
    next.sparklines = fallback.sparklines
  }
  if (fallback?.conditionalFormatArtifacts) {
    next.conditionalFormatArtifacts = fallback.conditionalFormatArtifacts
  }
  if (fallback?.styleArtifacts) {
    next.styleArtifacts = fallback.styleArtifacts
  }
  if (fallback?.pivotArtifacts) {
    next.pivotArtifacts = fallback.pivotArtifacts
  }
  if (fallback?.richTextArtifacts) {
    next.richTextArtifacts = fallback.richTextArtifacts
  }
  if (fallback?.threadedCommentArtifacts) {
    next.threadedCommentArtifacts = fallback.threadedCommentArtifacts
  }
  if (fallback?.viewState) {
    next.viewState = fallback.viewState
  }
  if (fallback?.printPageSetup) {
    next.printPageSetup = fallback.printPageSetup
  }
  if (fallback?.merges) {
    next.merges = fallback.merges
  }
  if (rowEntries.length > 0) {
    next.rowMetadata = mergeAxisMetadataEntries(rowEntries, fallback?.rowMetadata)
  } else if (fallback?.rowMetadata) {
    next.rowMetadata = fallback.rowMetadata
  }
  if (columnEntries.length > 0) {
    next.columnMetadata = mergeAxisMetadataEntries(columnEntries, fallback?.columnMetadata)
  } else if (fallback?.columnMetadata) {
    next.columnMetadata = fallback.columnMetadata
  }
  if (styleRanges.length > 0) {
    next.styleRanges = styleRanges.map((entry) => ({
      ...entry,
      range: {
        ...entry.range,
        sheetName,
      },
    }))
  } else if (fallback?.styleRanges) {
    next.styleRanges = fallback.styleRanges
  }
  if (formatRanges.length > 0) {
    next.formatRanges = formatRanges.map((entry) => ({
      ...entry,
      range: {
        ...entry.range,
        sheetName,
      },
    }))
  } else if (fallback?.formatRanges) {
    next.formatRanges = fallback.formatRanges
  }
  if (freezePane) {
    next.freezePane = freezePane
  } else if (fallback?.freezePane) {
    next.freezePane = fallback.freezePane
  }
  return Object.keys(next).length > 0 ? next : undefined
}

export function projectWorkbookToSnapshot(value: unknown, documentId: string) {
  if (!isRecord(value)) {
    return null
  }

  const baseSnapshot = isWorkbookSnapshot(value['snapshot']) ? value['snapshot'] : createEmptyWorkbookSnapshot(documentId)
  const workbookName = asString(value['name']) ?? baseSnapshot.workbook.name ?? documentId

  const workbookMetadata = parseWorkbookProperties(asArray(value['workbookMetadataEntries']))
  const definedNames = parseDefinedNames(asArray(value['definedNames']))
  const styles = parseStyleRecords(asArray(value['styles']))
  const numberFormats = parseNumberFormats(asArray(value['numberFormats']))
  const numberFormatCodeById = new Map(numberFormats.map((entry) => [entry.id, entry.code]))

  const calculationSettingsRecord = isRecord(value['calculationSettings']) ? value['calculationSettings'] : null
  const calculationMode = calculationSettingsRecord ? asString(calculationSettingsRecord['mode']) : undefined
  const compatibilityMode = asString(value['compatibilityMode'])
  const recalcEpoch =
    calculationSettingsRecord?.['recalcEpoch'] !== undefined
      ? asSafeNonNegativeInteger(calculationSettingsRecord['recalcEpoch'])
      : asSafeNonNegativeInteger(value['recalcEpoch'])

  const fallbackSheets = new Map(baseSnapshot.sheets.map((sheet) => [sheet.name, sheet]))
  const projectedSheets = asArray(value['sheets'])
    .map((sheetEntry) => {
      if (!isRecord(sheetEntry)) {
        return null
      }
      const sheetName = asString(sheetEntry['name'])
      const sortOrder = asSafeNonNegativeInteger(sheetEntry['sortOrder'])
      if (!sheetName || sortOrder === undefined) {
        return null
      }

      const cells = asArray(sheetEntry['cells'])
        .map((cellEntry) => {
          if (!isRecord(cellEntry)) {
            return null
          }
          const explicitFormatId = asString(cellEntry['explicitFormatId'])
          const rowNum = asSafeNonNegativeInteger(cellEntry['rowNum'])
          const colNum = asSafeNonNegativeInteger(cellEntry['colNum'])
          const address =
            asString(cellEntry['address']) ?? (rowNum !== undefined && colNum !== undefined ? formatAddress(rowNum, colNum) : undefined)
          if (!address) {
            return null
          }
          const inputValue = cellEntry['inputValue']
          const formula = asString(cellEntry['formula'])
          const format = asString(cellEntry['format']) ?? (explicitFormatId ? numberFormatCodeById.get(explicitFormatId) : undefined)
          const nextCell: WorkbookSnapshot['sheets'][number]['cells'][number] = { address }
          if (formula) {
            nextCell.formula = formula
          } else if (isLiteralInput(inputValue)) {
            nextCell.value = inputValue
          }
          if (format) {
            nextCell.format = format
          }
          return nextCell
        })
        .filter((entry): entry is WorkbookSnapshot['sheets'][number]['cells'][number] => entry !== null)

      const fallbackSheet = fallbackSheets.get(sheetName)
      const metadata = withSheetMetadataFallback(
        sheetName,
        parseAxisMetadata(asArray(sheetEntry['rowMetadata'])),
        parseAxisMetadata(asArray(sheetEntry['columnMetadata'])),
        parseStyleRanges(asArray(sheetEntry['styleRanges'])),
        parseFormatRanges(asArray(sheetEntry['formatRanges'])),
        parseFreezePane(sheetEntry['freezeRows'], sheetEntry['freezeCols'], fallbackSheet?.metadata?.freezePane),
        fallbackSheet?.metadata,
      )

      const id = asSafeNonNegativeInteger(sheetEntry['id']) ?? fallbackSheet?.id
      const nextSheet: WorkbookSnapshot['sheets'][number] = metadata
        ? { name: sheetName, order: sortOrder, metadata, cells }
        : { name: sheetName, order: sortOrder, cells }
      if (id !== undefined) {
        nextSheet.id = id
      }
      return nextSheet
    })
    .filter((entry): entry is WorkbookSnapshot['sheets'][number] => entry !== null)

  const workbookMetadataSnapshot = {
    ...baseSnapshot.workbook.metadata,
  }

  if (workbookMetadata.length > 0) {
    workbookMetadataSnapshot.properties = workbookMetadata
  }
  if (definedNames.length > 0) {
    workbookMetadataSnapshot.definedNames = definedNames
  }
  if (styles.length > 0) {
    workbookMetadataSnapshot.styles = styles
  }
  if (numberFormats.length > 0) {
    workbookMetadataSnapshot.formats = numberFormats
  }
  if ((calculationMode === 'automatic' || calculationMode === 'manual') && isCompatibilityMode(compatibilityMode)) {
    workbookMetadataSnapshot.calculationSettings = {
      ...baseSnapshot.workbook.metadata?.calculationSettings,
      mode: calculationMode,
      compatibilityMode,
    }
  } else if (calculationMode === 'automatic' || calculationMode === 'manual') {
    workbookMetadataSnapshot.calculationSettings = {
      ...baseSnapshot.workbook.metadata?.calculationSettings,
      mode: calculationMode,
    }
  }
  if (recalcEpoch !== undefined) {
    workbookMetadataSnapshot.volatileContext = {
      recalcEpoch,
    }
  }

  const workbook =
    Object.keys(workbookMetadataSnapshot).length > 0 ? { name: workbookName, metadata: workbookMetadataSnapshot } : { name: workbookName }

  return {
    version: 1,
    workbook,
    sheets: projectedSheets.length > 0 ? projectedSheets : baseSnapshot.sheets,
  }
}
