import { isLiteralInput, type CellRangeRef } from '@bilig/protocol'
import type { EngineOp, EngineOpBatch, WorkbookOp } from './ops.js'

const HORIZONTAL_ALIGNMENT_VALUES = new Set(['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'])
const VERTICAL_ALIGNMENT_VALUES = new Set(['top', 'middle', 'bottom', 'justify', 'distributed'])
const BORDER_STYLE_VALUES = new Set(['solid', 'dashed', 'dotted', 'double'])
const BORDER_WEIGHT_VALUES = new Set(['thin', 'medium', 'thick'])
const NUMBER_FORMAT_KIND_VALUES = new Set(['general', 'number', 'currency', 'accounting', 'percent', 'date', 'time', 'datetime', 'text'])
const COMPATIBILITY_MODE_VALUES = new Set(['excel-modern', 'odf-1.4'])
const SORT_DIRECTION_VALUES = new Set(['asc', 'desc'])
const PIVOT_AGGREGATION_VALUES = new Set(['sum', 'count'])
const VALIDATION_COMPARISON_OPERATOR_VALUES = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])
const VALIDATION_ERROR_STYLE_VALUES = new Set(['stop', 'warning', 'information'])

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function ownValue(value: Record<string, unknown>, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return isRecord(value) && hasString(value, 'sheetName') && hasString(value, 'startAddress') && hasString(value, 'endAddress')
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isSafePositiveInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0
}

function isStringArray(value: unknown): value is string[] {
  return arrayEvery(value, (entry) => typeof entry === 'string')
}

function stringSetHas(values: ReadonlySet<string>, value: unknown): boolean {
  return typeof value === 'string' && values.has(value)
}

function arrayEvery(value: unknown, predicate: (entry: unknown) => boolean): boolean {
  if (!Array.isArray(value)) {
    return false
  }

  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !('value' in descriptor) || !predicate(descriptor.value)) {
      return false
    }
  }

  return true
}

function nonEmptyArrayEvery(value: unknown, predicate: (entry: unknown) => boolean): boolean {
  return Array.isArray(value) && value.length > 0 && arrayEvery(value, predicate)
}

function isConditionalFormatArtifacts(value: unknown): boolean {
  return isRecord(value) && hasString(value, 'xml')
}

function isOptionalString(value: unknown): value is string | undefined {
  return value === undefined || typeof value === 'string'
}

function isOptionalNumber(value: unknown): value is number | undefined {
  return value === undefined || isFiniteNumber(value)
}

function isOptionalSafePositiveInteger(value: unknown): value is number | undefined {
  return value === undefined || isSafePositiveInteger(value)
}

function isOptionalSafeNonNegativeInteger(value: unknown): value is number | undefined {
  return value === undefined || isSafeNonNegativeInteger(value)
}

function isOptionalNullableNumber(value: unknown): value is number | null | undefined {
  return value === undefined || value === null || isFiniteNumber(value)
}

function isOptionalNullableSafePositiveInteger(value: unknown): value is number | null | undefined {
  return value === undefined || value === null || isSafePositiveInteger(value)
}

function isOptionalBoolean(value: unknown): value is boolean | undefined {
  return value === undefined || typeof value === 'boolean'
}

function isOptionalLiteralInput(value: unknown): boolean {
  return value === undefined || isLiteralInput(value)
}

function isOptionalNullableBoolean(value: unknown): value is boolean | null | undefined {
  return value === undefined || value === null || typeof value === 'boolean'
}

function hasString(value: Record<string, unknown>, key: string): boolean {
  return typeof ownValue(value, key) === 'string'
}

function hasFiniteNumber(value: Record<string, unknown>, key: string): boolean {
  return isFiniteNumber(ownValue(value, key))
}

function hasSafeNonNegativeInteger(value: Record<string, unknown>, key: string): boolean {
  return isSafeNonNegativeInteger(ownValue(value, key))
}

function hasSafePositiveInteger(value: Record<string, unknown>, key: string): boolean {
  return isSafePositiveInteger(ownValue(value, key))
}

function isWorkbookAxisEntry(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    hasSafeNonNegativeInteger(value, 'index') &&
    isOptionalNullableSafePositiveInteger(ownValue(value, 'size')) &&
    isOptionalNullableBoolean(ownValue(value, 'hidden')) &&
    isOptionalNullableBoolean(ownValue(value, 'filterHidden'))
  )
}

function isCellBorderSide(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'color') &&
    stringSetHas(BORDER_STYLE_VALUES, ownValue(value, 'style')) &&
    stringSetHas(BORDER_WEIGHT_VALUES, ownValue(value, 'weight'))
  )
}

function isCellStyleRecord(value: unknown): boolean {
  if (!isRecord(value) || !hasString(value, 'id')) {
    return false
  }

  const fill = ownValue(value, 'fill')
  if (fill !== undefined && (!isRecord(fill) || typeof ownValue(fill, 'backgroundColor') !== 'string')) {
    return false
  }

  const font = ownValue(value, 'font')
  if (
    font !== undefined &&
    (!isRecord(font) ||
      !isOptionalString(ownValue(font, 'family')) ||
      !isOptionalNumber(ownValue(font, 'size')) ||
      !isOptionalBoolean(ownValue(font, 'bold')) ||
      !isOptionalBoolean(ownValue(font, 'italic')) ||
      !isOptionalBoolean(ownValue(font, 'underline')) ||
      !isOptionalString(ownValue(font, 'color')))
  ) {
    return false
  }

  const alignment = ownValue(value, 'alignment')
  if (
    alignment !== undefined &&
    (!isRecord(alignment) ||
      !(ownValue(alignment, 'horizontal') === undefined || stringSetHas(HORIZONTAL_ALIGNMENT_VALUES, ownValue(alignment, 'horizontal'))) ||
      !(ownValue(alignment, 'vertical') === undefined || stringSetHas(VERTICAL_ALIGNMENT_VALUES, ownValue(alignment, 'vertical'))) ||
      !isOptionalBoolean(ownValue(alignment, 'wrap')) ||
      !isOptionalNumber(ownValue(alignment, 'indent')) ||
      !isOptionalBoolean(ownValue(alignment, 'shrinkToFit')) ||
      !isOptionalNumber(ownValue(alignment, 'readingOrder')) ||
      !isOptionalNumber(ownValue(alignment, 'textRotation')) ||
      !isOptionalBoolean(ownValue(alignment, 'justifyLastLine')))
  ) {
    return false
  }

  const borders = ownValue(value, 'borders')
  if (
    borders !== undefined &&
    (!isRecord(borders) ||
      !(ownValue(borders, 'top') === undefined || isCellBorderSide(ownValue(borders, 'top'))) ||
      !(ownValue(borders, 'right') === undefined || isCellBorderSide(ownValue(borders, 'right'))) ||
      !(ownValue(borders, 'bottom') === undefined || isCellBorderSide(ownValue(borders, 'bottom'))) ||
      !(ownValue(borders, 'left') === undefined || isCellBorderSide(ownValue(borders, 'left'))))
  ) {
    return false
  }

  const protection = ownValue(value, 'protection')
  if (
    protection !== undefined &&
    (!isRecord(protection) || !isOptionalBoolean(ownValue(protection, 'locked')) || !isOptionalBoolean(ownValue(protection, 'hidden')))
  ) {
    return false
  }

  return true
}

function isCellNumberFormatRecord(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    hasString(value, 'code') &&
    stringSetHas(NUMBER_FORMAT_KIND_VALUES, ownValue(value, 'kind'))
  )
}

function isWorkbookCalculationSettings(value: unknown): boolean {
  return (
    isRecord(value) &&
    (ownValue(value, 'mode') === 'automatic' || ownValue(value, 'mode') === 'manual') &&
    (ownValue(value, 'compatibilityMode') === undefined || stringSetHas(COMPATIBILITY_MODE_VALUES, ownValue(value, 'compatibilityMode')))
  )
}

function isWorkbookVolatileContext(value: unknown): boolean {
  return isRecord(value) && hasFiniteNumber(value, 'recalcEpoch')
}

function isWorkbookSortKey(value: unknown): boolean {
  return isRecord(value) && hasString(value, 'keyAddress') && stringSetHas(SORT_DIRECTION_VALUES, ownValue(value, 'direction'))
}

function isWorkbookValidationListSource(value: unknown): boolean {
  if (!isRecord(value) || typeof ownValue(value, 'kind') !== 'string') {
    return false
  }
  switch (ownValue(value, 'kind')) {
    case 'named-range':
      return hasString(value, 'name')
    case 'cell-ref':
      return hasString(value, 'sheetName') && hasString(value, 'address')
    case 'range-ref':
      return isCellRangeRef(value)
    case 'structured-ref':
      return hasString(value, 'tableName') && hasString(value, 'columnName')
    case 'formula':
      return hasString(value, 'formula')
    default:
      return false
  }
}

function isWorkbookDataValidationRule(value: unknown): boolean {
  if (!isRecord(value) || typeof ownValue(value, 'kind') !== 'string') {
    return false
  }
  switch (ownValue(value, 'kind')) {
    case 'list': {
      const hasValues = arrayEvery(ownValue(value, 'values'), isLiteralInput)
      const hasSource = ownValue(value, 'source') !== undefined && isWorkbookValidationListSource(ownValue(value, 'source'))
      return (hasValues ? 1 : 0) + (hasSource ? 1 : 0) === 1
    }
    case 'checkbox':
      return isOptionalLiteralInput(ownValue(value, 'checkedValue')) && isOptionalLiteralInput(ownValue(value, 'uncheckedValue'))
    case 'any':
      return true
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength':
      const operator = ownValue(value, 'operator')
      const values = ownValue(value, 'values')
      return (
        stringSetHas(VALIDATION_COMPARISON_OPERATOR_VALUES, operator) &&
        arrayEvery(values, isLiteralInput) &&
        (operator === 'between' || operator === 'notBetween'
          ? Array.isArray(values) && values.length === 2
          : Array.isArray(values) && values.length === 1)
      )
    default:
      return false
  }
}

function isWorkbookDataValidation(value: unknown): boolean {
  return (
    isRecord(value) &&
    isCellRangeRef(ownValue(value, 'range')) &&
    isWorkbookDataValidationRule(ownValue(value, 'rule')) &&
    isOptionalBoolean(ownValue(value, 'allowBlank')) &&
    isOptionalBoolean(ownValue(value, 'showDropdown')) &&
    isOptionalString(ownValue(value, 'promptTitle')) &&
    isOptionalString(ownValue(value, 'promptMessage')) &&
    (ownValue(value, 'errorStyle') === undefined || stringSetHas(VALIDATION_ERROR_STYLE_VALUES, ownValue(value, 'errorStyle'))) &&
    isOptionalString(ownValue(value, 'errorTitle')) &&
    isOptionalString(ownValue(value, 'errorMessage'))
  )
}

function isCellStylePatch(value: unknown): boolean {
  if (!isRecord(value)) {
    return false
  }

  const fill = ownValue(value, 'fill')
  if (
    fill !== undefined &&
    fill !== null &&
    (!isRecord(fill) ||
      !(
        ownValue(fill, 'backgroundColor') === undefined ||
        ownValue(fill, 'backgroundColor') === null ||
        hasString(fill, 'backgroundColor')
      ))
  ) {
    return false
  }

  const font = ownValue(value, 'font')
  if (
    font !== undefined &&
    font !== null &&
    (!isRecord(font) ||
      !(ownValue(font, 'family') === undefined || ownValue(font, 'family') === null || hasString(font, 'family')) ||
      !isOptionalNullableNumber(ownValue(font, 'size')) ||
      !isOptionalNullableBoolean(ownValue(font, 'bold')) ||
      !isOptionalNullableBoolean(ownValue(font, 'italic')) ||
      !isOptionalNullableBoolean(ownValue(font, 'underline')) ||
      !(ownValue(font, 'color') === undefined || ownValue(font, 'color') === null || hasString(font, 'color')))
  ) {
    return false
  }

  const alignment = ownValue(value, 'alignment')
  if (
    alignment !== undefined &&
    alignment !== null &&
    (!isRecord(alignment) ||
      !(
        ownValue(alignment, 'horizontal') === undefined ||
        ownValue(alignment, 'horizontal') === null ||
        stringSetHas(HORIZONTAL_ALIGNMENT_VALUES, ownValue(alignment, 'horizontal'))
      ) ||
      !(
        ownValue(alignment, 'vertical') === undefined ||
        ownValue(alignment, 'vertical') === null ||
        stringSetHas(VERTICAL_ALIGNMENT_VALUES, ownValue(alignment, 'vertical'))
      ) ||
      !isOptionalNullableBoolean(ownValue(alignment, 'wrap')) ||
      !isOptionalNullableNumber(ownValue(alignment, 'indent')) ||
      !isOptionalNullableBoolean(ownValue(alignment, 'shrinkToFit')) ||
      !isOptionalNullableNumber(ownValue(alignment, 'readingOrder')) ||
      !isOptionalNullableNumber(ownValue(alignment, 'textRotation')) ||
      !isOptionalNullableBoolean(ownValue(alignment, 'justifyLastLine')))
  ) {
    return false
  }

  const borders = ownValue(value, 'borders')
  if (borders !== undefined && borders !== null) {
    if (!isRecord(borders)) {
      return false
    }
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const sideValue = ownValue(borders, side)
      if (sideValue === undefined || sideValue === null) {
        continue
      }
      if (
        !isRecord(sideValue) ||
        !(
          ownValue(sideValue, 'style') === undefined ||
          ownValue(sideValue, 'style') === null ||
          stringSetHas(BORDER_STYLE_VALUES, ownValue(sideValue, 'style'))
        ) ||
        !(
          ownValue(sideValue, 'weight') === undefined ||
          ownValue(sideValue, 'weight') === null ||
          stringSetHas(BORDER_WEIGHT_VALUES, ownValue(sideValue, 'weight'))
        ) ||
        !(ownValue(sideValue, 'color') === undefined || ownValue(sideValue, 'color') === null || hasString(sideValue, 'color'))
      ) {
        return false
      }
    }
  }

  return true
}

function isWorkbookConditionalFormatRule(value: unknown): boolean {
  if (!isRecord(value) || typeof ownValue(value, 'kind') !== 'string') {
    return false
  }
  switch (ownValue(value, 'kind')) {
    case 'cellIs':
      const operator = ownValue(value, 'operator')
      const values = ownValue(value, 'values')
      return (
        stringSetHas(VALIDATION_COMPARISON_OPERATOR_VALUES, operator) &&
        arrayEvery(values, isLiteralInput) &&
        (operator === 'between' || operator === 'notBetween'
          ? Array.isArray(values) && values.length === 2
          : Array.isArray(values) && values.length === 1)
      )
    case 'textContains':
      return hasString(value, 'text') && isOptionalBoolean(ownValue(value, 'caseSensitive'))
    case 'formula':
      return hasString(value, 'formula')
    case 'blanks':
    case 'notBlanks':
      return true
    default:
      return false
  }
}

function isWorkbookConditionalFormat(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    isCellRangeRef(ownValue(value, 'range')) &&
    isWorkbookConditionalFormatRule(ownValue(value, 'rule')) &&
    isCellStylePatch(ownValue(value, 'style')) &&
    isOptionalBoolean(ownValue(value, 'stopIfTrue')) &&
    isOptionalSafeNonNegativeInteger(ownValue(value, 'priority'))
  )
}

function isWorkbookSheetProtection(value: unknown): boolean {
  return isRecord(value) && hasString(value, 'sheetName') && isOptionalBoolean(ownValue(value, 'hideFormulas'))
}

function isWorkbookRangeProtection(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    isCellRangeRef(ownValue(value, 'range')) &&
    isOptionalBoolean(ownValue(value, 'hideFormulas'))
  )
}

function isWorkbookCommentEntry(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    hasString(value, 'body') &&
    isOptionalString(ownValue(value, 'authorUserId')) &&
    isOptionalString(ownValue(value, 'authorDisplayName')) &&
    isOptionalSafeNonNegativeInteger(ownValue(value, 'createdAtUnixMs'))
  )
}

function isWorkbookCommentThread(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'threadId') &&
    hasString(value, 'sheetName') &&
    hasString(value, 'address') &&
    nonEmptyArrayEvery(ownValue(value, 'comments'), isWorkbookCommentEntry) &&
    isOptionalBoolean(ownValue(value, 'resolved')) &&
    isOptionalString(ownValue(value, 'resolvedByUserId')) &&
    isOptionalSafeNonNegativeInteger(ownValue(value, 'resolvedAtUnixMs'))
  )
}

function isWorkbookNote(value: unknown): boolean {
  return isRecord(value) && hasString(value, 'sheetName') && hasString(value, 'address') && hasString(value, 'text')
}

function isWorkbookHyperlink(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'sheetName') &&
    hasString(value, 'address') &&
    hasString(value, 'target') &&
    isOptionalString(ownValue(value, 'tooltip')) &&
    isOptionalString(ownValue(value, 'display'))
  )
}

function isWorkbookDefinedNameValue(value: unknown): boolean {
  if (isLiteralInput(value)) {
    return true
  }

  if (!isRecord(value) || typeof ownValue(value, 'kind') !== 'string') {
    return false
  }

  switch (ownValue(value, 'kind')) {
    case 'scalar':
      return isLiteralInput(ownValue(value, 'value'))
    case 'cell-ref':
      return hasString(value, 'sheetName') && hasString(value, 'address')
    case 'range-ref':
      return isCellRangeRef(value)
    case 'structured-ref':
      return hasString(value, 'tableName') && hasString(value, 'columnName')
    case 'formula':
      return hasString(value, 'formula')
    default:
      return false
  }
}

function isWorkbookTableOp(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'name') &&
    hasString(value, 'sheetName') &&
    hasString(value, 'startAddress') &&
    hasString(value, 'endAddress') &&
    isStringArray(ownValue(value, 'columnNames')) &&
    typeof ownValue(value, 'headerRow') === 'boolean' &&
    typeof ownValue(value, 'totalsRow') === 'boolean' &&
    (ownValue(value, 'columns') === undefined || isWorkbookTableColumns(ownValue(value, 'columns'))) &&
    (ownValue(value, 'style') === undefined || isWorkbookTableStyle(ownValue(value, 'style'))) &&
    (ownValue(value, 'autoFilter') === undefined || isCellRangeRef(ownValue(value, 'autoFilter'))) &&
    isOptionalString(ownValue(value, 'sortState'))
  )
}

function isWorkbookTableColumns(value: unknown): boolean {
  return arrayEvery(
    value,
    (entry) =>
      isRecord(entry) &&
      hasString(entry, 'name') &&
      isOptionalString(ownValue(entry, 'calculatedColumnFormula')) &&
      isOptionalString(ownValue(entry, 'totalsRowLabel')) &&
      isOptionalString(ownValue(entry, 'totalsRowFunction')) &&
      isOptionalString(ownValue(entry, 'totalsRowFormula')),
  )
}

function isWorkbookTableStyle(value: unknown): boolean {
  return (
    isRecord(value) &&
    isOptionalString(ownValue(value, 'name')) &&
    isOptionalBoolean(ownValue(value, 'showFirstColumn')) &&
    isOptionalBoolean(ownValue(value, 'showLastColumn')) &&
    isOptionalBoolean(ownValue(value, 'showRowStripes')) &&
    isOptionalBoolean(ownValue(value, 'showColumnStripes'))
  )
}

function isWorkbookPivotValue(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'sourceColumn') &&
    stringSetHas(PIVOT_AGGREGATION_VALUES, ownValue(value, 'summarizeBy')) &&
    isOptionalString(ownValue(value, 'outputLabel'))
  )
}

const CHART_TYPE_VALUES = new Set(['column', 'bar', 'line', 'area', 'pie', 'scatter'])
const CHART_SERIES_ORIENTATION_VALUES = new Set(['rows', 'columns'])
const CHART_LEGEND_POSITION_VALUES = new Set(['top', 'right', 'bottom', 'left', 'hidden'])
const DRAWING_ANCHOR_EDIT_AS_VALUES = new Set(['twoCell', 'oneCell', 'absolute'])
const SHAPE_TYPE_VALUES = new Set(['rectangle', 'roundedRectangle', 'ellipse', 'line', 'arrow', 'textBox'])

function isDrawingAnchorMarker(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasSafeNonNegativeInteger(value, 'row') &&
    hasSafeNonNegativeInteger(value, 'col') &&
    isOptionalSafeNonNegativeInteger(ownValue(value, 'rowOffset')) &&
    isOptionalSafeNonNegativeInteger(ownValue(value, 'colOffset'))
  )
}

function isDrawingAnchorExtent(value: unknown): boolean {
  return isRecord(value) && hasSafePositiveInteger(value, 'width') && hasSafePositiveInteger(value, 'height')
}

function isDrawingAnchorPosition(value: unknown): boolean {
  return isRecord(value) && hasSafeNonNegativeInteger(value, 'x') && hasSafeNonNegativeInteger(value, 'y')
}

function isWorkbookChartAnchor(value: unknown): boolean {
  if (!isRecord(value) || typeof ownValue(value, 'kind') !== 'string') {
    return false
  }
  switch (ownValue(value, 'kind')) {
    case 'twoCell':
      return (
        (ownValue(value, 'editAs') === undefined || stringSetHas(DRAWING_ANCHOR_EDIT_AS_VALUES, ownValue(value, 'editAs'))) &&
        isDrawingAnchorMarker(ownValue(value, 'from')) &&
        isDrawingAnchorMarker(ownValue(value, 'to'))
      )
    case 'oneCell':
      return isDrawingAnchorMarker(ownValue(value, 'from')) && isDrawingAnchorExtent(ownValue(value, 'extent'))
    case 'absolute':
      return isDrawingAnchorPosition(ownValue(value, 'position')) && isDrawingAnchorExtent(ownValue(value, 'extent'))
    default:
      return false
  }
}

function isWorkbookChart(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    hasString(value, 'sheetName') &&
    hasString(value, 'address') &&
    isCellRangeRef(ownValue(value, 'source')) &&
    stringSetHas(CHART_TYPE_VALUES, ownValue(value, 'chartType')) &&
    (ownValue(value, 'anchor') === undefined || isWorkbookChartAnchor(ownValue(value, 'anchor'))) &&
    (ownValue(value, 'seriesOrientation') === undefined ||
      stringSetHas(CHART_SERIES_ORIENTATION_VALUES, ownValue(value, 'seriesOrientation'))) &&
    isOptionalBoolean(ownValue(value, 'firstRowAsHeaders')) &&
    isOptionalBoolean(ownValue(value, 'firstColumnAsLabels')) &&
    isOptionalString(ownValue(value, 'title')) &&
    (ownValue(value, 'legendPosition') === undefined || stringSetHas(CHART_LEGEND_POSITION_VALUES, ownValue(value, 'legendPosition'))) &&
    hasSafePositiveInteger(value, 'rows') &&
    hasSafePositiveInteger(value, 'cols')
  )
}

function isWorkbookImage(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    hasString(value, 'sheetName') &&
    hasString(value, 'address') &&
    hasString(value, 'sourceUrl') &&
    hasSafePositiveInteger(value, 'rows') &&
    hasSafePositiveInteger(value, 'cols') &&
    isOptionalString(ownValue(value, 'altText'))
  )
}

function isWorkbookShape(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    hasString(value, 'sheetName') &&
    hasString(value, 'address') &&
    stringSetHas(SHAPE_TYPE_VALUES, ownValue(value, 'shapeType')) &&
    hasSafePositiveInteger(value, 'rows') &&
    hasSafePositiveInteger(value, 'cols') &&
    isOptionalString(ownValue(value, 'text')) &&
    isOptionalString(ownValue(value, 'fillColor')) &&
    isOptionalString(ownValue(value, 'strokeColor'))
  )
}

export function isWorkbookOp(value: unknown): value is WorkbookOp {
  if (!isRecord(value) || typeof ownValue(value, 'kind') !== 'string') {
    return false
  }

  switch (ownValue(value, 'kind')) {
    case 'upsertWorkbook':
      return hasString(value, 'name')
    case 'setWorkbookMetadata':
      return hasString(value, 'key') && isLiteralInput(ownValue(value, 'value'))
    case 'setCalculationSettings':
      return isWorkbookCalculationSettings(ownValue(value, 'settings'))
    case 'setVolatileContext':
      return isWorkbookVolatileContext(ownValue(value, 'context'))
    case 'upsertSheet':
      return hasString(value, 'name') && hasSafeNonNegativeInteger(value, 'order') && isOptionalSafePositiveInteger(ownValue(value, 'id'))
    case 'renameSheet':
      return hasString(value, 'oldName') && hasString(value, 'newName')
    case 'deleteSheet':
      return hasString(value, 'name')
    case 'insertRows':
    case 'insertColumns':
      return (
        hasString(value, 'sheetName') &&
        hasSafeNonNegativeInteger(value, 'start') &&
        hasSafePositiveInteger(value, 'count') &&
        (ownValue(value, 'entries') === undefined || arrayEvery(ownValue(value, 'entries'), isWorkbookAxisEntry))
      )
    case 'deleteRows':
    case 'deleteColumns':
      return hasString(value, 'sheetName') && hasSafeNonNegativeInteger(value, 'start') && hasSafePositiveInteger(value, 'count')
    case 'moveRows':
    case 'moveColumns':
      return (
        hasString(value, 'sheetName') &&
        hasSafeNonNegativeInteger(value, 'start') &&
        hasSafePositiveInteger(value, 'count') &&
        hasSafeNonNegativeInteger(value, 'target')
      )
    case 'updateRowMetadata':
    case 'updateColumnMetadata':
      return (
        hasString(value, 'sheetName') &&
        hasSafeNonNegativeInteger(value, 'start') &&
        hasSafePositiveInteger(value, 'count') &&
        isOptionalNullableSafePositiveInteger(ownValue(value, 'size')) &&
        isOptionalNullableBoolean(ownValue(value, 'hidden')) &&
        isOptionalNullableBoolean(ownValue(value, 'filterHidden'))
      )
    case 'setFreezePane':
      return hasString(value, 'sheetName') && hasSafeNonNegativeInteger(value, 'rows') && hasSafeNonNegativeInteger(value, 'cols')
    case 'clearFreezePane':
      return hasString(value, 'sheetName')
    case 'mergeCells':
    case 'unmergeCells':
      return isCellRangeRef(ownValue(value, 'range'))
    case 'setSheetProtection':
      return isWorkbookSheetProtection(ownValue(value, 'protection'))
    case 'clearSheetProtection':
      return hasString(value, 'sheetName')
    case 'setFilter':
    case 'clearFilter':
    case 'clearSort':
      return hasString(value, 'sheetName') && isCellRangeRef(ownValue(value, 'range'))
    case 'setSort':
      return (
        hasString(value, 'sheetName') && isCellRangeRef(ownValue(value, 'range')) && arrayEvery(ownValue(value, 'keys'), isWorkbookSortKey)
      )
    case 'setDataValidation':
      return isWorkbookDataValidation(ownValue(value, 'validation'))
    case 'clearDataValidation':
      return hasString(value, 'sheetName') && isCellRangeRef(ownValue(value, 'range'))
    case 'upsertConditionalFormat':
      return isWorkbookConditionalFormat(ownValue(value, 'format'))
    case 'deleteConditionalFormat':
      return hasString(value, 'id') && hasString(value, 'sheetName')
    case 'setConditionalFormatArtifacts':
      return hasString(value, 'sheetName') && isConditionalFormatArtifacts(ownValue(value, 'artifacts'))
    case 'clearConditionalFormatArtifacts':
      return hasString(value, 'sheetName')
    case 'upsertRangeProtection':
      return isWorkbookRangeProtection(ownValue(value, 'protection'))
    case 'deleteRangeProtection':
      return hasString(value, 'id') && hasString(value, 'sheetName')
    case 'upsertCommentThread':
      return isWorkbookCommentThread(ownValue(value, 'thread'))
    case 'deleteCommentThread':
    case 'deleteNote':
    case 'deleteHyperlink':
      return hasString(value, 'sheetName') && hasString(value, 'address')
    case 'upsertNote':
      return isWorkbookNote(ownValue(value, 'note'))
    case 'upsertHyperlink':
      return isWorkbookHyperlink(ownValue(value, 'hyperlink'))
    case 'setCellValue':
      return (
        hasString(value, 'sheetName') &&
        hasString(value, 'address') &&
        isLiteralInput(ownValue(value, 'value')) &&
        (ownValue(value, 'authoredBlank') === undefined || typeof ownValue(value, 'authoredBlank') === 'boolean')
      )
    case 'setCellFormula':
      return hasString(value, 'sheetName') && hasString(value, 'address') && hasString(value, 'formula')
    case 'setCellFormat':
      return (
        hasString(value, 'sheetName') &&
        hasString(value, 'address') &&
        (ownValue(value, 'format') === null || typeof ownValue(value, 'format') === 'string')
      )
    case 'upsertCellStyle':
      return isCellStyleRecord(ownValue(value, 'style'))
    case 'upsertCellNumberFormat':
      return isCellNumberFormatRecord(ownValue(value, 'format'))
    case 'setStyleRange':
      return isCellRangeRef(ownValue(value, 'range')) && hasString(value, 'styleId')
    case 'setFormatRange':
      return isCellRangeRef(ownValue(value, 'range')) && hasString(value, 'formatId')
    case 'clearCell':
      return hasString(value, 'sheetName') && hasString(value, 'address')
    case 'upsertDefinedName':
      return hasString(value, 'name') && isWorkbookDefinedNameValue(ownValue(value, 'value'))
    case 'deleteDefinedName':
    case 'deleteTable':
      return hasString(value, 'name')
    case 'upsertTable':
      return isWorkbookTableOp(ownValue(value, 'table'))
    case 'upsertSpillRange':
      return (
        hasString(value, 'sheetName') &&
        hasString(value, 'address') &&
        hasSafePositiveInteger(value, 'rows') &&
        hasSafePositiveInteger(value, 'cols')
      )
    case 'deleteSpillRange':
    case 'deletePivotTable':
      return hasString(value, 'sheetName') && hasString(value, 'address')
    case 'upsertPivotTable':
      return (
        hasString(value, 'name') &&
        hasString(value, 'sheetName') &&
        hasString(value, 'address') &&
        isCellRangeRef(ownValue(value, 'source')) &&
        isStringArray(ownValue(value, 'groupBy')) &&
        arrayEvery(ownValue(value, 'values'), isWorkbookPivotValue) &&
        hasSafePositiveInteger(value, 'rows') &&
        hasSafePositiveInteger(value, 'cols')
      )
    case 'upsertChart':
      return isWorkbookChart(ownValue(value, 'chart'))
    case 'deleteChart':
      return hasString(value, 'id')
    case 'upsertImage':
      return isWorkbookImage(ownValue(value, 'image'))
    case 'deleteImage':
      return hasString(value, 'id')
    case 'upsertShape':
      return isWorkbookShape(ownValue(value, 'shape'))
    case 'deleteShape':
      return hasString(value, 'id')
    default:
      return false
  }
}

export function isEngineOp(value: unknown): value is EngineOp {
  return isWorkbookOp(value)
}

export function isEngineOps(value: unknown): value is EngineOp[] {
  return arrayEvery(value, isEngineOp)
}

export function isEngineOpBatch(value: unknown): value is EngineOpBatch {
  if (!isRecord(value)) {
    return false
  }

  const clock = ownValue(value, 'clock')

  return (
    hasString(value, 'id') &&
    hasString(value, 'replicaId') &&
    isRecord(clock) &&
    hasSafeNonNegativeInteger(clock, 'counter') &&
    isEngineOps(ownValue(value, 'ops'))
  )
}
