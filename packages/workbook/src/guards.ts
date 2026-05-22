import { parseCellAddress, parseFormula } from '@bilig/formula'
import { isCellRangeRef as isProtocolCellRangeRef, isLiteralInput } from '@bilig/protocol'
import type { CellRangeRef } from '@bilig/protocol'
import type { EngineOp, EngineOpBatch, WorkbookOp } from './ops.js'

const HORIZONTAL_ALIGNMENT_VALUES = new Set(['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'])
const VERTICAL_ALIGNMENT_VALUES = new Set(['top', 'middle', 'bottom', 'justify', 'distributed'])
const BORDER_STYLE_VALUES = new Set(['solid', 'dashed', 'dotted', 'double'])
const BORDER_WEIGHT_VALUES = new Set(['thin', 'medium', 'thick'])
const NUMBER_FORMAT_KIND_VALUES = new Set(['general', 'number', 'currency', 'accounting', 'percent', 'date', 'time', 'datetime', 'text'])
const COMPATIBILITY_MODE_VALUES = new Set(['excel-modern', 'odf-1.4'])
const SORT_DIRECTION_VALUES = new Set(['asc', 'desc'])
const AUTO_FILTER_CUSTOM_OPERATOR_VALUES = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])
const PIVOT_AGGREGATION_VALUES = new Set(['sum', 'count', 'countNums', 'average', 'min', 'max', 'product'])
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
  return typeof value === 'object' && value !== null
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
  return Array.isArray(value) && value.every((entry) => typeof entry === 'string')
}

function isNonEmptyStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.length > 0 && value.every((entry) => typeof entry === 'string' && entry.trim() !== '')
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
  return typeof value[key] === 'string'
}

function hasNonEmptyString(value: Record<string, unknown>, key: string): boolean {
  return typeof value[key] === 'string' && value[key].trim() !== ''
}

function hasFiniteNumber(value: Record<string, unknown>, key: string): boolean {
  return isFiniteNumber(value[key])
}

function hasSafeNonNegativeInteger(value: Record<string, unknown>, key: string): boolean {
  return isSafeNonNegativeInteger(value[key])
}

function hasSafePositiveInteger(value: Record<string, unknown>, key: string): boolean {
  return isSafePositiveInteger(value[key])
}

function isCellAddressText(value: unknown): value is string {
  if (typeof value !== 'string' || value.trim() === '') {
    return false
  }
  try {
    return parseCellAddress(value).sheetName === undefined
  } catch {
    return false
  }
}

function hasCellAddress(value: Record<string, unknown>, key: string): boolean {
  return isCellAddressText(value[key])
}

function isFormulaText(value: unknown): value is string {
  if (typeof value !== 'string' || value.trim() === '') {
    return false
  }
  try {
    parseFormula(value)
    return true
  } catch {
    return false
  }
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  if (!isProtocolCellRangeRef(value) || value.sheetName.trim() === '') {
    return false
  }
  try {
    const start = parseCellAddress(value.startAddress)
    const end = parseCellAddress(value.endAddress)
    return start.sheetName === undefined && end.sheetName === undefined && end.row >= start.row && end.col >= start.col
  } catch {
    return false
  }
}

function isWorkbookAxisEntry(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'id') &&
    hasSafeNonNegativeInteger(value, 'index') &&
    isOptionalNullableSafePositiveInteger(value['size']) &&
    isOptionalNullableBoolean(value['hidden'])
  )
}

function isCellBorderSide(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'color') &&
    typeof value['style'] === 'string' &&
    BORDER_STYLE_VALUES.has(value['style']) &&
    typeof value['weight'] === 'string' &&
    BORDER_WEIGHT_VALUES.has(value['weight'])
  )
}

function isCellStyleRecord(value: unknown): boolean {
  if (!isRecord(value) || !hasString(value, 'id')) {
    return false
  }

  const fill = value['fill']
  if (fill !== undefined && (!isRecord(fill) || typeof fill['backgroundColor'] !== 'string')) {
    return false
  }

  const font = value['font']
  if (
    font !== undefined &&
    (!isRecord(font) ||
      !isOptionalString(font['family']) ||
      !isOptionalNumber(font['size']) ||
      !isOptionalBoolean(font['bold']) ||
      !isOptionalBoolean(font['italic']) ||
      !isOptionalBoolean(font['underline']) ||
      !isOptionalString(font['color']))
  ) {
    return false
  }

  const alignment = value['alignment']
  if (
    alignment !== undefined &&
    (!isRecord(alignment) ||
      !(
        alignment['horizontal'] === undefined ||
        (typeof alignment['horizontal'] === 'string' && HORIZONTAL_ALIGNMENT_VALUES.has(alignment['horizontal']))
      ) ||
      !(
        alignment['vertical'] === undefined ||
        (typeof alignment['vertical'] === 'string' && VERTICAL_ALIGNMENT_VALUES.has(alignment['vertical']))
      ) ||
      !isOptionalBoolean(alignment['wrap']) ||
      !isOptionalNumber(alignment['indent']) ||
      !isOptionalBoolean(alignment['shrinkToFit']) ||
      !isOptionalNumber(alignment['readingOrder']) ||
      !isOptionalNumber(alignment['textRotation']) ||
      !isOptionalBoolean(alignment['justifyLastLine']))
  ) {
    return false
  }

  const borders = value['borders']
  if (
    borders !== undefined &&
    (!isRecord(borders) ||
      !(borders['top'] === undefined || isCellBorderSide(borders['top'])) ||
      !(borders['right'] === undefined || isCellBorderSide(borders['right'])) ||
      !(borders['bottom'] === undefined || isCellBorderSide(borders['bottom'])) ||
      !(borders['left'] === undefined || isCellBorderSide(borders['left'])))
  ) {
    return false
  }

  const protection = value['protection']
  if (
    protection !== undefined &&
    (!isRecord(protection) || !isOptionalBoolean(protection['locked']) || !isOptionalBoolean(protection['hidden']))
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
    typeof value['kind'] === 'string' &&
    NUMBER_FORMAT_KIND_VALUES.has(value['kind'])
  )
}

function isWorkbookCalculationSettings(value: unknown): boolean {
  return (
    isRecord(value) &&
    (value['mode'] === 'automatic' || value['mode'] === 'manual') &&
    (value['compatibilityMode'] === undefined ||
      (typeof value['compatibilityMode'] === 'string' && COMPATIBILITY_MODE_VALUES.has(value['compatibilityMode'])))
  )
}

function isWorkbookVolatileContext(value: unknown): boolean {
  return isRecord(value) && hasFiniteNumber(value, 'recalcEpoch')
}

function isWorkbookSortKey(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasCellAddress(value, 'keyAddress') &&
    typeof value['direction'] === 'string' &&
    SORT_DIRECTION_VALUES.has(value['direction'])
  )
}

function isWorkbookValidationListSource(value: unknown): boolean {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'named-range':
      return hasNonEmptyString(value, 'name')
    case 'cell-ref':
      return hasNonEmptyString(value, 'sheetName') && hasCellAddress(value, 'address')
    case 'range-ref':
      return isCellRangeRef(value)
    case 'structured-ref':
      return hasNonEmptyString(value, 'tableName') && hasNonEmptyString(value, 'columnName')
    default:
      return false
  }
}

function isWorkbookDataValidationRule(value: unknown): boolean {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'list': {
      const hasValues = Array.isArray(value['values']) && value['values'].every((entry) => isLiteralInput(entry))
      const hasSource = value['source'] !== undefined && isWorkbookValidationListSource(value['source'])
      return (hasValues ? 1 : 0) + (hasSource ? 1 : 0) === 1
    }
    case 'checkbox':
      return isOptionalLiteralInput(value['checkedValue']) && isOptionalLiteralInput(value['uncheckedValue'])
    case 'any':
      return true
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time':
    case 'textLength':
      return (
        typeof value['operator'] === 'string' &&
        VALIDATION_COMPARISON_OPERATOR_VALUES.has(value['operator']) &&
        Array.isArray(value['values']) &&
        value['values'].every((entry) => isLiteralInput(entry)) &&
        (value['operator'] === 'between' || value['operator'] === 'notBetween'
          ? value['values'].length === 2
          : value['values'].length === 1)
      )
    default:
      return false
  }
}

function isWorkbookDataValidation(value: unknown): boolean {
  return (
    isRecord(value) &&
    isCellRangeRef(value['range']) &&
    isWorkbookDataValidationRule(value['rule']) &&
    isOptionalBoolean(value['allowBlank']) &&
    isOptionalBoolean(value['showDropdown']) &&
    isOptionalString(value['promptTitle']) &&
    isOptionalString(value['promptMessage']) &&
    (value['errorStyle'] === undefined ||
      (typeof value['errorStyle'] === 'string' && VALIDATION_ERROR_STYLE_VALUES.has(value['errorStyle']))) &&
    isOptionalString(value['errorTitle']) &&
    isOptionalString(value['errorMessage'])
  )
}

function isCellStylePatch(value: unknown): boolean {
  if (!isRecord(value)) {
    return false
  }

  const fill = value['fill']
  if (
    fill !== undefined &&
    fill !== null &&
    (!isRecord(fill) || !(fill['backgroundColor'] === undefined || fill['backgroundColor'] === null || hasString(fill, 'backgroundColor')))
  ) {
    return false
  }

  const font = value['font']
  if (
    font !== undefined &&
    font !== null &&
    (!isRecord(font) ||
      !(font['family'] === undefined || font['family'] === null || hasString(font, 'family')) ||
      !isOptionalNullableNumber(font['size']) ||
      !isOptionalNullableBoolean(font['bold']) ||
      !isOptionalNullableBoolean(font['italic']) ||
      !isOptionalNullableBoolean(font['underline']) ||
      !(font['color'] === undefined || font['color'] === null || hasString(font, 'color')))
  ) {
    return false
  }

  const alignment = value['alignment']
  if (
    alignment !== undefined &&
    alignment !== null &&
    (!isRecord(alignment) ||
      !(
        alignment['horizontal'] === undefined ||
        alignment['horizontal'] === null ||
        (typeof alignment['horizontal'] === 'string' && HORIZONTAL_ALIGNMENT_VALUES.has(alignment['horizontal']))
      ) ||
      !(
        alignment['vertical'] === undefined ||
        alignment['vertical'] === null ||
        (typeof alignment['vertical'] === 'string' && VERTICAL_ALIGNMENT_VALUES.has(alignment['vertical']))
      ) ||
      !isOptionalNullableBoolean(alignment['wrap']) ||
      !isOptionalNullableNumber(alignment['indent']) ||
      !isOptionalNullableBoolean(alignment['shrinkToFit']) ||
      !isOptionalNullableNumber(alignment['readingOrder']) ||
      !isOptionalNullableNumber(alignment['textRotation']) ||
      !isOptionalNullableBoolean(alignment['justifyLastLine']))
  ) {
    return false
  }

  const borders = value['borders']
  if (borders !== undefined && borders !== null) {
    if (!isRecord(borders)) {
      return false
    }
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const sideValue = borders[side]
      if (sideValue === undefined || sideValue === null) {
        continue
      }
      if (
        !isRecord(sideValue) ||
        !(
          sideValue['style'] === undefined ||
          sideValue['style'] === null ||
          (typeof sideValue['style'] === 'string' && BORDER_STYLE_VALUES.has(sideValue['style']))
        ) ||
        !(
          sideValue['weight'] === undefined ||
          sideValue['weight'] === null ||
          (typeof sideValue['weight'] === 'string' && BORDER_WEIGHT_VALUES.has(sideValue['weight']))
        ) ||
        !(sideValue['color'] === undefined || sideValue['color'] === null || hasString(sideValue, 'color'))
      ) {
        return false
      }
    }
  }

  return true
}

function isWorkbookConditionalFormatRule(value: unknown): boolean {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'cellIs':
      return (
        typeof value['operator'] === 'string' &&
        VALIDATION_COMPARISON_OPERATOR_VALUES.has(value['operator']) &&
        Array.isArray(value['values']) &&
        value['values'].every((entry) => isLiteralInput(entry)) &&
        (value['operator'] === 'between' || value['operator'] === 'notBetween'
          ? value['values'].length === 2
          : value['values'].length === 1)
      )
    case 'textContains':
      return hasString(value, 'text') && isOptionalBoolean(value['caseSensitive'])
    case 'formula':
      return isFormulaText(value['formula'])
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
    hasNonEmptyString(value, 'id') &&
    isCellRangeRef(value['range']) &&
    isWorkbookConditionalFormatRule(value['rule']) &&
    isCellStylePatch(value['style']) &&
    isOptionalBoolean(value['stopIfTrue']) &&
    isOptionalSafeNonNegativeInteger(value['priority'])
  )
}

function isWorkbookSheetProtection(value: unknown): boolean {
  return isRecord(value) && hasNonEmptyString(value, 'sheetName') && isOptionalBoolean(value['hideFormulas'])
}

function isWorkbookRangeProtection(value: unknown): boolean {
  return isRecord(value) && hasNonEmptyString(value, 'id') && isCellRangeRef(value['range']) && isOptionalBoolean(value['hideFormulas'])
}

function isWorkbookCommentEntry(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasNonEmptyString(value, 'id') &&
    hasString(value, 'body') &&
    isOptionalString(value['authorUserId']) &&
    isOptionalString(value['authorDisplayName']) &&
    isOptionalSafeNonNegativeInteger(value['createdAtUnixMs'])
  )
}

function isWorkbookCommentThread(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasNonEmptyString(value, 'threadId') &&
    hasNonEmptyString(value, 'sheetName') &&
    hasCellAddress(value, 'address') &&
    Array.isArray(value['comments']) &&
    value['comments'].length > 0 &&
    value['comments'].every((entry) => isWorkbookCommentEntry(entry)) &&
    isOptionalBoolean(value['resolved']) &&
    isOptionalString(value['resolvedByUserId']) &&
    isOptionalSafeNonNegativeInteger(value['resolvedAtUnixMs'])
  )
}

function isWorkbookNote(value: unknown): boolean {
  return isRecord(value) && hasNonEmptyString(value, 'sheetName') && hasCellAddress(value, 'address') && hasString(value, 'text')
}

function isWorkbookDefinedNameValue(value: unknown): boolean {
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
      return hasNonEmptyString(value, 'sheetName') && hasCellAddress(value, 'address')
    case 'range-ref':
      return isCellRangeRef(value)
    case 'structured-ref':
      return hasNonEmptyString(value, 'tableName') && hasNonEmptyString(value, 'columnName')
    case 'formula':
      return isFormulaText(value['formula'])
    default:
      return false
  }
}

function isWorkbookTableOp(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasNonEmptyString(value, 'name') &&
    isCellRangeRef({
      sheetName: value['sheetName'],
      startAddress: value['startAddress'],
      endAddress: value['endAddress'],
    }) &&
    isNonEmptyStringArray(value['columnNames']) &&
    typeof value['headerRow'] === 'boolean' &&
    typeof value['totalsRow'] === 'boolean' &&
    (value['columns'] === undefined || isWorkbookTableColumns(value['columns'])) &&
    (value['style'] === undefined || isWorkbookTableStyle(value['style'])) &&
    (value['autoFilter'] === undefined || isWorkbookAutoFilter(value['autoFilter'])) &&
    isOptionalString(value['sortState'])
  )
}

function isWorkbookAutoFilter(value: unknown): boolean {
  if (!isCellRangeRef(value) || !isRecord(value)) {
    return false
  }
  return (
    value['criteria'] === undefined ||
    (Array.isArray(value['criteria']) && value['criteria'].every((entry) => isWorkbookAutoFilterColumn(entry)))
  )
}

function isWorkbookAutoFilterColumn(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasSafeNonNegativeInteger(value, 'colId') &&
    isOptionalBoolean(value['hiddenButton']) &&
    isOptionalBoolean(value['showButton']) &&
    (value['filters'] === undefined || isWorkbookAutoFilterValueCriteria(value['filters'])) &&
    (value['customFilters'] === undefined || isWorkbookAutoFilterCustomCriteria(value['customFilters']))
  )
}

function isWorkbookAutoFilterValueCriteria(value: unknown): boolean {
  return isRecord(value) && isOptionalBoolean(value['blank']) && isStringArray(value['values'])
}

function isWorkbookAutoFilterCustomCriteria(value: unknown): boolean {
  return (
    isRecord(value) &&
    isOptionalBoolean(value['and']) &&
    Array.isArray(value['filters']) &&
    value['filters'].every((entry) => isWorkbookAutoFilterCustomCriterion(entry))
  )
}

function isWorkbookAutoFilterCustomCriterion(value: unknown): boolean {
  return (
    isRecord(value) &&
    (value['operator'] === undefined ||
      (typeof value['operator'] === 'string' && AUTO_FILTER_CUSTOM_OPERATOR_VALUES.has(value['operator']))) &&
    typeof value['value'] === 'string'
  )
}

function isWorkbookTableColumns(value: unknown): boolean {
  return (
    Array.isArray(value) &&
    value.every(
      (entry) =>
        isRecord(entry) &&
        hasString(entry, 'name') &&
        isOptionalString(entry['totalsRowLabel']) &&
        isOptionalString(entry['totalsRowFunction']),
    )
  )
}

function isWorkbookTableStyle(value: unknown): boolean {
  return (
    isRecord(value) &&
    isOptionalString(value['name']) &&
    isOptionalBoolean(value['showFirstColumn']) &&
    isOptionalBoolean(value['showLastColumn']) &&
    isOptionalBoolean(value['showRowStripes']) &&
    isOptionalBoolean(value['showColumnStripes'])
  )
}

function isWorkbookPivotValue(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'sourceColumn') &&
    typeof value['summarizeBy'] === 'string' &&
    PIVOT_AGGREGATION_VALUES.has(value['summarizeBy']) &&
    isOptionalString(value['outputLabel'])
  )
}

function isOptionalLiteralInputArray(value: unknown): boolean {
  return value === undefined || (Array.isArray(value) && value.every((entry) => isLiteralInput(entry)))
}

function isWorkbookPivotFilter(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'sourceColumn') &&
    isOptionalLiteralInputArray(value['includedValues']) &&
    isOptionalLiteralInputArray(value['hiddenValues'])
  )
}

function isWorkbookPivotPageField(value: unknown): boolean {
  return isRecord(value) && hasString(value, 'sourceColumn') && isOptionalLiteralInput(value['selectedValue'])
}

function isWorkbookPivotHiddenItems(value: unknown): boolean {
  return isRecord(value) && hasString(value, 'sourceColumn') && Array.isArray(value['values']) && value['values'].every(isLiteralInput)
}

function isWorkbookPivotCalculatedFormula(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, 'name') &&
    hasString(value, 'formula') &&
    (value['clause'] === '18.10' || value['clause'] === '3.2.3.1')
  )
}

function isOptionalStringArray(value: unknown): boolean {
  return value === undefined || isStringArray(value)
}

function isOptionalPivotFilterArray(value: unknown): boolean {
  return value === undefined || (Array.isArray(value) && value.every(isWorkbookPivotFilter))
}

function isOptionalPivotPageFieldArray(value: unknown): boolean {
  return value === undefined || (Array.isArray(value) && value.every(isWorkbookPivotPageField))
}

function isOptionalPivotHiddenItemsArray(value: unknown): boolean {
  return value === undefined || (Array.isArray(value) && value.every(isWorkbookPivotHiddenItems))
}

function isOptionalPivotCalculatedFormulaArray(value: unknown): boolean {
  return value === undefined || (Array.isArray(value) && value.every(isWorkbookPivotCalculatedFormula))
}

function isOptionalCachedRecords(value: unknown): boolean {
  return value === undefined || (Array.isArray(value) && value.every((row) => Array.isArray(row) && row.every(isLiteralInput)))
}

const CHART_TYPE_VALUES = new Set(['column', 'bar', 'line', 'area', 'pie', 'scatter'])
const CHART_SERIES_ORIENTATION_VALUES = new Set(['rows', 'columns'])
const CHART_LEGEND_POSITION_VALUES = new Set(['top', 'right', 'bottom', 'left', 'hidden'])
const SHAPE_TYPE_VALUES = new Set(['rectangle', 'roundedRectangle', 'ellipse', 'line', 'arrow', 'textBox'])

function isWorkbookChart(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasNonEmptyString(value, 'id') &&
    hasNonEmptyString(value, 'sheetName') &&
    hasCellAddress(value, 'address') &&
    isCellRangeRef(value['source']) &&
    typeof value['chartType'] === 'string' &&
    CHART_TYPE_VALUES.has(value['chartType']) &&
    (value['seriesOrientation'] === undefined ||
      (typeof value['seriesOrientation'] === 'string' && CHART_SERIES_ORIENTATION_VALUES.has(value['seriesOrientation']))) &&
    isOptionalBoolean(value['firstRowAsHeaders']) &&
    isOptionalBoolean(value['firstColumnAsLabels']) &&
    isOptionalString(value['title']) &&
    (value['legendPosition'] === undefined ||
      (typeof value['legendPosition'] === 'string' && CHART_LEGEND_POSITION_VALUES.has(value['legendPosition']))) &&
    hasSafePositiveInteger(value, 'rows') &&
    hasSafePositiveInteger(value, 'cols')
  )
}

function isWorkbookImage(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasNonEmptyString(value, 'id') &&
    hasNonEmptyString(value, 'sheetName') &&
    hasCellAddress(value, 'address') &&
    hasNonEmptyString(value, 'sourceUrl') &&
    hasSafePositiveInteger(value, 'rows') &&
    hasSafePositiveInteger(value, 'cols') &&
    isOptionalString(value['altText'])
  )
}

function isWorkbookShape(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasNonEmptyString(value, 'id') &&
    hasNonEmptyString(value, 'sheetName') &&
    hasCellAddress(value, 'address') &&
    typeof value['shapeType'] === 'string' &&
    SHAPE_TYPE_VALUES.has(value['shapeType']) &&
    hasSafePositiveInteger(value, 'rows') &&
    hasSafePositiveInteger(value, 'cols') &&
    isOptionalString(value['text']) &&
    isOptionalString(value['fillColor']) &&
    isOptionalString(value['strokeColor'])
  )
}

export function isWorkbookOp(value: unknown): value is WorkbookOp {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }

  switch (value['kind']) {
    case 'upsertWorkbook':
      return hasNonEmptyString(value, 'name')
    case 'setWorkbookMetadata':
      return hasNonEmptyString(value, 'key') && isLiteralInput(value['value'])
    case 'setCalculationSettings':
      return isWorkbookCalculationSettings(value['settings'])
    case 'setVolatileContext':
      return isWorkbookVolatileContext(value['context'])
    case 'upsertSheet':
      return hasNonEmptyString(value, 'name') && hasSafeNonNegativeInteger(value, 'order') && isOptionalSafePositiveInteger(value['id'])
    case 'renameSheet':
      return hasNonEmptyString(value, 'oldName') && hasNonEmptyString(value, 'newName')
    case 'deleteSheet':
      return hasNonEmptyString(value, 'name')
    case 'insertRows':
    case 'insertColumns':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasSafeNonNegativeInteger(value, 'start') &&
        hasSafePositiveInteger(value, 'count') &&
        (value['entries'] === undefined ||
          (Array.isArray(value['entries']) && value['entries'].every((entry) => isWorkbookAxisEntry(entry))))
      )
    case 'deleteRows':
    case 'deleteColumns':
      return hasNonEmptyString(value, 'sheetName') && hasSafeNonNegativeInteger(value, 'start') && hasSafePositiveInteger(value, 'count')
    case 'moveRows':
    case 'moveColumns':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasSafeNonNegativeInteger(value, 'start') &&
        hasSafePositiveInteger(value, 'count') &&
        hasSafeNonNegativeInteger(value, 'target')
      )
    case 'updateRowMetadata':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasSafeNonNegativeInteger(value, 'start') &&
        hasSafePositiveInteger(value, 'count') &&
        isOptionalNullableSafePositiveInteger(value['size']) &&
        isOptionalNullableBoolean(value['hidden']) &&
        isOptionalNullableBoolean(value['filtered'])
      )
    case 'updateColumnMetadata':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasSafeNonNegativeInteger(value, 'start') &&
        hasSafePositiveInteger(value, 'count') &&
        isOptionalNullableSafePositiveInteger(value['size']) &&
        isOptionalNullableBoolean(value['hidden'])
      )
    case 'setFreezePane':
      return hasNonEmptyString(value, 'sheetName') && hasSafeNonNegativeInteger(value, 'rows') && hasSafeNonNegativeInteger(value, 'cols')
    case 'clearFreezePane':
      return hasNonEmptyString(value, 'sheetName')
    case 'mergeCells':
    case 'unmergeCells':
      return isCellRangeRef(value['range'])
    case 'setSheetProtection':
      return isWorkbookSheetProtection(value['protection'])
    case 'clearSheetProtection':
      return hasNonEmptyString(value, 'sheetName')
    case 'setFilter':
    case 'clearFilter':
    case 'clearSort':
      return hasNonEmptyString(value, 'sheetName') && isCellRangeRef(value['range'])
    case 'setSort':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        isCellRangeRef(value['range']) &&
        Array.isArray(value['keys']) &&
        value['keys'].every((entry) => isWorkbookSortKey(entry))
      )
    case 'setDataValidation':
      return isWorkbookDataValidation(value['validation'])
    case 'clearDataValidation':
      return hasNonEmptyString(value, 'sheetName') && isCellRangeRef(value['range'])
    case 'upsertConditionalFormat':
      return isWorkbookConditionalFormat(value['format'])
    case 'deleteConditionalFormat':
      return hasNonEmptyString(value, 'id') && hasNonEmptyString(value, 'sheetName')
    case 'upsertRangeProtection':
      return isWorkbookRangeProtection(value['protection'])
    case 'deleteRangeProtection':
      return hasNonEmptyString(value, 'id') && hasNonEmptyString(value, 'sheetName')
    case 'upsertCommentThread':
      return isWorkbookCommentThread(value['thread'])
    case 'deleteCommentThread':
    case 'deleteNote':
      return hasNonEmptyString(value, 'sheetName') && hasCellAddress(value, 'address')
    case 'upsertNote':
      return isWorkbookNote(value['note'])
    case 'setCellValue':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasCellAddress(value, 'address') &&
        isLiteralInput(value['value']) &&
        (value['authoredBlank'] === undefined || typeof value['authoredBlank'] === 'boolean') &&
        (value['skipTableHeaderRename'] === undefined || typeof value['skipTableHeaderRename'] === 'boolean')
      )
    case 'setCellFormula':
      return hasNonEmptyString(value, 'sheetName') && hasCellAddress(value, 'address') && isFormulaText(value['formula'])
    case 'setCellFormat':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasCellAddress(value, 'address') &&
        (value['format'] === null || typeof value['format'] === 'string')
      )
    case 'upsertCellStyle':
      return isCellStyleRecord(value['style'])
    case 'upsertCellNumberFormat':
      return isCellNumberFormatRecord(value['format'])
    case 'setStyleRange':
      return isCellRangeRef(value['range']) && hasNonEmptyString(value, 'styleId')
    case 'setFormatRange':
      return isCellRangeRef(value['range']) && hasNonEmptyString(value, 'formatId')
    case 'clearCell':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasCellAddress(value, 'address') &&
        (value['skipTableHeaderRename'] === undefined || typeof value['skipTableHeaderRename'] === 'boolean')
      )
    case 'upsertDefinedName':
      return hasNonEmptyString(value, 'name') && isWorkbookDefinedNameValue(value['value'])
    case 'deleteDefinedName':
    case 'deleteTable':
      return hasNonEmptyString(value, 'name')
    case 'upsertTable':
      return isWorkbookTableOp(value['table'])
    case 'upsertSpillRange':
      return (
        hasNonEmptyString(value, 'sheetName') &&
        hasCellAddress(value, 'address') &&
        hasSafePositiveInteger(value, 'rows') &&
        hasSafePositiveInteger(value, 'cols')
      )
    case 'deleteSpillRange':
    case 'deletePivotTable':
      return hasNonEmptyString(value, 'sheetName') && hasCellAddress(value, 'address')
    case 'upsertPivotTable':
      return (
        hasNonEmptyString(value, 'name') &&
        hasNonEmptyString(value, 'sheetName') &&
        hasCellAddress(value, 'address') &&
        isCellRangeRef(value['source']) &&
        isStringArray(value['groupBy']) &&
        isOptionalStringArray(value['columnFields']) &&
        isOptionalPivotPageFieldArray(value['pageFields']) &&
        isOptionalPivotFilterArray(value['filters']) &&
        isOptionalPivotHiddenItemsArray(value['hiddenItems']) &&
        isOptionalPivotCalculatedFormulaArray(value['calculatedFields']) &&
        isOptionalPivotCalculatedFormulaArray(value['calculatedItems']) &&
        (value['sourceKind'] === undefined ||
          value['sourceKind'] === 'worksheet' ||
          value['sourceKind'] === 'table' ||
          value['sourceKind'] === 'named-range' ||
          value['sourceKind'] === 'external-cache-only') &&
        (value['cacheOnly'] === undefined || typeof value['cacheOnly'] === 'boolean') &&
        (value['cacheId'] === undefined || isSafeNonNegativeInteger(value['cacheId'])) &&
        isOptionalStringArray(value['cacheFields']) &&
        isOptionalCachedRecords(value['cachedRecords']) &&
        Array.isArray(value['values']) &&
        value['values'].every((entry) => isWorkbookPivotValue(entry)) &&
        hasSafePositiveInteger(value, 'rows') &&
        hasSafePositiveInteger(value, 'cols')
      )
    case 'upsertChart':
      return isWorkbookChart(value['chart'])
    case 'deleteChart':
      return hasNonEmptyString(value, 'id')
    case 'upsertImage':
      return isWorkbookImage(value['image'])
    case 'deleteImage':
      return hasNonEmptyString(value, 'id')
    case 'upsertShape':
      return isWorkbookShape(value['shape'])
    case 'deleteShape':
      return hasNonEmptyString(value, 'id')
    default:
      return false
  }
}

export function isEngineOp(value: unknown): value is EngineOp {
  return isWorkbookOp(value)
}

export function isEngineOps(value: unknown): value is EngineOp[] {
  return Array.isArray(value) && value.every((entry) => isEngineOp(entry))
}

export function isEngineOpBatch(value: unknown): value is EngineOpBatch {
  return (
    isRecord(value) &&
    hasNonEmptyString(value, 'id') &&
    hasNonEmptyString(value, 'replicaId') &&
    isRecord(value['clock']) &&
    hasSafeNonNegativeInteger(value['clock'], 'counter') &&
    isEngineOps(value['ops'])
  )
}
