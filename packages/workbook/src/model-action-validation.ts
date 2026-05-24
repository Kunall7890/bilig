import {
  CELL_BORDER_STYLE_VALUES,
  CELL_BORDER_WEIGHT_VALUES,
  CELL_HORIZONTAL_ALIGNMENT_VALUES,
  CELL_VERTICAL_ALIGNMENT_VALUES,
  isLiteralInput,
  type CellStylePatch,
  type LiteralInput,
} from '@bilig/protocol'
import { optionalDataProperty, isObject, isObjectRecord } from './data-properties.js'
import { isWorkbookRef, type WorkbookRef } from './find.js'
import { isWorkbookOp } from './guards.js'
import type { WorkbookOp } from './ops.js'

export interface NormalizedWorkbookAddOpOptions {
  readonly target?: WorkbookRef
  readonly message?: string
}

export function normalizeWorkbookActionTarget(action: string, target: unknown): WorkbookRef {
  if (!isWorkbookRef(target)) {
    throw new Error(`Workbook action ${action} target must be a workbook ref`)
  }
  return target
}

export function normalizeWorkbookActionLiteralInput(action: string, value: unknown): LiteralInput {
  if (!isLiteralInput(value)) {
    throw new Error(`Workbook action ${action} value must be a finite JSON literal`)
  }
  return value
}

export function normalizeWorkbookActionFormatOptions(options: unknown): {
  readonly style?: CellStylePatch
  readonly numberFormat?: string | null
} {
  if (!isObjectRecord(options)) {
    throw new Error('Workbook action format options must be an object')
  }
  assertOnlyDataProperties(options, 'Workbook action format options')

  const style = optionalDataProperty(options, 'style', 'Workbook action format style')
  const numberFormat = optionalDataProperty(options, 'numberFormat', 'Workbook action format numberFormat')
  const normalized: {
    style?: CellStylePatch
    numberFormat?: string | null
  } = {}

  if (style.status === 'present' && style.value !== undefined) {
    if (!isObject(style.value) || Array.isArray(style.value)) {
      throw new Error('Workbook action format style must be an object')
    }
    const clonedStyle = freezeData(cloneData(style.value))
    if (!isCellStylePatchValue(clonedStyle)) {
      throw new Error('Workbook action format style must be a valid cell style patch')
    }
    normalized.style = clonedStyle
  }

  if (numberFormat.status === 'present' && numberFormat.value !== undefined) {
    if (numberFormat.value !== null && typeof numberFormat.value !== 'string') {
      throw new Error('Workbook action format numberFormat must be a string, null, or undefined')
    }
    normalized.numberFormat = numberFormat.value
  }

  return Object.freeze(normalized)
}

export function normalizeWorkbookAddOpOptions(options: unknown): NormalizedWorkbookAddOpOptions {
  if (!isObjectRecord(options)) {
    throw new Error('Workbook action addOp options must be an object')
  }
  assertOnlyDataProperties(options, 'Workbook action addOp options')

  const target = optionalDataProperty(options, 'target', 'Workbook action addOp target')
  const message = optionalDataProperty(options, 'message', 'Workbook action addOp message')
  const normalized: {
    target?: WorkbookRef
    message?: string
  } = {}

  if (target.status === 'present' && target.value !== undefined) {
    normalized.target = normalizeWorkbookActionTarget('addOp', target.value)
  }
  if (message.status === 'present' && message.value !== undefined) {
    if (typeof message.value !== 'string') {
      throw new Error('Workbook action addOp message must be a string')
    }
    normalized.message = message.value
  }

  return Object.freeze(normalized)
}

export function normalizeWorkbookActionOp(op: unknown): WorkbookOp {
  assertOnlyDataProperties(op, 'Workbook action op')
  if (!isWorkbookOp(op)) {
    throw new Error('Workbook op is not a valid WorkbookOp')
  }
  const clonedOp = freezeData(cloneData(op))
  if (!isWorkbookOp(clonedOp)) {
    throw new Error('Workbook op is not a valid WorkbookOp')
  }
  return clonedOp
}

function isCellStylePatchValue(value: unknown): value is CellStylePatch {
  if (!isPlainRecord(value)) {
    return false
  }

  const fill = ownDataValue(value, 'fill')
  if (fill !== undefined && fill !== null && (!isPlainRecord(fill) || !isOptionalNullableString(fill, 'backgroundColor'))) {
    return false
  }

  const font = ownDataValue(value, 'font')
  if (
    font !== undefined &&
    font !== null &&
    (!isPlainRecord(font) ||
      !isOptionalNullableString(font, 'family') ||
      !isOptionalNullableNumber(font, 'size') ||
      !isOptionalNullableBoolean(font, 'bold') ||
      !isOptionalNullableBoolean(font, 'italic') ||
      !isOptionalNullableBoolean(font, 'underline') ||
      !isOptionalNullableString(font, 'color'))
  ) {
    return false
  }

  const alignment = ownDataValue(value, 'alignment')
  if (
    alignment !== undefined &&
    alignment !== null &&
    (!isPlainRecord(alignment) ||
      !isOptionalNullableStringValue(CELL_HORIZONTAL_ALIGNMENT_VALUES, ownDataValue(alignment, 'horizontal')) ||
      !isOptionalNullableStringValue(CELL_VERTICAL_ALIGNMENT_VALUES, ownDataValue(alignment, 'vertical')) ||
      !isOptionalNullableBoolean(alignment, 'wrap') ||
      !isOptionalNullableNumber(alignment, 'indent') ||
      !isOptionalNullableBoolean(alignment, 'shrinkToFit') ||
      !isOptionalNullableNumber(alignment, 'readingOrder') ||
      !isOptionalNullableNumber(alignment, 'textRotation') ||
      !isOptionalNullableBoolean(alignment, 'justifyLastLine'))
  ) {
    return false
  }

  const borders = ownDataValue(value, 'borders')
  if (borders !== undefined && borders !== null) {
    if (!isPlainRecord(borders)) {
      return false
    }
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const border = ownDataValue(borders, side)
      if (border !== undefined && border !== null && !isCellBorderSidePatchValue(border)) {
        return false
      }
    }
  }

  return true
}

function isCellBorderSidePatchValue(value: unknown): boolean {
  return (
    isPlainRecord(value) &&
    isOptionalNullableStringValue(CELL_BORDER_STYLE_VALUES, ownDataValue(value, 'style')) &&
    isOptionalNullableStringValue(CELL_BORDER_WEIGHT_VALUES, ownDataValue(value, 'weight')) &&
    isOptionalNullableString(value, 'color')
  )
}

function assertOnlyDataProperties(value: unknown, path: string): void {
  const invalidPath = firstNonDataPropertyPath(value, path)
  if (invalidPath !== null) {
    throw new Error(`${invalidPath} must be a data property`)
  }
}

function isPlainRecord(value: unknown): value is object {
  if (!isObject(value) || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function ownDataValue(value: object, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function isOptionalNullableString(value: object, key: string): boolean {
  const entry = ownDataValue(value, key)
  return entry === undefined || entry === null || typeof entry === 'string'
}

function isOptionalNullableNumber(value: object, key: string): boolean {
  const entry = ownDataValue(value, key)
  return entry === undefined || entry === null || (typeof entry === 'number' && Number.isFinite(entry))
}

function isOptionalNullableBoolean(value: object, key: string): boolean {
  const entry = ownDataValue(value, key)
  return entry === undefined || entry === null || typeof entry === 'boolean'
}

function isOptionalNullableStringValue(values: readonly string[], value: unknown): boolean {
  return value === undefined || value === null || (typeof value === 'string' && values.includes(value))
}

function firstNonDataPropertyPath(value: unknown, path: string, seen = new WeakSet<object>()): string | null {
  if (typeof value !== 'object' || value === null) {
    return null
  }
  if (seen.has(value)) {
    return null
  }
  seen.add(value)

  for (const key of Reflect.ownKeys(value)) {
    if (typeof key === 'symbol') {
      return `${path}[${String(key)}]`
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    const child = childPath(path, value, key)
    if (descriptor === undefined || !('value' in descriptor)) {
      return child
    }
    const nested = firstNonDataPropertyPath(descriptor.value, child, seen)
    if (nested !== null) {
      return nested
    }
  }
  return null
}

function childPath(parent: string, value: object, key: string): string {
  if (Array.isArray(value) && /^\d+$/.test(key)) {
    return `${parent}[${key}]`
  }
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(key) ? `${parent}.${key}` : `${parent}[${JSON.stringify(key)}]`
}

function freezeData<T>(value: T, seen = new WeakSet<object>()): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeData(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}

function cloneData(value: unknown, seen = new WeakMap<object, unknown>()): unknown {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  const existing = seen.get(value)
  if (existing !== undefined) {
    return existing
  }
  if (Array.isArray(value)) {
    const cloned: unknown[] = []
    seen.set(value, cloned)
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
      if (descriptor !== undefined && descriptor.enumerable && 'value' in descriptor) {
        cloned[index] = cloneData(descriptor.value, seen)
      }
    }
    return cloned
  }
  const cloned: Record<string, unknown> = {}
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (descriptor.enumerable && 'value' in descriptor) {
      Object.defineProperty(cloned, key, {
        configurable: true,
        enumerable: true,
        value: cloneData(descriptor.value, seen),
        writable: true,
      })
    }
  })
  return cloned
}
