import type { LiteralInput, WorkbookSnapshot } from '@bilig/protocol'
import { binarySearchUint32Prefix, encodeCellAddress } from './xlsx-large-simple-arena-helpers.js'
import {
  noPoolId,
  valueKindBoolean,
  valueKindEmpty,
  valueKindInteger,
  valueKindNull,
  valueKindNumber,
  valueKindSharedStringRef,
  valueKindSmallInteger,
  valueKindString,
  valueKindTinyInteger,
} from './xlsx-large-simple-arena-constants.js'
import { createLazyWorkbookSheetCells } from './xlsx-large-simple-lazy-sheet-cells.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'

type WorkbookSheetCells = WorkbookSnapshot['sheets'][number]['cells']
type WorkbookSheetCell = WorkbookSheetCells[number]

export interface DetachedLazySheetCellSource {
  readonly cellCount: number
  readonly arenaIndexes: Uint32Array | number
  readonly length: number
  readonly denseRowMajorWidth: number | null
  readonly linearCellIndexes?: Uint32Array
  readonly linearRowMajorWidth: number | null
  readonly rows: Uint32Array
  readonly columns: Uint16Array
  readonly valueKinds: Uint8Array
  readonly numberValues?: Float64Array
  readonly tinyIntegerValues?: Int8Array
  readonly smallIntegerValues?: Int16Array
  readonly sparseSmallIntegerCellIndexes?: Uint32Array
  readonly sparseSmallIntegerValues?: Int16Array
  readonly integerValues?: Int32Array
  readonly sparseIntegerCellIndexes?: Uint32Array
  readonly sparseIntegerValues?: Int32Array
  readonly stringIds?: Uint32Array
  readonly sparseStringCellIndexes?: Uint32Array
  readonly sparseStringIds?: Uint32Array
  readonly booleanValues?: Uint8Array
  readonly formulaIds?: Uint32Array
  readonly strings: readonly string[]
  readonly formulas: readonly string[]
  readonly sharedStrings?: LargeSimpleSharedStrings
  readonly sharedStringRefsInNumberValues: boolean
}

export function createDetachedLazyWorkbookSheetCells(source: DetachedLazySheetCellSource): WorkbookSheetCells {
  return createLazyWorkbookSheetCells(source.cellCount, (index) => materializeDetachedLazyCell(source, index))
}

function materializeDetachedLazyCell(source: DetachedLazySheetCellSource, index: number): WorkbookSheetCell | undefined {
  if (!Number.isInteger(index) || index < 0 || index >= source.cellCount) {
    return undefined
  }
  const arenaIndex = typeof source.arenaIndexes === 'number' ? index : (source.arenaIndexes[index] ?? -1)
  if (arenaIndex < 0 || arenaIndex >= source.length) {
    return undefined
  }
  const value = materializeDetachedValue(source, arenaIndex)
  const formulaId = source.formulaIds?.[arenaIndex] ?? noPoolId
  const formula = formulaId === noPoolId ? undefined : source.formulas[formulaId]
  if (value === undefined && formula === undefined) {
    return undefined
  }
  const row = detachedRowAt(source, arenaIndex)
  const col = detachedColumnAt(source, arenaIndex)
  return {
    address: encodeCellAddress(row, col),
    ...(value !== undefined ? { value } : {}),
    ...(formula !== undefined ? { formula } : {}),
  }
}

function materializeDetachedValue(source: DetachedLazySheetCellSource, index: number): LiteralInput | undefined {
  const valueKind = source.valueKinds[index] ?? valueKindEmpty
  switch (valueKind) {
    case valueKindNumber:
      return source.numberValues?.[index]
    case valueKindTinyInteger:
      return source.tinyIntegerValues?.[index]
    case valueKindSmallInteger:
      return smallIntegerValueAt(source, index)
    case valueKindInteger:
      return integerValueAt(source, index)
    case valueKindString: {
      const stringId = stringIdAt(source, index)
      return stringId === noPoolId ? undefined : source.strings[stringId]
    }
    case valueKindSharedStringRef: {
      const sharedStringIndex = sharedStringIndexAt(source, index)
      return sharedStringIndex === noPoolId ? undefined : source.sharedStrings?.[sharedStringIndex]?.text
    }
    case valueKindBoolean:
      return (source.booleanValues?.[index] ?? 0) === 1
    case valueKindNull:
      return null
    default:
      return undefined
  }
}

function detachedRowAt(source: DetachedLazySheetCellSource, index: number): number {
  if (source.denseRowMajorWidth !== null) {
    return Math.floor(index / source.denseRowMajorWidth)
  }
  if (source.linearCellIndexes && source.linearRowMajorWidth !== null) {
    return Math.floor((source.linearCellIndexes[index] ?? 0) / source.linearRowMajorWidth)
  }
  return source.rows[index] ?? 0
}

function detachedColumnAt(source: DetachedLazySheetCellSource, index: number): number {
  if (source.denseRowMajorWidth !== null) {
    return index % source.denseRowMajorWidth
  }
  if (source.linearCellIndexes && source.linearRowMajorWidth !== null) {
    return (source.linearCellIndexes[index] ?? 0) % source.linearRowMajorWidth
  }
  return source.columns[index] ?? 0
}

function smallIntegerValueAt(source: DetachedLazySheetCellSource, index: number): number | undefined {
  if (source.smallIntegerValues) {
    return source.smallIntegerValues[index]
  }
  const sparseIndexes = source.sparseSmallIntegerCellIndexes
  const sparseValues = source.sparseSmallIntegerValues
  if (!sparseIndexes || !sparseValues || sparseIndexes.length === 0) {
    return undefined
  }
  const offset = binarySearchUint32Prefix(sparseIndexes, sparseIndexes.length, index)
  return offset === -1 ? undefined : sparseValues[offset]
}

function integerValueAt(source: DetachedLazySheetCellSource, index: number): number | undefined {
  if (source.integerValues) {
    return source.integerValues[index]
  }
  const sparseIndexes = source.sparseIntegerCellIndexes
  const sparseValues = source.sparseIntegerValues
  if (!sparseIndexes || !sparseValues || sparseIndexes.length === 0) {
    return undefined
  }
  const offset = binarySearchUint32Prefix(sparseIndexes, sparseIndexes.length, index)
  return offset === -1 ? undefined : sparseValues[offset]
}

function sharedStringIndexAt(source: DetachedLazySheetCellSource, index: number): number {
  if (source.sharedStringRefsInNumberValues) {
    const value = source.numberValues?.[index]
    if (value !== undefined && !Number.isNaN(value)) {
      return Math.trunc(value)
    }
  }
  return stringIdAt(source, index)
}

function stringIdAt(source: DetachedLazySheetCellSource, index: number): number {
  if (source.stringIds) {
    return source.stringIds[index] ?? noPoolId
  }
  const sparseIndexes = source.sparseStringCellIndexes
  const sparseIds = source.sparseStringIds
  if (!sparseIndexes || !sparseIds) {
    return noPoolId
  }
  const offset = binarySearchUint32Prefix(sparseIndexes, sparseIndexes.length, index)
  return offset === -1 ? noPoolId : (sparseIds[offset] ?? noPoolId)
}
