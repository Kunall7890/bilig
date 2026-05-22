import type { LiteralInput } from '@bilig/protocol'
import {
  filledUint32Array,
  growFloat64Array,
  growInt16Array,
  growInt32Array,
  growInt8Array,
  growUint8Array,
  growUint16Array,
  growUint32Array,
} from './xlsx-large-simple-array-storage.js'
import { canStoreLinearCoordinate, previewCellCount, previewIndex } from './xlsx-large-simple-arena-helpers.js'
import { initialCellCapacity, noPoolId } from './xlsx-large-simple-arena-constants.js'
import type { ImportedWorkbookArenaDedupeMode, ImportedWorkbookArenaOptions } from './xlsx-large-simple-arena-types.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import { compactLinearRowMajorCoordinates } from './xlsx-large-simple-row-run-coordinate-compaction.js'

const denseRowMajorEarlyMismatchCapacityDivisor = 4
const narrowStringNoPoolId = 0xffff
export const maxNarrowStringId = narrowStringNoPoolId - 1

export type StringIdStorage = Uint16Array<ArrayBuffer> | Uint32Array<ArrayBuffer>

export function createStringIdStorage(length: number, wide = false): StringIdStorage {
  if (wide) {
    return filledUint32Array(length, noPoolId)
  }
  const output = new Uint16Array(length)
  output.fill(narrowStringNoPoolId)
  return output
}

export function growStringIdStorage(source: StringIdStorage, nextCapacity: number): StringIdStorage {
  if (source instanceof Uint32Array) {
    return growUint32Array(source, nextCapacity, noPoolId)
  }
  const output = growUint16Array(source, nextCapacity)
  if (nextCapacity > source.length) {
    output.fill(narrowStringNoPoolId, source.length)
  }
  return output
}

export function widenStringIdStorage(source: StringIdStorage): Uint32Array<ArrayBuffer> {
  if (source instanceof Uint32Array) {
    return source
  }
  const output = filledUint32Array(source.length, noPoolId)
  for (let index = 0; index < source.length; index += 1) {
    const value = source[index]
    if (value !== undefined && value !== narrowStringNoPoolId) {
      output[index] = value
    }
  }
  return output
}

export function canStoreStringIdInStorage(storage: StringIdStorage, stringId: number): boolean {
  return storage instanceof Uint32Array || stringId <= maxNarrowStringId
}

export function stringNoPoolIdForStorage(storage: StringIdStorage): number {
  return storage instanceof Uint16Array ? narrowStringNoPoolId : noPoolId
}

export function normalizeStoredStringId(value: number | undefined, storage: StringIdStorage): number {
  if (value === undefined) {
    return noPoolId
  }
  return storage instanceof Uint16Array && value === narrowStringNoPoolId ? noPoolId : value
}

function findRowRunIndex(rowRunIndexes: Uint32Array, rowRunCount: number, cellIndex: number): number {
  let low = 0
  let high = rowRunCount - 1
  let result = -1
  while (low <= high) {
    const mid = (low + high) >>> 1
    const start = rowRunIndexes[mid] ?? 0
    if (start <= cellIndex) {
      result = mid
      low = mid + 1
    } else {
      high = mid - 1
    }
  }
  return result
}

export abstract class ImportedWorkbookArenaCoordinateBase {
  protected sheetIndex: number | null = null
  protected sheetIndexes: Uint32Array<ArrayBuffer> | undefined
  protected rows: Uint32Array<ArrayBuffer> = new Uint32Array(initialCellCapacity)
  protected columns: Uint16Array<ArrayBuffer> = new Uint16Array(initialCellCapacity)
  protected rowRunIndexes: Uint32Array<ArrayBuffer> | undefined
  protected rowRunRows: Uint32Array<ArrayBuffer> | undefined
  protected rowRunStartColumns: Uint16Array<ArrayBuffer> | undefined
  protected columnPattern: Uint16Array<ArrayBuffer> | undefined
  protected rowRunCount = 0
  protected valueKinds: Uint8Array<ArrayBuffer> = new Uint8Array(initialCellCapacity)
  protected numberValues: Float64Array<ArrayBuffer> | undefined
  protected sparseNumberCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected sparseNumberValues: Float64Array<ArrayBuffer> | undefined
  protected sparseNumberCount = 0
  protected tinyIntegerValues: Int8Array<ArrayBuffer> | undefined
  protected smallIntegerValues: Int16Array<ArrayBuffer> | undefined
  protected sparseSmallIntegerCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected sparseSmallIntegerValues: Int16Array<ArrayBuffer> | undefined
  protected sparseSmallIntegerCount = 0
  protected integerValues: Int32Array<ArrayBuffer> | undefined
  protected sparseIntegerCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected sparseIntegerValues: Int32Array<ArrayBuffer> | undefined
  protected sparseIntegerCount = 0
  protected stringIds: StringIdStorage | undefined
  protected sparseStringCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected sparseStringIds: StringIdStorage | undefined
  protected sparseStringCount = 0
  protected booleanValues: Uint8Array<ArrayBuffer> | undefined
  protected formulaIds: Uint32Array<ArrayBuffer> | undefined
  protected length = 0
  protected denseRowMajorWidth: number | null = null
  protected linearCellIndexes: Uint32Array<ArrayBuffer> | undefined
  protected linearRowMajorWidth: number | null = null
  protected linearRowMajorCapacityLimit: number | null = null
  private streamRowRunCoordinates = false
  private streamRowRunLastRow = -1
  private streamRowRunLastColumn = -1
  protected strings: string[] = []
  protected readonly stringIdsByValue = new Map<string, number>()
  protected formulas: string[] = []
  protected readonly formulaIdsByValue = new Map<string, number>()
  protected sharedStrings: LargeSimpleSharedStrings | undefined
  protected stringValueCount = 0
  protected sharedStringRefCount = 0
  protected sharedStringRefsInNumberValues = false
  protected readonly previewValues: (LiteralInput | undefined)[] = Array.from({ length: previewCellCount })
  protected readonly previewValueSet = new Uint8Array(previewCellCount)
  protected readonly stringDedupeMode: ImportedWorkbookArenaDedupeMode
  protected readonly formulaDedupeMode: ImportedWorkbookArenaDedupeMode
  protected readonly dedupeMaxEntries: number
  protected readonly stringDedupeKeys: string[] = []
  protected stringDedupeEvictionIndex = 0
  protected readonly formulaDedupeKeys: string[] = []
  protected formulaDedupeEvictionIndex = 0

  constructor(
    protected readonly stringPool?: ImportedWorkbookStringPool,
    options: ImportedWorkbookArenaOptions = {},
  ) {
    this.stringDedupeMode = options.deduplicateStrings ?? true
    this.formulaDedupeMode = options.deduplicateFormulas ?? true
    this.dedupeMaxEntries = Math.max(0, Math.trunc(options.dedupeMaxEntries ?? 8192))
  }

  get cellCount(): number {
    return this.length
  }

  retainedStorageByteLength(): number {
    return (
      (this.sheetIndexes?.byteLength ?? 0) +
      this.rows.byteLength +
      this.columns.byteLength +
      (this.rowRunIndexes?.byteLength ?? 0) +
      (this.rowRunRows?.byteLength ?? 0) +
      (this.rowRunStartColumns?.byteLength ?? 0) +
      (this.columnPattern?.byteLength ?? 0) +
      this.valueKinds.byteLength +
      (this.numberValues?.byteLength ?? 0) +
      (this.sparseNumberCellIndexes?.byteLength ?? 0) +
      (this.sparseNumberValues?.byteLength ?? 0) +
      (this.tinyIntegerValues?.byteLength ?? 0) +
      (this.smallIntegerValues?.byteLength ?? 0) +
      (this.sparseSmallIntegerCellIndexes?.byteLength ?? 0) +
      (this.sparseSmallIntegerValues?.byteLength ?? 0) +
      (this.integerValues?.byteLength ?? 0) +
      (this.sparseIntegerCellIndexes?.byteLength ?? 0) +
      (this.sparseIntegerValues?.byteLength ?? 0) +
      (this.stringIds?.byteLength ?? 0) +
      (this.sparseStringCellIndexes?.byteLength ?? 0) +
      (this.sparseStringIds?.byteLength ?? 0) +
      (this.booleanValues?.byteLength ?? 0) +
      (this.formulaIds?.byteLength ?? 0) +
      (this.linearCellIndexes?.byteLength ?? 0)
    )
  }

  reserveCellCapacity(capacity: number): void {
    if (!Number.isSafeInteger(capacity) || capacity <= this.valueKinds.length) {
      return
    }
    this.resizeStorage(capacity)
  }

  reserveDenseRowMajorCellCapacity(sheetIndex: number, width: number, height: number): void {
    const capacity = width * height
    if (!Number.isSafeInteger(capacity) || width <= 0 || height <= 0) {
      return
    }
    if (this.length !== 0 || this.sheetIndexes || (this.sheetIndex !== null && this.sheetIndex !== sheetIndex)) {
      this.reserveCellCapacity(capacity)
      return
    }
    this.sheetIndex = sheetIndex
    this.denseRowMajorWidth = width
    this.linearRowMajorCapacityLimit = null
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
    this.reserveCellCapacity(capacity)
  }

  trackLinearRowMajorCellCoordinates(sheetIndex: number, width: number, height: number): void {
    const capacity = width * height
    if (
      this.length !== 0 ||
      this.sheetIndexes ||
      (this.sheetIndex !== null && this.sheetIndex !== sheetIndex) ||
      !Number.isSafeInteger(capacity) ||
      capacity <= 0 ||
      !canStoreLinearCoordinate(width, height - 1, width - 1)
    ) {
      return
    }
    this.sheetIndex = sheetIndex
    this.linearRowMajorWidth = width
    this.linearRowMajorCapacityLimit = capacity
    this.linearCellIndexes = new Uint32Array(this.valueKinds.length)
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
  }

  trackRowRunCellCoordinates(sheetIndex: number): void {
    if (this.length !== 0 || this.sheetIndexes || (this.sheetIndex !== null && this.sheetIndex !== sheetIndex)) {
      return
    }
    this.sheetIndex = sheetIndex
    this.denseRowMajorWidth = null
    this.linearCellIndexes = undefined
    this.linearRowMajorWidth = null
    this.linearRowMajorCapacityLimit = null
    this.rowRunIndexes = new Uint32Array(initialCellCapacity)
    this.rowRunRows = new Uint32Array(initialCellCapacity)
    this.rowRunStartColumns = new Uint16Array(initialCellCapacity)
    this.columnPattern = undefined
    this.rowRunCount = 0
    this.streamRowRunCoordinates = true
    this.streamRowRunLastRow = -1
    this.streamRowRunLastColumn = -1
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
  }

  protected appendCell(sheetIndex: number, row: number, column: number): number {
    if (this.rowRunRows && !this.streamRowRunCoordinates) {
      this.materializeRowRunCoordinateStorage()
    }
    const index = this.length
    this.length += 1
    this.recordSheetIndex(index, sheetIndex)
    if (this.streamRowRunCoordinates) {
      if (this.appendStreamedRowRunCoordinate(index, row, column)) {
        return index
      }
      this.materializeRowRunCoordinateStorage(index)
    }
    if (this.denseRowMajorWidth !== null) {
      const expectedRow = Math.floor(index / this.denseRowMajorWidth)
      if (
        this.sheetIndexes === undefined &&
        this.sheetIndex === sheetIndex &&
        row === expectedRow &&
        column === index % this.denseRowMajorWidth
      ) {
        return index
      }
      if (this.shouldAbandonDenseRowMajorPreallocation(index)) {
        if (this.switchDenseRowMajorPreallocationToRowRunStorage(index)) {
          if (this.appendStreamedRowRunCoordinate(index, row, column)) {
            return index
          }
          this.materializeRowRunCoordinateStorage(index)
        } else {
          this.shrinkDenseRowMajorPreallocationToCoordinateStorage(index)
        }
      } else if (canStoreLinearCoordinate(this.denseRowMajorWidth, row, column)) {
        this.materializeLinearCoordinateStorage(index)
      } else {
        this.materializeCoordinateStorage(index)
      }
    }
    if (this.linearCellIndexes && this.linearRowMajorWidth !== null) {
      if (canStoreLinearCoordinate(this.linearRowMajorWidth, row, column)) {
        this.linearCellIndexes[index] = row * this.linearRowMajorWidth + column
        return index
      }
      this.materializeCoordinateStorage(index)
    }
    this.rows[index] = row
    this.columns[index] = column
    return index
  }

  private appendStreamedRowRunCoordinate(index: number, row: number, column: number): boolean {
    if (!this.rowRunIndexes || !this.rowRunRows) {
      return false
    }
    if (this.rowRunCount === 0 || row > this.streamRowRunLastRow) {
      this.ensureRowRunCapacity(this.rowRunCount + 1)
      this.rowRunIndexes[this.rowRunCount] = index
      this.rowRunRows[this.rowRunCount] = row
      if (this.rowRunStartColumns) {
        this.rowRunStartColumns[this.rowRunCount] = column
      }
      if (this.columns.length > 0) {
        this.columns[index] = column
      }
      this.rowRunCount += 1
      this.streamRowRunLastRow = row
      this.streamRowRunLastColumn = column
      return true
    }
    if (row === this.streamRowRunLastRow && column > this.streamRowRunLastColumn) {
      if (column !== this.streamRowRunLastColumn + 1) {
        this.materializeStreamRowRunColumnStorage(index)
      }
      if (this.columns.length > 0) {
        this.columns[index] = column
      }
      this.streamRowRunLastColumn = column
      return true
    }
    this.streamRowRunCoordinates = false
    this.streamRowRunLastRow = -1
    this.streamRowRunLastColumn = -1
    return false
  }

  private materializeStreamRowRunColumnStorage(upToIndex: number): void {
    if (this.columns.length > 0 || !this.rowRunIndexes || !this.rowRunStartColumns) {
      return
    }
    this.columns = new Uint16Array(this.valueKinds.length)
    for (let run = 0; run < this.rowRunCount; run += 1) {
      const start = this.rowRunIndexes[run] ?? 0
      const end = run + 1 < this.rowRunCount ? (this.rowRunIndexes[run + 1] ?? upToIndex) : upToIndex
      const startColumn = this.rowRunStartColumns[run] ?? 0
      for (let index = start; index < Math.min(end, this.columns.length, upToIndex); index += 1) {
        this.columns[index] = startColumn + index - start
      }
    }
    this.rowRunStartColumns = undefined
  }

  private ensureRowRunCapacity(nextCount: number): void {
    if (!this.rowRunIndexes || !this.rowRunRows || nextCount <= this.rowRunIndexes.length) {
      return
    }
    const nextCapacity = this.rowRunIndexes.length * 2
    this.rowRunIndexes = growUint32Array(this.rowRunIndexes, nextCapacity)
    this.rowRunRows = growUint32Array(this.rowRunRows, nextCapacity)
    if (this.rowRunStartColumns) {
      this.rowRunStartColumns = growUint16Array(this.rowRunStartColumns, nextCapacity)
    }
  }

  private shouldAbandonDenseRowMajorPreallocation(mismatchIndex: number): boolean {
    return mismatchIndex < initialCellCapacity || mismatchIndex * denseRowMajorEarlyMismatchCapacityDivisor < this.valueKinds.length
  }

  protected recordSheetIndex(index: number, sheetIndex: number): void {
    if (this.sheetIndexes) {
      this.sheetIndexes[index] = sheetIndex
      return
    }
    if (this.sheetIndex === null) {
      this.sheetIndex = sheetIndex
      return
    }
    if (this.sheetIndex === sheetIndex) {
      return
    }
    const sheetIndexes = new Uint32Array(this.valueKinds.length)
    sheetIndexes.fill(this.sheetIndex, 0, index)
    sheetIndexes[index] = sheetIndex
    this.sheetIndexes = sheetIndexes
  }

  protected rowAt(index: number): number {
    if (this.denseRowMajorWidth !== null) {
      return Math.floor(index / this.denseRowMajorWidth)
    }
    if (this.linearCellIndexes && this.linearRowMajorWidth !== null) {
      return Math.floor((this.linearCellIndexes[index] ?? 0) / this.linearRowMajorWidth)
    }
    if (this.rowRunIndexes && this.rowRunRows) {
      const runIndex = findRowRunIndex(this.rowRunIndexes, this.rowRunCount, index)
      return runIndex === -1 ? 0 : (this.rowRunRows[runIndex] ?? 0)
    }
    return this.rows[index] ?? 0
  }

  protected columnAt(index: number): number {
    if (this.denseRowMajorWidth !== null) {
      return index % this.denseRowMajorWidth
    }
    if (this.linearCellIndexes && this.linearRowMajorWidth !== null) {
      return (this.linearCellIndexes[index] ?? 0) % this.linearRowMajorWidth
    }
    if (this.rowRunIndexes && this.rowRunStartColumns) {
      const runIndex = findRowRunIndex(this.rowRunIndexes, this.rowRunCount, index)
      const start = runIndex === -1 ? 0 : (this.rowRunIndexes[runIndex] ?? 0)
      return (runIndex === -1 ? 0 : (this.rowRunStartColumns[runIndex] ?? 0)) + index - start
    }
    if (this.rowRunIndexes && this.columnPattern) {
      const runIndex = findRowRunIndex(this.rowRunIndexes, this.rowRunCount, index)
      const start = runIndex === -1 ? 0 : (this.rowRunIndexes[runIndex] ?? 0)
      return this.columnPattern[index - start] ?? 0
    }
    return this.columns[index] ?? 0
  }

  protected materializeCoordinateStorage(upToIndex: number): void {
    if (this.rowRunRows) {
      this.materializeRowRunCoordinateStorage()
    }
    const width = this.denseRowMajorWidth ?? this.linearRowMajorWidth
    if (width === null) {
      return
    }
    this.rows = new Uint32Array(this.valueKinds.length)
    this.columns = new Uint16Array(this.valueKinds.length)
    const linearCellIndexes = this.linearCellIndexes
    if (linearCellIndexes) {
      for (let index = 0; index < upToIndex; index += 1) {
        const linearCellIndex = linearCellIndexes[index] ?? 0
        this.rows[index] = Math.floor(linearCellIndex / width)
        this.columns[index] = linearCellIndex % width
      }
    } else {
      for (let index = 0; index < upToIndex; index += 1) {
        this.rows[index] = Math.floor(index / width)
        this.columns[index] = index % width
      }
    }
    this.denseRowMajorWidth = null
    this.linearCellIndexes = undefined
    this.linearRowMajorWidth = null
    this.linearRowMajorCapacityLimit = null
  }

  protected compactRowMajorRowCoordinates(): void {
    if (this.linearCellIndexes && this.linearRowMajorWidth !== null && !this.sheetIndexes && !this.rowRunRows) {
      const compacted = compactLinearRowMajorCoordinates(this.linearCellIndexes, this.length, this.linearRowMajorWidth)
      if (!compacted) {
        return
      }
      this.rowRunIndexes = compacted.rowRunIndexes
      this.rowRunRows = compacted.rowRunRows
      this.rowRunStartColumns = compacted.rowRunStartColumns
      this.columnPattern = compacted.columnPattern
      this.rowRunCount = compacted.runCount
      this.rows = new Uint32Array(0)
      this.columns = compacted.columns ?? new Uint16Array(0)
      this.linearCellIndexes = undefined
      this.linearRowMajorWidth = null
      this.linearRowMajorCapacityLimit = null
      return
    }
    if (
      this.length === 0 ||
      this.denseRowMajorWidth !== null ||
      this.linearCellIndexes ||
      this.rowRunRows ||
      this.sheetIndexes ||
      this.rows.length < this.length ||
      this.columns.length < this.length
    ) {
      return
    }
    let runCount = 0
    let previousRow = -1
    let previousColumn = -1
    let nextContiguousColumn = -1
    let columnsAreContiguousByRun = true
    for (let index = 0; index < this.length; index += 1) {
      const row = this.rows[index] ?? 0
      const column = this.columns[index] ?? 0
      if (row < previousRow || (row === previousRow && column <= previousColumn)) {
        return
      }
      if (row !== previousRow) {
        runCount += 1
        previousRow = row
        nextContiguousColumn = column + 1
      } else if (column === nextContiguousColumn) {
        nextContiguousColumn += 1
      } else {
        columnsAreContiguousByRun = false
      }
      previousColumn = column
    }
    if (runCount * 2 >= this.length) {
      return
    }
    const rowRunIndexes = new Uint32Array(runCount)
    const rowRunRows = new Uint32Array(runCount)
    const rowRunStartColumns = columnsAreContiguousByRun ? new Uint16Array(runCount) : undefined
    let outputIndex = 0
    previousRow = -1
    for (let index = 0; index < this.length; index += 1) {
      const row = this.rows[index] ?? 0
      if (row === previousRow) {
        continue
      }
      rowRunIndexes[outputIndex] = index
      rowRunRows[outputIndex] = row
      if (rowRunStartColumns) {
        rowRunStartColumns[outputIndex] = this.columns[index] ?? 0
      }
      outputIndex += 1
      previousRow = row
    }
    this.rowRunIndexes = rowRunIndexes
    this.rowRunRows = rowRunRows
    this.rowRunStartColumns = rowRunStartColumns
    this.rowRunCount = runCount
    this.rows = new Uint32Array(0)
    if (rowRunStartColumns) {
      this.columns = new Uint16Array(0)
    }
  }

  private materializeRowRunCoordinateStorage(upToIndex = this.length): void {
    const rowRunIndexes = this.rowRunIndexes
    const rowRunRows = this.rowRunRows
    if (!rowRunIndexes || !rowRunRows) {
      return
    }
    this.rows = new Uint32Array(this.valueKinds.length)
    for (let run = 0; run < this.rowRunCount; run += 1) {
      const start = rowRunIndexes[run] ?? 0
      const end = run + 1 < this.rowRunCount ? (rowRunIndexes[run + 1] ?? upToIndex) : upToIndex
      this.rows.fill(rowRunRows[run] ?? 0, start, Math.min(end, this.rows.length, upToIndex))
    }
    const columnPattern = this.columnPattern
    if (this.rowRunStartColumns || columnPattern) {
      this.columns = new Uint16Array(this.valueKinds.length)
      for (let run = 0; run < this.rowRunCount; run += 1) {
        const start = rowRunIndexes[run] ?? 0
        const end = run + 1 < this.rowRunCount ? (rowRunIndexes[run + 1] ?? upToIndex) : upToIndex
        if (this.rowRunStartColumns) {
          const startColumn = this.rowRunStartColumns[run] ?? 0
          for (let index = start; index < Math.min(end, this.columns.length, upToIndex); index += 1) {
            this.columns[index] = startColumn + index - start
          }
        } else {
          for (let index = start; index < Math.min(end, this.columns.length, upToIndex); index += 1) {
            this.columns[index] = columnPattern?.[index - start] ?? 0
          }
        }
      }
    }
    this.rowRunIndexes = undefined
    this.rowRunRows = undefined
    this.rowRunStartColumns = undefined
    this.columnPattern = undefined
    this.rowRunCount = 0
    this.streamRowRunCoordinates = false
    this.streamRowRunLastRow = -1
    this.streamRowRunLastColumn = -1
  }

  private switchDenseRowMajorPreallocationToRowRunStorage(upToIndex: number): boolean {
    const width = this.denseRowMajorWidth
    if (width === null || this.sheetIndexes || width <= 0 || upToIndex < 0) {
      return false
    }
    const nextCapacity = Math.max(initialCellCapacity, this.length)
    const copyLength = Math.min(upToIndex, nextCapacity)
    const previousValueKinds = this.valueKinds
    const previousNumberValues = this.numberValues
    const previousTinyIntegerValues = this.tinyIntegerValues
    const previousSmallIntegerValues = this.smallIntegerValues
    const previousIntegerValues = this.integerValues
    const previousStringIds = this.stringIds
    const previousBooleanValues = this.booleanValues
    const previousFormulaIds = this.formulaIds

    this.valueKinds = new Uint8Array(nextCapacity)
    this.valueKinds.set(previousValueKinds.subarray(0, copyLength))
    if (previousNumberValues) {
      this.numberValues = new Float64Array(nextCapacity)
      this.numberValues.fill(Number.NaN)
      this.numberValues.set(previousNumberValues.subarray(0, copyLength))
    }
    if (previousTinyIntegerValues) {
      this.tinyIntegerValues = new Int8Array(nextCapacity)
      this.tinyIntegerValues.set(previousTinyIntegerValues.subarray(0, copyLength))
    }
    if (previousSmallIntegerValues) {
      this.smallIntegerValues = new Int16Array(nextCapacity)
      this.smallIntegerValues.set(previousSmallIntegerValues.subarray(0, copyLength))
    }
    if (previousIntegerValues) {
      this.integerValues = new Int32Array(nextCapacity)
      this.integerValues.set(previousIntegerValues.subarray(0, copyLength))
    }
    if (previousStringIds) {
      this.stringIds = createStringIdStorage(nextCapacity, previousStringIds instanceof Uint32Array)
      this.stringIds.set(previousStringIds.subarray(0, copyLength))
    }
    if (previousBooleanValues) {
      this.booleanValues = new Uint8Array(nextCapacity)
      this.booleanValues.set(previousBooleanValues.subarray(0, copyLength))
    }
    if (previousFormulaIds) {
      this.formulaIds = filledUint32Array(nextCapacity, noPoolId)
      this.formulaIds.set(previousFormulaIds.subarray(0, copyLength))
    }
    const runCount = upToIndex === 0 ? 0 : Math.floor((upToIndex - 1) / width) + 1
    const runCapacity = Math.max(16, runCount)
    this.rowRunIndexes = new Uint32Array(runCapacity)
    this.rowRunRows = new Uint32Array(runCapacity)
    this.rowRunStartColumns = new Uint16Array(runCapacity)
    for (let run = 0; run < runCount; run += 1) {
      this.rowRunIndexes[run] = run * width
      this.rowRunRows[run] = run
      this.rowRunStartColumns[run] = 0
    }
    this.rowRunCount = runCount
    this.denseRowMajorWidth = null
    this.linearCellIndexes = undefined
    this.linearRowMajorWidth = null
    this.linearRowMajorCapacityLimit = null
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
    this.streamRowRunCoordinates = true
    this.streamRowRunLastRow = runCount === 0 ? -1 : runCount - 1
    this.streamRowRunLastColumn = upToIndex === 0 ? -1 : (upToIndex - 1) % width
    return true
  }

  private shrinkDenseRowMajorPreallocationToCoordinateStorage(upToIndex: number): void {
    const width = this.denseRowMajorWidth
    if (width === null) {
      return
    }
    const nextCapacity = Math.max(initialCellCapacity, this.length)
    const copyLength = Math.min(upToIndex, nextCapacity)
    const previousValueKinds = this.valueKinds
    const previousSheetIndexes = this.sheetIndexes
    const previousNumberValues = this.numberValues
    const previousTinyIntegerValues = this.tinyIntegerValues
    const previousSmallIntegerValues = this.smallIntegerValues
    const previousIntegerValues = this.integerValues
    const previousStringIds = this.stringIds
    const previousBooleanValues = this.booleanValues
    const previousFormulaIds = this.formulaIds

    this.rows = new Uint32Array(nextCapacity)
    this.columns = new Uint16Array(nextCapacity)
    for (let index = 0; index < copyLength; index += 1) {
      this.rows[index] = Math.floor(index / width)
      this.columns[index] = index % width
    }
    if (previousSheetIndexes) {
      this.sheetIndexes = new Uint32Array(nextCapacity)
      this.sheetIndexes.set(previousSheetIndexes.subarray(0, copyLength))
    }
    this.valueKinds = new Uint8Array(nextCapacity)
    this.valueKinds.set(previousValueKinds.subarray(0, copyLength))
    if (previousNumberValues) {
      this.numberValues = new Float64Array(nextCapacity)
      this.numberValues.fill(Number.NaN)
      this.numberValues.set(previousNumberValues.subarray(0, copyLength))
    }
    if (previousTinyIntegerValues) {
      this.tinyIntegerValues = new Int8Array(nextCapacity)
      this.tinyIntegerValues.set(previousTinyIntegerValues.subarray(0, copyLength))
    }
    if (previousSmallIntegerValues) {
      this.smallIntegerValues = new Int16Array(nextCapacity)
      this.smallIntegerValues.set(previousSmallIntegerValues.subarray(0, copyLength))
    }
    if (previousIntegerValues) {
      this.integerValues = new Int32Array(nextCapacity)
      this.integerValues.set(previousIntegerValues.subarray(0, copyLength))
    }
    if (previousStringIds) {
      this.stringIds = createStringIdStorage(nextCapacity, previousStringIds instanceof Uint32Array)
      this.stringIds.set(previousStringIds.subarray(0, copyLength))
    }
    if (previousBooleanValues) {
      this.booleanValues = new Uint8Array(nextCapacity)
      this.booleanValues.set(previousBooleanValues.subarray(0, copyLength))
    }
    if (previousFormulaIds) {
      this.formulaIds = filledUint32Array(nextCapacity, noPoolId)
      this.formulaIds.set(previousFormulaIds.subarray(0, copyLength))
    }
    this.denseRowMajorWidth = null
    this.linearCellIndexes = undefined
    this.linearRowMajorWidth = null
    this.linearRowMajorCapacityLimit = null
  }

  protected materializeLinearCoordinateStorage(upToIndex: number): void {
    const width = this.denseRowMajorWidth
    if (width === null) {
      return
    }
    this.linearCellIndexes = new Uint32Array(this.valueKinds.length)
    for (let index = 0; index < upToIndex; index += 1) {
      this.linearCellIndexes[index] = index
    }
    this.linearRowMajorWidth = width
    this.linearRowMajorCapacityLimit = this.valueKinds.length
    this.denseRowMajorWidth = null
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
  }

  protected setPreviewValue(row: number, column: number, value: LiteralInput): void {
    const index = previewIndex(row, column)
    if (index === -1) {
      return
    }
    this.previewValues[index] = value
    this.previewValueSet[index] = 1
  }

  protected hasPreviewValue(row: number, column: number): boolean {
    const index = previewIndex(row, column)
    return index !== -1 && this.previewValueSet[index] === 1
  }

  protected readPreviewValue(row: number, column: number): LiteralInput | undefined {
    const index = previewIndex(row, column)
    return index === -1 || this.previewValueSet[index] !== 1 ? undefined : this.previewValues[index]
  }

  protected ensureCapacity(nextLength: number): void {
    if (nextLength <= this.valueKinds.length) {
      return
    }
    let nextCapacity = this.valueKinds.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    if (
      this.linearRowMajorCapacityLimit !== null &&
      nextLength <= this.linearRowMajorCapacityLimit &&
      nextCapacity > this.linearRowMajorCapacityLimit
    ) {
      nextCapacity = this.linearRowMajorCapacityLimit
    }
    this.resizeStorage(nextCapacity)
  }

  protected resizeStorage(nextCapacity: number): void {
    if (this.sheetIndexes) {
      this.sheetIndexes = growUint32Array(this.sheetIndexes, nextCapacity)
    }
    if (this.linearCellIndexes) {
      this.linearCellIndexes = growUint32Array(this.linearCellIndexes, nextCapacity)
    } else if (this.denseRowMajorWidth === null && this.streamRowRunCoordinates) {
      if (this.columns.length > 0) {
        this.columns = growUint16Array(this.columns, nextCapacity)
      }
    } else if (this.denseRowMajorWidth === null) {
      if (this.rowRunRows) {
        this.materializeRowRunCoordinateStorage()
      }
      this.rows = growUint32Array(this.rows, nextCapacity)
      this.columns = growUint16Array(this.columns, nextCapacity)
    }
    this.valueKinds = growUint8Array(this.valueKinds, nextCapacity)
    if (this.numberValues) {
      this.numberValues = growFloat64Array(this.numberValues, nextCapacity)
    }
    if (this.tinyIntegerValues) {
      this.tinyIntegerValues = growInt8Array(this.tinyIntegerValues, nextCapacity)
    }
    if (this.smallIntegerValues) {
      this.smallIntegerValues = growInt16Array(this.smallIntegerValues, nextCapacity)
    }
    if (this.integerValues) {
      this.integerValues = growInt32Array(this.integerValues, nextCapacity)
    }
    if (this.stringIds) {
      this.stringIds = growStringIdStorage(this.stringIds, nextCapacity)
    }
    if (this.booleanValues) {
      this.booleanValues = growUint8Array(this.booleanValues, nextCapacity)
    }
    if (this.formulaIds) {
      this.formulaIds = growUint32Array(this.formulaIds, nextCapacity, noPoolId)
    }
  }
}
