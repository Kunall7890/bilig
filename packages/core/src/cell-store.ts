import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { CellValue } from '@bilig/protocol'

const EMPTY_U8 = new Uint8Array(0)
const EMPTY_U16 = new Uint16Array(0)
const EMPTY_U32 = new Uint32Array(0)
const EMPTY_I32 = new Int32Array(0)
const EMPTY_F64 = new Float64Array(0)
const MAX_POOLED_CELL_STORE_CAPACITY = 131_072
const MAX_POOLED_CELL_STORE_BUFFER_SETS = 8

interface CellStoreBuffers {
  readonly capacity: number
  readonly tags: Uint8Array
  readonly numbers: Float64Array
  readonly stringIds: Uint32Array
  readonly errors: Uint16Array
  readonly formulaIds: Uint32Array
  readonly versions: Uint32Array
  readonly flags: Uint32Array
  readonly sheetIds: Uint16Array
  readonly rows: Uint32Array
  readonly cols: Uint16Array
  readonly topoRanks: Uint32Array
  readonly cycleGroupIds: Int32Array
}

const pooledCellStoreBuffers = new Map<number, CellStoreBuffers[]>()
let pooledCellStoreBufferSetCount = 0

export const enum CellFlags {
  HasFormula = 1 << 1,
  JsOnly = 1 << 2,
  InCycle = 1 << 3,
  Materialized = 1 << 4,
  PendingDelete = 1 << 5,
  SpillChild = 1 << 6,
  PivotOutput = 1 << 7,
  AuthoredBlank = 1 << 8,
}

export class CellStore {
  size = 0
  capacity: number
  onSetValue: ((index: number) => void) | null = null
  tags: Uint8Array = EMPTY_U8
  numbers: Float64Array = EMPTY_F64
  stringIds: Uint32Array = EMPTY_U32
  errors: Uint16Array = EMPTY_U16
  formulaIds: Uint32Array = EMPTY_U32
  versions: Uint32Array = EMPTY_U32
  flags: Uint32Array = EMPTY_U32
  sheetIds: Uint16Array = EMPTY_U16
  rows: Uint32Array = EMPTY_U32
  cols: Uint16Array = EMPTY_U16
  topoRanks: Uint32Array = EMPTY_U32
  cycleGroupIds: Int32Array = EMPTY_I32

  constructor(initialCapacity = 0) {
    this.capacity = initialCapacity
    if (initialCapacity === 0) {
      this.tags = EMPTY_U8
      this.numbers = EMPTY_F64
      this.stringIds = EMPTY_U32
      this.errors = EMPTY_U16
      this.formulaIds = EMPTY_U32
      this.versions = EMPTY_U32
      this.flags = EMPTY_U32
      this.sheetIds = EMPTY_U16
      this.rows = EMPTY_U32
      this.cols = EMPTY_U16
      this.topoRanks = EMPTY_U32
      this.cycleGroupIds = EMPTY_I32
    } else {
      this.assignBuffers(takeCellStoreBuffers(initialCapacity) ?? createCellStoreBuffers(initialCapacity))
    }
  }

  ensureCapacity(nextSize: number): void {
    if (nextSize <= this.capacity) return
    let nextCapacity = Math.max(this.capacity, 1)
    while (nextCapacity < nextSize) nextCapacity *= 2
    if (this.size === 0) {
      this.assignBuffers(takeCellStoreBuffers(nextCapacity) ?? createCellStoreBuffers(nextCapacity))
      return
    }
    this.tags = grow(this.tags, nextCapacity)
    this.numbers = grow(this.numbers, nextCapacity)
    this.stringIds = grow(this.stringIds, nextCapacity)
    this.errors = grow(this.errors, nextCapacity)
    this.formulaIds = grow(this.formulaIds, nextCapacity)
    this.versions = grow(this.versions, nextCapacity)
    this.flags = grow(this.flags, nextCapacity)
    this.sheetIds = grow(this.sheetIds, nextCapacity)
    this.rows = grow(this.rows, nextCapacity)
    this.cols = grow(this.cols, nextCapacity)
    this.topoRanks = grow(this.topoRanks, nextCapacity)
    const nextCycle = grow(this.cycleGroupIds, nextCapacity)
    nextCycle.fill(-1, this.capacity)
    this.cycleGroupIds = nextCycle
    this.capacity = nextCapacity
  }

  allocate(sheetId: number, row: number, col: number): number {
    this.ensureCapacity(this.size + 1)
    return this.allocateReserved(sheetId, row, col)
  }

  allocateReserved(sheetId: number, row: number, col: number): number {
    const index = this.size
    this.size += 1
    this.sheetIds[index] = sheetId
    this.rows[index] = row
    this.cols[index] = col
    this.tags[index] = ValueTag.Empty
    this.errors[index] = ErrorCode.None
    this.flags[index] = CellFlags.Materialized
    return index
  }

  allocateDenseRowMajorReserved(sheetId: number, rowCount: number, colCount: number): number {
    return this.allocateDenseRowMajorAtReserved(sheetId, 0, rowCount, 0, colCount)
  }

  allocateDenseRowMajorAtReserved(sheetId: number, rowStart: number, rowCount: number, colStart: number, colCount: number): number {
    const count = rowCount * colCount
    const firstIndex = this.size
    if (count <= 0) {
      return firstIndex
    }
    this.ensureCapacity(firstIndex + count)
    this.size = firstIndex + count

    this.sheetIds.fill(sheetId, firstIndex, firstIndex + count)
    this.tags.fill(ValueTag.Empty, firstIndex, firstIndex + count)
    this.errors.fill(ErrorCode.None, firstIndex, firstIndex + count)
    this.flags.fill(CellFlags.Materialized, firstIndex, firstIndex + count)

    if (colCount <= 4) {
      for (let row = 0; row < rowCount; row += 1) {
        const rowBase = firstIndex + row * colCount
        const physicalRow = rowStart + row
        for (let colOffset = 0; colOffset < colCount; colOffset += 1) {
          const index = rowBase + colOffset
          this.rows[index] = physicalRow
          this.cols[index] = colStart + colOffset
        }
      }
      return firstIndex
    }
    const colPattern = materializeDenseColumnPattern(colStart, colCount)
    for (let row = 0; row < rowCount; row += 1) {
      const rowBase = firstIndex + row * colCount
      this.rows.fill(rowStart + row, rowBase, rowBase + colCount)
      this.cols.set(colPattern, rowBase)
    }
    return firstIndex
  }

  allocateDenseSingleColumnReserved(sheetId: number, rowStart: number, rowCount: number, col: number): number {
    const firstIndex = this.size
    if (rowCount <= 0) {
      return firstIndex
    }
    this.ensureCapacity(firstIndex + rowCount)
    this.size = firstIndex + rowCount

    this.sheetIds.fill(sheetId, firstIndex, firstIndex + rowCount)
    this.cols.fill(col, firstIndex, firstIndex + rowCount)
    this.tags.fill(ValueTag.Empty, firstIndex, firstIndex + rowCount)
    this.errors.fill(ErrorCode.None, firstIndex, firstIndex + rowCount)
    this.flags.fill(CellFlags.Materialized, firstIndex, firstIndex + rowCount)

    for (let offset = 0; offset < rowCount; offset += 1) {
      this.rows[firstIndex + offset] = rowStart + offset
    }
    return firstIndex
  }

  reset(): void {
    this.size = 0
    clearCellStoreBuffers(this)
  }

  releaseBuffersToPool(): void {
    if (this.capacity <= 0) {
      this.size = 0
      this.onSetValue = null
      return
    }
    clearCellStoreBuffers(this)
    const buffers: CellStoreBuffers = {
      capacity: this.capacity,
      tags: this.tags,
      numbers: this.numbers,
      stringIds: this.stringIds,
      errors: this.errors,
      formulaIds: this.formulaIds,
      versions: this.versions,
      flags: this.flags,
      sheetIds: this.sheetIds,
      rows: this.rows,
      cols: this.cols,
      topoRanks: this.topoRanks,
      cycleGroupIds: this.cycleGroupIds,
    }
    releaseCellStoreBuffers(buffers)
    this.size = 0
    this.capacity = 0
    this.onSetValue = null
    this.tags = EMPTY_U8
    this.numbers = EMPTY_F64
    this.stringIds = EMPTY_U32
    this.errors = EMPTY_U16
    this.formulaIds = EMPTY_U32
    this.versions = EMPTY_U32
    this.flags = EMPTY_U32
    this.sheetIds = EMPTY_U16
    this.rows = EMPTY_U32
    this.cols = EMPTY_U16
    this.topoRanks = EMPTY_U32
    this.cycleGroupIds = EMPTY_I32
  }

  setValue(index: number, value: CellValue, stringId = 0): void {
    this.tags[index] = value.tag
    this.errors[index] = value.tag === ValueTag.Error ? value.code : ErrorCode.None
    this.stringIds[index] = value.tag === ValueTag.String ? stringId : 0
    this.numbers[index] = value.tag === ValueTag.Number ? value.value : value.tag === ValueTag.Boolean ? (value.value ? 1 : 0) : 0
    this.versions[index] = (this.versions[index] ?? 0) + 1
    this.onSetValue?.(index)
  }

  getValue(index: number, stringLookup: (id: number) => string): CellValue {
    const rawTag = this.tags[index]
    const readValue = rawTag === undefined ? undefined : valueReaders[rawTag]
    if (!readValue) {
      return { tag: ValueTag.Empty }
    }
    return readValue(this, index, stringLookup)
  }

  private assignBuffers(buffers: CellStoreBuffers): void {
    this.capacity = buffers.capacity
    this.tags = buffers.tags
    this.numbers = buffers.numbers
    this.stringIds = buffers.stringIds
    this.errors = buffers.errors
    this.formulaIds = buffers.formulaIds
    this.versions = buffers.versions
    this.flags = buffers.flags
    this.sheetIds = buffers.sheetIds
    this.rows = buffers.rows
    this.cols = buffers.cols
    this.topoRanks = buffers.topoRanks
    this.cycleGroupIds = buffers.cycleGroupIds
  }
}

function createCellStoreBuffers(capacity: number): CellStoreBuffers {
  const buffers = {
    capacity,
    tags: new Uint8Array(capacity),
    numbers: new Float64Array(capacity),
    stringIds: new Uint32Array(capacity),
    errors: new Uint16Array(capacity),
    formulaIds: new Uint32Array(capacity),
    versions: new Uint32Array(capacity),
    flags: new Uint32Array(capacity),
    sheetIds: new Uint16Array(capacity),
    rows: new Uint32Array(capacity),
    cols: new Uint16Array(capacity),
    topoRanks: new Uint32Array(capacity),
    cycleGroupIds: new Int32Array(capacity),
  }
  buffers.cycleGroupIds.fill(-1)
  return buffers
}

function takeCellStoreBuffers(capacity: number): CellStoreBuffers | undefined {
  const buffers = pooledCellStoreBuffers.get(capacity)
  const bufferSet = buffers?.pop()
  if (bufferSet !== undefined) {
    pooledCellStoreBufferSetCount -= 1
  }
  if (buffers?.length === 0) {
    pooledCellStoreBuffers.delete(capacity)
  }
  return bufferSet
}

function releaseCellStoreBuffers(buffers: CellStoreBuffers): void {
  if (
    buffers.capacity <= 0 ||
    buffers.capacity > MAX_POOLED_CELL_STORE_CAPACITY ||
    pooledCellStoreBufferSetCount >= MAX_POOLED_CELL_STORE_BUFFER_SETS
  ) {
    return
  }
  const pool = pooledCellStoreBuffers.get(buffers.capacity) ?? []
  pool.push(buffers)
  pooledCellStoreBuffers.set(buffers.capacity, pool)
  pooledCellStoreBufferSetCount += 1
}

function clearCellStoreBuffers(buffers: {
  readonly tags: Uint8Array
  readonly numbers: Float64Array
  readonly stringIds: Uint32Array
  readonly errors: Uint16Array
  readonly formulaIds: Uint32Array
  readonly versions: Uint32Array
  readonly flags: Uint32Array
  readonly sheetIds: Uint16Array
  readonly rows: Uint32Array
  readonly cols: Uint16Array
  readonly topoRanks: Uint32Array
  readonly cycleGroupIds: Int32Array
}): void {
  buffers.tags.fill(0)
  buffers.numbers.fill(0)
  buffers.stringIds.fill(0)
  buffers.errors.fill(0)
  buffers.formulaIds.fill(0)
  buffers.versions.fill(0)
  buffers.flags.fill(0)
  buffers.sheetIds.fill(0)
  buffers.rows.fill(0)
  buffers.cols.fill(0)
  buffers.topoRanks.fill(0)
  buffers.cycleGroupIds.fill(-1)
}

const valueReaders: Array<((store: CellStore, index: number, stringLookup: (id: number) => string) => CellValue) | undefined> = []

valueReaders[ValueTag.Empty] = () => ({ tag: ValueTag.Empty })
valueReaders[ValueTag.Number] = (store, index) => ({
  tag: ValueTag.Number,
  value: store.numbers[index]!,
})
valueReaders[ValueTag.Boolean] = (store, index) => ({
  tag: ValueTag.Boolean,
  value: store.numbers[index]! !== 0,
})
valueReaders[ValueTag.String] = (store, index, stringLookup) => ({
  tag: ValueTag.String,
  value: stringLookup(store.stringIds[index]!),
  stringId: store.stringIds[index]!,
})
valueReaders[ValueTag.Error] = (store, index) => ({
  tag: ValueTag.Error,
  code: store.errors[index]!,
})

function grow(buffer: Uint8Array, capacity: number): Uint8Array
function grow(buffer: Uint16Array, capacity: number): Uint16Array
function grow(buffer: Uint32Array, capacity: number): Uint32Array
function grow(buffer: Int32Array, capacity: number): Int32Array
function grow(buffer: Float64Array, capacity: number): Float64Array
function grow(
  buffer: Uint8Array | Uint16Array | Uint32Array | Int32Array | Float64Array,
  capacity: number,
): Uint8Array | Uint16Array | Uint32Array | Int32Array | Float64Array {
  if (buffer instanceof Uint8Array) {
    const next = new Uint8Array(capacity)
    next.set(buffer)
    return next
  }
  if (buffer instanceof Uint16Array) {
    const next = new Uint16Array(capacity)
    next.set(buffer)
    return next
  }
  if (buffer instanceof Uint32Array) {
    const next = new Uint32Array(capacity)
    next.set(buffer)
    return next
  }
  if (buffer instanceof Int32Array) {
    const next = new Int32Array(capacity)
    next.set(buffer)
    return next
  }
  const next = new Float64Array(capacity)
  next.set(buffer)
  return next
}

function materializeDenseColumnPattern(colStart: number, colCount: number): Uint16Array {
  const pattern = new Uint16Array(colCount)
  for (let col = 0; col < colCount; col += 1) {
    pattern[col] = colStart + col
  }
  return pattern
}
