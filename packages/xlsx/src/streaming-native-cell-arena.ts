import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

export interface StreamingNativeSharedStringReference {
  readonly kind: 'shared-string'
  readonly index: number
}

export type StreamingNativePendingCellValue = CellValue | StreamingNativeSharedStringReference

export interface StreamingNativePendingCellRow {
  readonly size: number
  get(col: number): StreamingNativePendingCellValue | undefined
  entries(): IterableIterator<[number, StreamingNativePendingCellValue]>
  values(): IterableIterator<StreamingNativePendingCellValue>
  [Symbol.iterator](): IterableIterator<[number, StreamingNativePendingCellValue]>
}

export interface StreamingNativeMutablePendingCellRow extends StreamingNativePendingCellRow {
  set(col: number, value: StreamingNativePendingCellValue): this
}

export interface StreamingNativePendingCellRows {
  get(row: number): StreamingNativePendingCellRow | undefined
  entries(): IterableIterator<[number, StreamingNativePendingCellRow]>
  values(): IterableIterator<StreamingNativePendingCellRow>
  [Symbol.iterator](): IterableIterator<[number, StreamingNativePendingCellRow]>
}

export interface StreamingNativeMutablePendingCellRows extends StreamingNativePendingCellRows {
  get(row: number): StreamingNativeMutablePendingCellRow | undefined
  getOrCreate(row: number): StreamingNativeMutablePendingCellRow
  set(row: number, values: StreamingNativePendingCellRow): this
  clear(): void
}

const storedCellTag = {
  Empty: 0,
  Number: 1,
  Boolean: 2,
  String: 3,
  Error: 4,
  SharedString: 5,
} as const

const errorCodeByStoredNumber: Readonly<Record<number, ErrorCode>> = {
  [ErrorCode.None]: ErrorCode.None,
  [ErrorCode.Div0]: ErrorCode.Div0,
  [ErrorCode.Ref]: ErrorCode.Ref,
  [ErrorCode.Value]: ErrorCode.Value,
  [ErrorCode.Name]: ErrorCode.Name,
  [ErrorCode.NA]: ErrorCode.NA,
  [ErrorCode.Cycle]: ErrorCode.Cycle,
  [ErrorCode.Spill]: ErrorCode.Spill,
  [ErrorCode.Blocked]: ErrorCode.Blocked,
  [ErrorCode.Num]: ErrorCode.Num,
  [ErrorCode.Field]: ErrorCode.Field,
  [ErrorCode.Null]: ErrorCode.Null,
}

const initialCapacity = 64

export class StreamingNativeSheetCellArena implements StreamingNativeMutablePendingCellRows {
  private rowNumbers: Int32Array = new Int32Array(initialCapacity)
  private columnNumbers: Int32Array = new Int32Array(initialCapacity)
  private valueTags: Uint8Array = new Uint8Array(initialCapacity)
  private numberValues: Float64Array = new Float64Array(initialCapacity)
  private stringIds: Int32Array = new Int32Array(initialCapacity)
  private errorCodes: Int16Array = new Int16Array(initialCapacity)
  private stringValues: (string | undefined)[] = []
  private readonly slotsByRow = new Map<number, number[]>()
  private readonly rowViews = new Map<number, StreamingNativeCellRowView>()
  private cellCount = 0

  get(row: number): StreamingNativeMutablePendingCellRow | undefined {
    return this.rowViews.get(row)
  }

  getOrCreate(row: number): StreamingNativeMutablePendingCellRow {
    let rowView = this.rowViews.get(row)
    if (!rowView) {
      rowView = new StreamingNativeCellRowView(this, row)
      this.rowViews.set(row, rowView)
      this.slotsByRow.set(row, [])
    }
    return rowView
  }

  set(row: number, values: StreamingNativePendingCellRow): this {
    const target = this.getOrCreate(row)
    for (const [col, value] of values.entries()) {
      target.set(col, value)
    }
    return this
  }

  *entries(): IterableIterator<[number, StreamingNativeMutablePendingCellRow]> {
    for (const [row, rowView] of this.rowViews.entries()) {
      yield [row, rowView]
    }
  }

  values(): IterableIterator<StreamingNativeMutablePendingCellRow> {
    return this.rowViews.values()
  }

  [Symbol.iterator](): IterableIterator<[number, StreamingNativeMutablePendingCellRow]> {
    return this.entries()
  }

  clear(): void {
    for (let index = 0; index < this.cellCount; index += 1) {
      this.stringValues[index] = undefined
    }
    this.cellCount = 0
    this.slotsByRow.clear()
    this.rowViews.clear()
  }

  rowSize(row: number): number {
    return this.slotsByRow.get(row)?.length ?? 0
  }

  getCell(row: number, col: number): StreamingNativePendingCellValue | undefined {
    const slot = this.slotForCell(row, col)
    return slot === undefined ? undefined : this.valueAtSlot(slot)
  }

  setCell(row: number, col: number, value: StreamingNativePendingCellValue): void {
    const slot = this.slotForCell(row, col) ?? this.appendCell(row, col)
    this.storeValue(slot, value)
  }

  *rowEntries(row: number): IterableIterator<[number, StreamingNativePendingCellValue]> {
    for (const slot of this.slotsByRow.get(row) ?? []) {
      yield [this.columnNumbers[slot]!, this.valueAtSlot(slot)]
    }
  }

  *rowValues(row: number): IterableIterator<StreamingNativePendingCellValue> {
    for (const slot of this.slotsByRow.get(row) ?? []) {
      yield this.valueAtSlot(slot)
    }
  }

  private slotForCell(row: number, col: number): number | undefined {
    const slots = this.slotsByRow.get(row)
    if (!slots) {
      return undefined
    }
    return slots.find((slot) => this.columnNumbers[slot] === col)
  }

  private appendCell(row: number, col: number): number {
    this.ensureCapacity(this.cellCount + 1)
    const slot = this.cellCount
    this.cellCount += 1
    this.rowNumbers[slot] = row
    this.columnNumbers[slot] = col
    const slots = this.slotsByRow.get(row) ?? []
    slots.push(slot)
    this.slotsByRow.set(row, slots)
    if (!this.rowViews.has(row)) {
      this.rowViews.set(row, new StreamingNativeCellRowView(this, row))
    }
    return slot
  }

  private ensureCapacity(requiredCapacity: number): void {
    if (requiredCapacity <= this.rowNumbers.length) {
      return
    }
    let nextCapacity = this.rowNumbers.length
    while (nextCapacity < requiredCapacity) {
      nextCapacity *= 2
    }
    this.rowNumbers = growInt32Array(this.rowNumbers, nextCapacity)
    this.columnNumbers = growInt32Array(this.columnNumbers, nextCapacity)
    this.valueTags = growUint8Array(this.valueTags, nextCapacity)
    this.numberValues = growFloat64Array(this.numberValues, nextCapacity)
    this.stringIds = growInt32Array(this.stringIds, nextCapacity)
    this.errorCodes = growInt16Array(this.errorCodes, nextCapacity)
  }

  private storeValue(slot: number, value: StreamingNativePendingCellValue): void {
    this.stringValues[slot] = undefined
    if (isSharedStringReference(value)) {
      this.valueTags[slot] = storedCellTag.SharedString
      this.stringIds[slot] = value.index
      return
    }
    switch (value.tag) {
      case ValueTag.Empty:
        this.valueTags[slot] = storedCellTag.Empty
        return
      case ValueTag.Number:
        this.valueTags[slot] = storedCellTag.Number
        this.numberValues[slot] = value.value
        return
      case ValueTag.Boolean:
        this.valueTags[slot] = storedCellTag.Boolean
        this.numberValues[slot] = value.value ? 1 : 0
        return
      case ValueTag.String:
        this.valueTags[slot] = storedCellTag.String
        this.stringValues[slot] = value.value
        this.stringIds[slot] = value.stringId ?? 0
        return
      case ValueTag.Error:
        this.valueTags[slot] = storedCellTag.Error
        this.errorCodes[slot] = value.code
        return
    }
  }

  private valueAtSlot(slot: number): StreamingNativePendingCellValue {
    const tag = this.valueTags[slot] ?? storedCellTag.Empty
    switch (tag) {
      case storedCellTag.Number:
        return { tag: ValueTag.Number, value: this.numberValues[slot]! }
      case storedCellTag.Boolean:
        return { tag: ValueTag.Boolean, value: this.numberValues[slot] === 1 }
      case storedCellTag.String:
        return { tag: ValueTag.String, value: this.stringValues[slot] ?? '', stringId: this.stringIds[slot] ?? 0 }
      case storedCellTag.Error:
        return { tag: ValueTag.Error, code: errorCodeForStoredCode(this.errorCodes[slot] ?? ErrorCode.Value) }
      case storedCellTag.SharedString:
        return { kind: 'shared-string', index: this.stringIds[slot] ?? 0 }
      case storedCellTag.Empty:
      default:
        return { tag: ValueTag.Empty }
    }
  }
}

class StreamingNativeCellRowView implements StreamingNativeMutablePendingCellRow {
  constructor(
    private readonly arena: StreamingNativeSheetCellArena,
    private readonly row: number,
  ) {}

  get size(): number {
    return this.arena.rowSize(this.row)
  }

  get(col: number): StreamingNativePendingCellValue | undefined {
    return this.arena.getCell(this.row, col)
  }

  set(col: number, value: StreamingNativePendingCellValue): this {
    this.arena.setCell(this.row, col, value)
    return this
  }

  entries(): IterableIterator<[number, StreamingNativePendingCellValue]> {
    return this.arena.rowEntries(this.row)
  }

  values(): IterableIterator<StreamingNativePendingCellValue> {
    return this.arena.rowValues(this.row)
  }

  [Symbol.iterator](): IterableIterator<[number, StreamingNativePendingCellValue]> {
    return this.entries()
  }
}

function isSharedStringReference(value: StreamingNativePendingCellValue): value is StreamingNativeSharedStringReference {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'shared-string'
}

function errorCodeForStoredCode(code: number): ErrorCode {
  return errorCodeByStoredNumber[code] ?? ErrorCode.Value
}

function growInt32Array(source: Int32Array, length: number): Int32Array {
  const next = new Int32Array(length)
  next.set(source)
  return next
}

function growInt16Array(source: Int16Array, length: number): Int16Array {
  const next = new Int16Array(length)
  next.set(source)
  return next
}

function growUint8Array(source: Uint8Array, length: number): Uint8Array {
  const next = new Uint8Array(length)
  next.set(source)
  return next
}

function growFloat64Array(source: Float64Array, length: number): Float64Array {
  const next = new Float64Array(length)
  next.set(source)
  return next
}
