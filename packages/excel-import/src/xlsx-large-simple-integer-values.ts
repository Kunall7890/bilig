import { initialSparseIntegerCapacity } from './xlsx-large-simple-arena-constants.js'
import { binarySearchUint32Prefix } from './xlsx-large-simple-arena-helpers.js'
import { growInt16Array, growInt32Array, growUint32Array } from './xlsx-large-simple-array-storage.js'

export class SparseInt16ArenaValues {
  private denseValues: Int16Array<ArrayBuffer> | undefined
  private sparseCellIndexes: Uint32Array<ArrayBuffer> | undefined
  private sparseValues: Int16Array<ArrayBuffer> | undefined
  private sparseCount = 0

  retainedStorageByteLength(): number {
    return (this.denseValues?.byteLength ?? 0) + (this.sparseCellIndexes?.byteLength ?? 0) + (this.sparseValues?.byteLength ?? 0)
  }

  snapshot(length: number): {
    readonly smallIntegerValues?: Int16Array
    readonly sparseSmallIntegerCellIndexes?: Uint32Array
    readonly sparseSmallIntegerValues?: Int16Array
  } {
    return {
      ...(this.denseValues ? { smallIntegerValues: this.denseValues.subarray(0, length) } : {}),
      ...(this.sparseCount > 0 && this.sparseCellIndexes
        ? { sparseSmallIntegerCellIndexes: this.sparseCellIndexes.subarray(0, this.sparseCount) }
        : {}),
      ...(this.sparseCount > 0 && this.sparseValues ? { sparseSmallIntegerValues: this.sparseValues.subarray(0, this.sparseCount) } : {}),
    }
  }

  release(): void {
    this.denseValues = undefined
    this.sparseCellIndexes = undefined
    this.sparseValues = undefined
    this.sparseCount = 0
  }

  compact(length: number): void {
    if (this.denseValues && this.denseValues.length !== length) {
      this.denseValues = this.denseValues.slice(0, length)
    }
    if (this.sparseCellIndexes && this.sparseCellIndexes.length !== this.sparseCount) {
      this.sparseCellIndexes = this.sparseCellIndexes.slice(0, this.sparseCount)
    }
    if (this.sparseValues && this.sparseValues.length !== this.sparseCount) {
      this.sparseValues = this.sparseValues.slice(0, this.sparseCount)
    }
  }

  resize(nextCapacity: number): void {
    if (this.denseValues) {
      this.denseValues = growInt16Array(this.denseValues, nextCapacity)
    }
  }

  store(index: number, value: number, capacity: number, denseThreshold: number): void {
    if (this.denseValues) {
      this.denseValues[index] = value
      return
    }
    if (this.sparseCount >= denseThreshold) {
      this.ensureDense(capacity)[index] = value
      return
    }
    if (!this.sparseCellIndexes || !this.sparseValues) {
      this.sparseCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseValues = new Int16Array(initialSparseIntegerCapacity)
    } else if (this.sparseCount >= this.sparseCellIndexes.length) {
      const nextSparseCapacity = this.sparseCellIndexes.length * 2
      this.sparseCellIndexes = growUint32Array(this.sparseCellIndexes, nextSparseCapacity)
      this.sparseValues = growInt16Array(this.sparseValues, nextSparseCapacity)
    }
    this.sparseCellIndexes[this.sparseCount] = index
    this.sparseValues[this.sparseCount] = value
    this.sparseCount += 1
  }

  valueAt(index: number): number | undefined {
    if (this.denseValues) {
      return this.denseValues[index]
    }
    if (!this.sparseCellIndexes || !this.sparseValues || this.sparseCount === 0) {
      return undefined
    }
    const offset = binarySearchUint32Prefix(this.sparseCellIndexes, this.sparseCount, index)
    return offset === -1 ? undefined : this.sparseValues[offset]
  }

  private ensureDense(capacity: number): Int16Array<ArrayBuffer> {
    if (this.denseValues) {
      return this.denseValues
    }
    this.denseValues = new Int16Array(capacity)
    this.copySparseIntoDense(this.denseValues)
    this.sparseCellIndexes = undefined
    this.sparseValues = undefined
    this.sparseCount = 0
    return this.denseValues
  }

  private copySparseIntoDense(output: Int16Array<ArrayBuffer>): void {
    if (!this.sparseCellIndexes || !this.sparseValues) {
      return
    }
    for (let offset = 0; offset < this.sparseCount; offset += 1) {
      const cellIndex = this.sparseCellIndexes[offset] ?? -1
      if (cellIndex >= 0 && cellIndex < output.length) {
        output[cellIndex] = this.sparseValues[offset] ?? 0
      }
    }
  }
}

export class SparseInt32ArenaValues {
  private denseValues: Int32Array<ArrayBuffer> | undefined
  private sparseCellIndexes: Uint32Array<ArrayBuffer> | undefined
  private sparseValues: Int32Array<ArrayBuffer> | undefined
  private sparseCount = 0

  retainedStorageByteLength(): number {
    return (this.denseValues?.byteLength ?? 0) + (this.sparseCellIndexes?.byteLength ?? 0) + (this.sparseValues?.byteLength ?? 0)
  }

  snapshot(length: number): {
    readonly integerValues?: Int32Array
    readonly sparseIntegerCellIndexes?: Uint32Array
    readonly sparseIntegerValues?: Int32Array
  } {
    return {
      ...(this.denseValues ? { integerValues: this.denseValues.subarray(0, length) } : {}),
      ...(this.sparseCount > 0 && this.sparseCellIndexes
        ? { sparseIntegerCellIndexes: this.sparseCellIndexes.subarray(0, this.sparseCount) }
        : {}),
      ...(this.sparseCount > 0 && this.sparseValues ? { sparseIntegerValues: this.sparseValues.subarray(0, this.sparseCount) } : {}),
    }
  }

  release(): void {
    this.denseValues = undefined
    this.sparseCellIndexes = undefined
    this.sparseValues = undefined
    this.sparseCount = 0
  }

  compact(length: number): void {
    if (this.denseValues && this.denseValues.length !== length) {
      this.denseValues = this.denseValues.slice(0, length)
    }
    if (this.sparseCellIndexes && this.sparseCellIndexes.length !== this.sparseCount) {
      this.sparseCellIndexes = this.sparseCellIndexes.slice(0, this.sparseCount)
    }
    if (this.sparseValues && this.sparseValues.length !== this.sparseCount) {
      this.sparseValues = this.sparseValues.slice(0, this.sparseCount)
    }
  }

  resize(nextCapacity: number): void {
    if (this.denseValues) {
      this.denseValues = growInt32Array(this.denseValues, nextCapacity)
    }
  }

  store(index: number, value: number, capacity: number, denseThreshold: number): void {
    if (this.denseValues) {
      this.denseValues[index] = value
      return
    }
    if (this.sparseCount >= denseThreshold) {
      this.ensureDense(capacity)[index] = value
      return
    }
    if (!this.sparseCellIndexes || !this.sparseValues) {
      this.sparseCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseValues = new Int32Array(initialSparseIntegerCapacity)
    } else if (this.sparseCount >= this.sparseCellIndexes.length) {
      const nextSparseCapacity = this.sparseCellIndexes.length * 2
      this.sparseCellIndexes = growUint32Array(this.sparseCellIndexes, nextSparseCapacity)
      this.sparseValues = growInt32Array(this.sparseValues, nextSparseCapacity)
    }
    this.sparseCellIndexes[this.sparseCount] = index
    this.sparseValues[this.sparseCount] = value
    this.sparseCount += 1
  }

  valueAt(index: number): number | undefined {
    if (this.denseValues) {
      return this.denseValues[index]
    }
    if (!this.sparseCellIndexes || !this.sparseValues || this.sparseCount === 0) {
      return undefined
    }
    const offset = binarySearchUint32Prefix(this.sparseCellIndexes, this.sparseCount, index)
    return offset === -1 ? undefined : this.sparseValues[offset]
  }

  private ensureDense(capacity: number): Int32Array<ArrayBuffer> {
    if (this.denseValues) {
      return this.denseValues
    }
    this.denseValues = new Int32Array(capacity)
    this.copySparseIntoDense(this.denseValues)
    this.sparseCellIndexes = undefined
    this.sparseValues = undefined
    this.sparseCount = 0
    return this.denseValues
  }

  private copySparseIntoDense(output: Int32Array<ArrayBuffer>): void {
    if (!this.sparseCellIndexes || !this.sparseValues) {
      return
    }
    for (let offset = 0; offset < this.sparseCount; offset += 1) {
      const cellIndex = this.sparseCellIndexes[offset] ?? -1
      if (cellIndex >= 0 && cellIndex < output.length) {
        output[cellIndex] = this.sparseValues[offset] ?? 0
      }
    }
  }
}
