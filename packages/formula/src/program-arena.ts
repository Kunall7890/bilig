export interface ArenaSlice {
  offset: number
  length: number
}

const EMPTY_U32 = new Uint32Array(0)
const EMPTY_F64 = new Float64Array(0)

export class Uint32Arena {
  private data = EMPTY_U32
  private used = 0

  reset(): void {
    this.used = 0
  }

  append(values: Uint32Array): ArenaSlice {
    const slice = { offset: this.used, length: values.length }
    this.ensureCapacity(this.used + values.length)
    this.data.set(values, this.used)
    this.used += values.length
    return slice
  }

  view(): Uint32Array {
    return this.data.subarray(0, this.used)
  }

  get size(): number {
    return this.used
  }

  private ensureCapacity(required: number): void {
    if (required <= this.data.length) {
      return
    }
    let capacity = Math.max(this.data.length, 1)
    while (capacity < required) {
      capacity *= 2
    }
    const next = new Uint32Array(capacity)
    next.set(this.data)
    this.data = next
  }
}

export class Float64Arena {
  private data = EMPTY_F64
  private used = 0

  reset(): void {
    this.used = 0
  }

  append(values: ArrayLike<number>): ArenaSlice {
    const slice = { offset: this.used, length: values.length }
    this.ensureCapacity(this.used + values.length)
    for (let index = 0; index < values.length; index += 1) {
      this.data[this.used + index] = values[index]!
    }
    this.used += values.length
    return slice
  }

  view(): Float64Array {
    return this.data.subarray(0, this.used)
  }

  get size(): number {
    return this.used
  }

  private ensureCapacity(required: number): void {
    if (required <= this.data.length) {
      return
    }
    let capacity = Math.max(this.data.length, 1)
    while (capacity < required) {
      capacity *= 2
    }
    const next = new Float64Array(capacity)
    next.set(this.data)
    this.data = next
  }
}
