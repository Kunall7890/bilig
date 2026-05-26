export interface EdgeSlice {
  ptr: number
  len: number
  cap: number
}

const EMPTY_SLICE: EdgeSlice = { ptr: -1, len: 0, cap: 0 }
const EMPTY_U32 = new Uint32Array(0)

export class EdgeArena {
  private buffer = new Uint32Array(64)
  private freeList: EdgeSlice[] = []
  private maxFreeCapacity = 0
  private nextPtr = 0

  reset(): void {
    this.buffer.fill(0)
    this.freeList = []
    this.maxFreeCapacity = 0
    this.nextPtr = 0
  }

  empty(): EdgeSlice {
    return EMPTY_SLICE
  }

  alloc(size: number): EdgeSlice {
    if (size <= 0) {
      return EMPTY_SLICE
    }

    if (this.freeList.length > 0 && this.maxFreeCapacity >= size) {
      const freeIndex = this.freeList.findIndex((slice) => slice.cap >= size)
      if (freeIndex !== -1) {
        const [slice] = this.freeList.splice(freeIndex, 1)
        if (slice!.cap === this.maxFreeCapacity) {
          this.recomputeMaxFreeCapacity()
        }
        return { ptr: slice!.ptr, len: 0, cap: slice!.cap }
      }
      this.recomputeMaxFreeCapacity()
    }

    const ptr = this.nextPtr
    this.ensureCapacity(ptr + size)
    this.nextPtr += size
    return { ptr, len: 0, cap: size }
  }

  replace(slice: EdgeSlice, nextValues: Uint32Array | readonly number[]): EdgeSlice {
    const values = nextValues instanceof Uint32Array ? nextValues : Uint32Array.from(nextValues)
    if (values.length === 0) {
      this.free(slice)
      return EMPTY_SLICE
    }

    let target = slice
    if (target.cap < values.length || target.ptr < 0) {
      this.free(slice)
      target = this.alloc(values.length)
    }

    this.buffer.set(values, target.ptr)
    return {
      ptr: target.ptr,
      len: values.length,
      cap: target.cap,
    }
  }

  replaceSmall(slice: EdgeSlice, length: 0 | 1 | 2, value0 = 0, value1 = 0): EdgeSlice {
    if (length === 0) {
      this.free(slice)
      return EMPTY_SLICE
    }

    let target = slice
    if (target.cap < length || target.ptr < 0) {
      this.free(slice)
      target = this.alloc(length)
    }

    this.buffer[target.ptr] = value0
    if (length === 2) {
      this.buffer[target.ptr + 1] = value1
    }
    return {
      ptr: target.ptr,
      len: length,
      cap: target.cap,
    }
  }

  read(slice: EdgeSlice): Uint32Array {
    if (slice.ptr < 0 || slice.len <= 0) {
      return EMPTY_U32
    }
    return this.buffer.slice(slice.ptr, slice.ptr + slice.len)
  }

  readView(slice: EdgeSlice): Uint32Array {
    if (slice.ptr < 0 || slice.len <= 0) {
      return EMPTY_U32
    }
    return this.buffer.subarray(slice.ptr, slice.ptr + slice.len)
  }

  singleton(value: number): EdgeSlice {
    const target = this.alloc(1)
    this.buffer[target.ptr] = value
    return {
      ptr: target.ptr,
      len: 1,
      cap: target.cap,
    }
  }

  valueAt(slice: EdgeSlice, index: number): number {
    if (slice.ptr < 0 || index < 0 || index >= slice.len) {
      return -1
    }
    return this.buffer[slice.ptr + index] ?? -1
  }

  view(): Uint32Array {
    return this.buffer.subarray(0, this.nextPtr)
  }

  appendUnique(slice: EdgeSlice, value: number): EdgeSlice {
    const values = this.readView(slice)
    for (let index = 0; index < values.length; index += 1) {
      if (values[index] === value) {
        return slice
      }
    }
    if (slice.ptr < 0 || slice.cap <= 0) {
      const target = this.alloc(1)
      this.buffer[target.ptr] = value
      return {
        ptr: target.ptr,
        len: 1,
        cap: target.cap,
      }
    }
    if (slice.len < slice.cap) {
      this.buffer[slice.ptr + slice.len] = value
      return {
        ptr: slice.ptr,
        len: slice.len + 1,
        cap: slice.cap,
      }
    }
    const nextCap = Math.max(slice.cap * 2, slice.len + 1)
    const target = this.alloc(nextCap)
    this.buffer.set(values, target.ptr)
    this.buffer[target.ptr + values.length] = value
    this.free(slice)
    return {
      ptr: target.ptr,
      len: values.length + 1,
      cap: target.cap,
    }
  }

  appendKnownUnique(slice: EdgeSlice, value: number): EdgeSlice {
    const values = this.readView(slice)
    if (slice.ptr < 0 || slice.cap <= 0) {
      const target = this.alloc(1)
      this.buffer[target.ptr] = value
      return {
        ptr: target.ptr,
        len: 1,
        cap: target.cap,
      }
    }
    if (slice.len < slice.cap) {
      this.buffer[slice.ptr + slice.len] = value
      return {
        ptr: slice.ptr,
        len: slice.len + 1,
        cap: slice.cap,
      }
    }
    const nextCap = Math.max(slice.cap * 2, slice.len + 1)
    const target = this.alloc(nextCap)
    this.buffer.set(values, target.ptr)
    this.buffer[target.ptr + values.length] = value
    this.free(slice)
    return {
      ptr: target.ptr,
      len: values.length + 1,
      cap: target.cap,
    }
  }

  removeValue(slice: EdgeSlice, value: number): EdgeSlice {
    const values = this.readView(slice)
    if (values.length === 0) {
      return slice
    }

    let found = false
    const next = new Uint32Array(values.length)
    let cursor = 0
    for (let index = 0; index < values.length; index += 1) {
      const current = values[index]!
      if (current === value) {
        found = true
        continue
      }
      next[cursor] = current
      cursor += 1
    }

    if (!found) {
      return slice
    }
    return this.replace(slice, next.subarray(0, cursor))
  }

  free(slice: EdgeSlice): void {
    if (slice.ptr < 0 || slice.cap <= 0) {
      return
    }
    this.freeList.push({
      ptr: slice.ptr,
      len: 0,
      cap: slice.cap,
    })
    this.maxFreeCapacity = Math.max(this.maxFreeCapacity, slice.cap)
  }

  private ensureCapacity(nextSize: number): void {
    if (nextSize <= this.buffer.length) {
      return
    }
    let capacity = this.buffer.length
    while (capacity < nextSize) {
      capacity *= 2
    }
    const next = new Uint32Array(capacity)
    next.set(this.buffer)
    this.buffer = next
  }

  private recomputeMaxFreeCapacity(): void {
    let maxCapacity = 0
    for (let index = 0; index < this.freeList.length; index += 1) {
      maxCapacity = Math.max(maxCapacity, this.freeList[index]!.cap)
    }
    this.maxFreeCapacity = maxCapacity
  }
}
