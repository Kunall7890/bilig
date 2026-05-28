import type { InitialFormulaCellIndexList } from './formula-initialization-refs.js'

const DENSE_MEMBERSHIP_SMALL_CAPACITY = 256
const DENSE_MEMBERSHIP_MAX_SPARSE_FACTOR = 8
const SMALL_LIST_MEMBERSHIP_MAX_CELLS = 32

export interface InitialFormulaCellMembership {
  readonly has: (cellIndex: number) => boolean
  readonly add: (cellIndex: number) => void
  readonly delete: (cellIndex: number) => void
}

export function createInitialFormulaCellMembership(args: {
  readonly cellIndices?: InitialFormulaCellIndexList
  readonly cellCount?: number
  readonly maxCellIndex?: number
  readonly expectedCellCount?: number
}): InitialFormulaCellMembership {
  const cellCount = Math.min(args.cellCount ?? args.cellIndices?.length ?? 0, args.cellIndices?.length ?? 0)
  const expectedCellCount = Math.max(args.expectedCellCount ?? cellCount, cellCount)
  const maxCellIndex = args.maxCellIndex ?? findMaxCellIndex(args.cellIndices, cellCount)
  if (expectedCellCount <= SMALL_LIST_MEMBERSHIP_MAX_CELLS) {
    return new SmallInitialFormulaCellMembership(args.cellIndices, cellCount, expectedCellCount)
  }
  if (shouldUseDenseMembership(maxCellIndex, expectedCellCount)) {
    return new DenseInitialFormulaCellMembership(args.cellIndices, cellCount, maxCellIndex + 1)
  }

  return new SetInitialFormulaCellMembership(args.cellIndices, cellCount)
}

class SmallInitialFormulaCellMembership implements InitialFormulaCellMembership {
  private cells: Int32Array
  private length = 0

  constructor(cellIndices: InitialFormulaCellIndexList | undefined, cellCount: number, expectedCellCount: number) {
    this.cells = new Int32Array(Math.max(expectedCellCount, 1))
    if (cellIndices === undefined || cellCount === 0) {
      return
    }
    for (let index = 0; index < cellCount; index += 1) {
      this.add(cellIndices[index]!)
    }
  }

  has(cellIndex: number): boolean {
    for (let index = 0; index < this.length; index += 1) {
      if (this.cells[index] === cellIndex) {
        return true
      }
    }
    return false
  }

  add(cellIndex: number): void {
    if (this.has(cellIndex)) {
      return
    }
    for (let index = 0; index < this.length; index += 1) {
      if (this.cells[index] === -1) {
        this.cells[index] = cellIndex
        return
      }
    }
    if (this.length === this.cells.length) {
      const next = new Int32Array(this.cells.length * 2)
      next.set(this.cells)
      this.cells = next
    }
    this.cells[this.length] = cellIndex
    this.length += 1
  }

  delete(cellIndex: number): void {
    for (let index = 0; index < this.length; index += 1) {
      if (this.cells[index] === cellIndex) {
        this.cells[index] = -1
        return
      }
    }
  }
}

class DenseInitialFormulaCellMembership implements InitialFormulaCellMembership {
  private cells: Uint8Array

  constructor(cellIndices: InitialFormulaCellIndexList | undefined, cellCount: number, capacity: number) {
    this.cells = new Uint8Array(capacity)
    if (cellIndices === undefined) {
      return
    }
    for (let index = 0; index < cellCount; index += 1) {
      this.cells[cellIndices[index]!] = 1
    }
  }

  has(cellIndex: number): boolean {
    return (this.cells[cellIndex] ?? 0) !== 0
  }

  add(cellIndex: number): void {
    if (cellIndex >= this.cells.length) {
      const next = new Uint8Array(cellIndex + 1)
      next.set(this.cells)
      this.cells = next
    }
    this.cells[cellIndex] = 1
  }

  delete(cellIndex: number): void {
    this.cells[cellIndex] = 0
  }
}

class SetInitialFormulaCellMembership implements InitialFormulaCellMembership {
  private readonly cells = new Set<number>()

  constructor(cellIndices: InitialFormulaCellIndexList | undefined, cellCount: number) {
    if (cellIndices === undefined) {
      return
    }
    for (let index = 0; index < cellCount; index += 1) {
      this.cells.add(cellIndices[index]!)
    }
  }

  has(cellIndex: number): boolean {
    return this.cells.has(cellIndex)
  }

  add(cellIndex: number): void {
    this.cells.add(cellIndex)
  }

  delete(cellIndex: number): void {
    this.cells.delete(cellIndex)
  }
}

function shouldUseDenseMembership(maxCellIndex: number, expectedCellCount: number): boolean {
  if (maxCellIndex < 0) {
    return false
  }
  const capacity = maxCellIndex + 1
  return capacity <= DENSE_MEMBERSHIP_SMALL_CAPACITY || capacity <= expectedCellCount * DENSE_MEMBERSHIP_MAX_SPARSE_FACTOR
}

function findMaxCellIndex(cellIndices: InitialFormulaCellIndexList | undefined, cellCount: number): number {
  let maxCellIndex = -1
  if (cellIndices === undefined) {
    return maxCellIndex
  }
  for (let index = 0; index < cellCount; index += 1) {
    const cellIndex = cellIndices[index]!
    if (cellIndex > maxCellIndex) {
      maxCellIndex = cellIndex
    }
  }
  return maxCellIndex
}
