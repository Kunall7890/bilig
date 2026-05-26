import type { InitialFormulaCellIndexList } from './formula-initialization-refs.js'

const DENSE_MEMBERSHIP_SMALL_CAPACITY = 256
const DENSE_MEMBERSHIP_MAX_SPARSE_FACTOR = 8

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
  if (shouldUseDenseMembership(maxCellIndex, expectedCellCount)) {
    let cells = new Uint8Array(maxCellIndex + 1)
    for (let index = 0; index < cellCount; index += 1) {
      cells[args.cellIndices![index]!] = 1
    }
    return {
      has: (cellIndex: number): boolean => (cells[cellIndex] ?? 0) !== 0,
      add: (cellIndex: number): void => {
        if (cellIndex >= cells.length) {
          const next = new Uint8Array(cellIndex + 1)
          next.set(cells)
          cells = next
        }
        cells[cellIndex] = 1
      },
      delete: (cellIndex: number): void => {
        cells[cellIndex] = 0
      },
    }
  }

  const cells = new Set<number>()
  for (let index = 0; index < cellCount; index += 1) {
    cells.add(args.cellIndices![index]!)
  }
  return {
    has: (cellIndex: number): boolean => cells.has(cellIndex),
    add: (cellIndex: number): void => {
      cells.add(cellIndex)
    },
    delete: (cellIndex: number): void => {
      cells.delete(cellIndex)
    },
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
