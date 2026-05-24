import type { U32 } from '../runtime-state.js'

export interface RecalcValueChangeCollector {
  readonly note: (cellIndex: number) => void
  readonly toChangedSet: () => U32
}

export function createRecalcValueChangeCollector(initialCellStoreSize: number): RecalcValueChangeCollector {
  const seen = new Uint8Array(Math.max(0, initialCellStoreSize + 1))
  const overflowSeen = new Set<number>()
  const changed: number[] = []

  const note = (cellIndex: number): void => {
    if (!Number.isSafeInteger(cellIndex) || cellIndex < 0) {
      return
    }
    if (cellIndex < seen.length) {
      if (seen[cellIndex] !== 0) {
        return
      }
      seen[cellIndex] = 1
    } else {
      if (overflowSeen.has(cellIndex)) {
        return
      }
      overflowSeen.add(cellIndex)
    }
    changed.push(cellIndex)
  }

  return {
    note,
    toChangedSet: () => (changed.length === 0 ? new Uint32Array(0) : Uint32Array.from(changed)),
  }
}
