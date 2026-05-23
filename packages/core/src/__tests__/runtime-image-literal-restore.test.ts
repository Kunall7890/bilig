import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'
import { CellFlags, CellStore } from '../cell-store.js'
import { restoreFreshRuntimeLiteralCell } from '../snapshot/runtime-image-literal-restore.js'
import { StringPool } from '../string-pool.js'

describe('restoreFreshRuntimeLiteralCell', () => {
  it('restores fresh literal values without cell-wise notifications', () => {
    const cellStore = new CellStore()
    const strings = new StringPool()
    const onSetValue = vi.fn()
    cellStore.onSetValue = onSetValue
    const cache = new Map<string, number>()
    const emptyIndex = cellStore.allocateReserved(1, 0, 0)
    const blankIndex = cellStore.allocateReserved(1, 0, 1)
    const numberIndex = cellStore.allocateReserved(1, 0, 2)
    const booleanIndex = cellStore.allocateReserved(1, 0, 3)
    const stringIndex = cellStore.allocateReserved(1, 0, 4)

    restoreFreshRuntimeLiteralCell(cellStore, strings, emptyIndex, undefined, cache)
    restoreFreshRuntimeLiteralCell(cellStore, strings, blankIndex, null, cache)
    restoreFreshRuntimeLiteralCell(cellStore, strings, numberIndex, 42, cache)
    restoreFreshRuntimeLiteralCell(cellStore, strings, booleanIndex, true, cache)
    restoreFreshRuntimeLiteralCell(cellStore, strings, stringIndex, 'North', cache)

    expect(onSetValue).not.toHaveBeenCalled()
    expect(cellStore.tags[emptyIndex]).toBe(ValueTag.Empty)
    expect(cellStore.flags[emptyIndex] & CellFlags.AuthoredBlank).toBe(0)
    expect(cellStore.tags[blankIndex]).toBe(ValueTag.Empty)
    expect(cellStore.flags[blankIndex] & CellFlags.AuthoredBlank).not.toBe(0)
    expect(cellStore.tags[numberIndex]).toBe(ValueTag.Number)
    expect(cellStore.numbers[numberIndex]).toBe(42)
    expect(cellStore.tags[booleanIndex]).toBe(ValueTag.Boolean)
    expect(cellStore.numbers[booleanIndex]).toBe(1)
    expect(cellStore.tags[stringIndex]).toBe(ValueTag.String)
    expect(cellStore.stringIds[stringIndex]).toBeGreaterThan(0)
    for (const index of [emptyIndex, blankIndex, numberIndex, booleanIndex, stringIndex]) {
      expect(cellStore.errors[index]).toBe(ErrorCode.None)
      expect(cellStore.versions[index]).toBe(1)
    }
  })

  it('reuses string ids through the restore cache', () => {
    const cellStore = new CellStore()
    const strings = new StringPool()
    const intern = vi.spyOn(strings, 'intern')
    const cache = new Map<string, number>()
    const first = cellStore.allocateReserved(1, 0, 0)
    const second = cellStore.allocateReserved(1, 0, 1)

    restoreFreshRuntimeLiteralCell(cellStore, strings, first, 'North', cache)
    restoreFreshRuntimeLiteralCell(cellStore, strings, second, 'North', cache)

    expect(intern).toHaveBeenCalledOnce()
    expect(cellStore.stringIds[first]).toBe(cellStore.stringIds[second])
  })
})
