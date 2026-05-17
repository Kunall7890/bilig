import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import {
  createClearTombstoneSnapshot,
  isClearCellSnapshot,
  isResetEmptySnapshot,
  prepareIncomingSnapshot,
  shouldKeepCurrentSnapshot,
} from '../projected-viewport-cell-snapshot-policy.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'

function textSnapshot(overrides: Partial<CellSnapshot> = {}): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address: 'B2',
    input: 'value',
    value: { tag: ValueTag.String, value: 'value', stringId: 1 },
    flags: 0,
    version: 7,
    ...overrides,
  }
}

function clearSnapshot(overrides: Partial<CellSnapshot> = {}): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address: 'B2',
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 7,
    ...overrides,
  }
}

describe('projected viewport cell snapshot policy', () => {
  it('keeps confirmed clear tombstones ahead of same-version stale content', () => {
    expect(shouldKeepCurrentSnapshot(clearSnapshot(), textSnapshot())).toBe(true)
  })

  it('releases optimistic clear protection when authoritative empty confirmation arrives', () => {
    const current = clearSnapshot({ flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG, version: 8 })
    const incoming = clearSnapshot({ version: 9 })

    expect(prepareIncomingSnapshot(current, incoming, { releaseConfirmedOptimisticClear: true })).toEqual(clearSnapshot({ version: 9 }))
  })

  it('recognizes reset-empty and clear-tombstone snapshots consistently', () => {
    const reset = clearSnapshot({ version: 0 })

    expect(isResetEmptySnapshot(reset)).toBe(true)
    expect(isClearCellSnapshot(reset)).toBe(true)
    expect(createClearTombstoneSnapshot(clearSnapshot({ flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG, styleId: 'stale' }))).toEqual(clearSnapshot())
  })
})
