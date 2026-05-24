import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

async function engineFromSnapshot(snapshot: WorkbookSnapshot): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: `structural-insert-history-${Math.random()}`,
  })
  await engine.ready()
  engine.importSnapshot(structuredClone(snapshot))
  return engine
}

describe('engine no-op history', () => {
  const blankSheetSnapshot: WorkbookSnapshot = {
    version: 1,
    workbook: { name: 'History' },
    sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [] }],
  }

  it('does not record row inserts that have no semantic impact', async () => {
    const engine = await engineFromSnapshot(blankSheetSnapshot)

    engine.insertRows('Sheet1', 0, 1)

    expect(engine.exportSnapshot()).toEqual(blankSheetSnapshot)
    expect(engine.undo()).toBe(false)
  })

  it('does not let a no-op column insert intercept undo for the previous semantic mutation', async () => {
    const engine = await engineFromSnapshot({
      ...blankSheetSnapshot,
      sheets: [
        ...blankSheetSnapshot.sheets,
        { id: 2, name: 'Summary', order: 1, cells: [] },
      ],
    })

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, [[0]])
    engine.insertColumns('Sheet1', 3, 1)

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual({
      ...blankSheetSnapshot,
      sheets: [
        ...blankSheetSnapshot.sheets,
        { id: 2, name: 'Summary', order: 1, cells: [] },
      ],
    })
    expect(engine.undo()).toBe(false)
  })

  it('does not let repeated filter sets intercept undo for the previous semantic filter mutation', async () => {
    const engine = await engineFromSnapshot(blankSheetSnapshot)
    const range = { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C5' }

    engine.setFilter('Sheet1', range)
    const filtered = engine.exportSnapshot()
    engine.setFilter('Sheet1', range)

    expect(engine.exportSnapshot()).toEqual(filtered)
    expect(engine.undo()).toBe(true)
    expect(engine.getFilters('Sheet1')).toEqual([])
    expect(engine.undo()).toBe(false)
  })
})
