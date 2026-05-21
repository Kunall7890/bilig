import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine facade coverage', () => {
  it('keeps direct sheet-id mutation helpers and history state predictable', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'facade-coverage' })
    await engine.ready()
    const sheetId = engine.createSheetForInitialization('Sheet1')

    expect(engine.canUndo()).toBe(false)
    expect(engine.canRedo()).toBe(false)
    expect(engine.setCellValueAt(sheetId, 0, 0, 7)).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.setCellFormulaAt(sheetId, 0, 1, 'A1*2')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.setCellFormulaAt(sheetId, 0, 1, 'A1*2')).toEqual({ tag: ValueTag.Number, value: 14 })

    engine.clearCellAt(sheetId, 0, 0)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.canUndo()).toBe(true)

    engine.clearHistory()
    expect(engine.canUndo()).toBe(false)
    expect(engine.canRedo()).toBe(false)
  })

  it('rejects direct sheet-id helpers for unknown sheets', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'facade-unknown-sheet' })
    await engine.ready()

    expect(() => engine.setCellValueAt(999, 0, 0, 1)).toThrow('Unknown sheet id: 999')
    expect(() => engine.setCellFormulaAt(999, 0, 0, '1+1')).toThrow('Unknown sheet id: 999')
  })

  it('routes selection, subscriptions, ranges, snapshots, and metrics through the facade', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'facade-routes' })
    await engine.ready()
    engine.createSheetForInitialization('Sheet1')

    const eventListener = vi.fn()
    const selectionListener = vi.fn()
    const batches: unknown[] = []
    const unsubscribeEvents = engine.subscribe(eventListener)
    const unsubscribeSelection = engine.subscribeSelection(selectionListener)
    const unsubscribeBatches = engine.subscribeBatches((batch) => batches.push(batch))

    engine.setSelection('Sheet1', 'B2', {
      anchorAddress: 'A1',
      range: { startAddress: 'A1', endAddress: 'B2' },
      editMode: 'cell',
    })
    expect(engine.getSelectionState()).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      anchorAddress: 'A1',
      range: { startAddress: 'A1', endAddress: 'B2' },
      editMode: 'cell',
    })
    expect(selectionListener).toHaveBeenCalledTimes(1)

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [
      [1, 2],
      [3, 4],
    ])
    expect(engine.getRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' })).toEqual([
      [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ],
      [
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ],
    ])

    engine.setRangeFormulas({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' }, [['A1+B1'], ['A2+B2']])
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getDependencies('Sheet1', 'C2').directPrecedents).toContain('Sheet1!A2')
    expect(engine.getDependents('Sheet1', 'A2').directDependents).toContain('Sheet1!C2')
    expect(engine.explainCell('Sheet1', 'C2').formula).toBe('A2+B2')

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'B4' },
    )
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'B4' },
      { sheetName: 'Sheet1', startAddress: 'A5', endAddress: 'B5' },
    )
    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'A5', endAddress: 'B5' },
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E5' },
    )
    engine.pasteRange(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E5' },
      { sheetName: 'Sheet1', startAddress: 'D6', endAddress: 'E6' },
    )
    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E5' })
    expect(engine.getCellValue('Sheet1', 'D5')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'D6')).toEqual({ tag: ValueTag.Number, value: 1 })

    const csv = engine.exportSheetCsv('Sheet1')
    expect(csv).toContain('1,2,=A1+B1')
    engine.importSheetCsv('Sheet1', 'sku,amount\n001,4')
    expect(engine.getCellValue('Sheet1', 'A2')).toMatchObject({ tag: ValueTag.String, value: '001' })

    const snapshot = engine.exportSnapshot()
    engine.importSnapshot(snapshot)
    expect(engine.exportSnapshot().sheets[0]?.name).toBe('Sheet1')

    const replicaSnapshot = engine.exportReplicaSnapshot()
    engine.importReplicaSnapshot(replicaSnapshot)

    const captured = engine.captureUndoOps(() => engine.setCellValue('Sheet1', 'B2', 9))
    expect(captured.result).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(captured.undoOps?.length).toBeGreaterThan(0)

    const undoOps = engine.applyOps([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B3', value: true }], {
      captureUndo: true,
    })
    expect(undoOps?.length).toBeGreaterThan(0)
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Boolean, value: true })

    expect(engine.getLastMetrics()).toBeDefined()
    expect(engine.getPerformanceCounters()).toBeDefined()
    engine.resetPerformanceCounters()
    engine.setUseColumnIndexEnabled(false)
    engine.setUseColumnIndexEnabled(true)
    expect(engine.getSyncState()).toBe('local-only')

    expect(eventListener).toHaveBeenCalled()
    expect(batches.length).toBeGreaterThan(0)
    unsubscribeBatches()
    unsubscribeSelection()
    unsubscribeEvents()
  })

  it('covers sheet lifecycle no-ops and sync-disabled rejection', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'facade-sheet-lifecycle' })
    await engine.ready()
    engine.createSheetForInitialization('Sheet1')
    engine.createSheet('Sheet2')

    engine.renameSheet('Sheet1', '   ')
    engine.renameSheet('Sheet1', 'Sheet2')
    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Sheet2'])

    engine.renameSheet('Sheet1', 'Data')
    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Sheet2'])
    engine.deleteSheet('Sheet2')
    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toEqual(['Data'])

    const syncDisabledEngine = new SpreadsheetEngine({ workbookName: 'facade-sync-disabled', trackReplicaVersions: false })
    await syncDisabledEngine.ready()
    const client = {
      connect: () => ({
        send: () => {},
        disconnect: () => {},
      }),
    }
    await expect(syncDisabledEngine.connectSyncClient(client)).rejects.toThrow('Sync is unavailable')
    await expect(syncDisabledEngine.disconnectSyncClient()).resolves.toBeUndefined()
  })
})
