import { describe, expect, it } from 'vitest'
import { createBatch, createReplicaState } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine protections', () => {
  it('roundtrips sheet and range protections through snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-roundtrip' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setSheetProtection({ sheetName: 'Sheet1', hideFormulas: true })
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
      hideFormulas: true,
    })

    const snapshot = engine.exportSnapshot()
    const metadata = snapshot.sheets.find((sheet) => sheet.name === 'Sheet1')?.metadata
    expect(metadata?.sheetProtection).toEqual({
      sheetName: 'Sheet1',
      hideFormulas: true,
    })
    expect(metadata?.protectedRanges).toEqual([
      {
        id: 'protect-a1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        hideFormulas: true,
      },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'protection-roundtrip-restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)
    expect(restored.getSheetProtection('Sheet1')).toEqual({
      sheetName: 'Sheet1',
      hideFormulas: true,
    })
    expect(restored.getRangeProtections('Sheet1')).toEqual([
      {
        id: 'protect-a1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        hideFormulas: true,
      },
    ])
  })

  it('blocks writes to protected sheets and ranges while allowing explicit unprotect', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-enforcement' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
    })

    expect(() => engine.setCellValue('Sheet1', 'A1', 7)).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.setCellValue('Sheet1', 'C3', 7)).not.toThrow()

    engine.setSheetProtection({ sheetName: 'Sheet1', hideFormulas: true })
    expect(() => engine.setCellValue('Sheet1', 'D4', 9)).toThrow(/Workbook protection blocks this change/)

    expect(engine.clearSheetProtection('Sheet1')).toBe(true)
    expect(engine.deleteRangeProtection('protect-a1')).toBe(true)
    expect(() => engine.setCellValue('Sheet1', 'A1', 7)).not.toThrow()
  })

  it('reports missing range protections after deletion through the public API', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-delete-missing' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
      hideFormulas: true,
    })

    expect(engine.getRangeProtection('protect-a1')).toMatchObject({ id: 'protect-a1' })
    expect(engine.deleteRangeProtection('protect-a1')).toBe(true)
    expect(engine.getRangeProtection('protect-a1')).toBeUndefined()
    expect(engine.getRangeProtections('Sheet1')).toEqual([])
    expect(engine.deleteRangeProtection('protect-a1')).toBe(false)
  })

  it('blocks sheet topology mutations when workbook structure is protected', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'workbook-structure-protection' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Report')
    engine.workbook.setWorkbookProtection({
      lockStructure: true,
      xmlAttributes: [{ name: 'lockStructure', value: '1' }],
    })

    expect(() => engine.createSheet('Added')).toThrow(/workbook structure is protected/)
    expect(() => engine.moveSheet('Report', 0)).toThrow(/workbook structure is protected/)
    expect(() => engine.renameSheet('Data', 'Source')).toThrow(/workbook structure is protected/)
    expect(() => engine.deleteSheet('Data')).toThrow(/workbook structure is protected/)
    expect(() => engine.renderCommit([{ kind: 'upsertSheet', name: 'Committed', order: 2 }])).toThrow(
      /Failed to execute render commit transaction/,
    )
    expect(() =>
      engine.applyRemoteBatch(createBatch(createReplicaState('remote-lock'), [{ kind: 'upsertSheet', name: 'Remote', order: 2 }])),
    ).toThrow(/Failed to apply remote batch/)
    expect(engine.renameSheetMetadataOnly('Data', 'Source')).toBe(false)
    expect(engine.renameSheetMetadataOnlyById(engine.workbook.getSheet('Data')!.id, 'Source')).toBe(false)

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Report'])
    expect(engine.exportSnapshot().workbook.metadata?.workbookProtection).toEqual({
      lockStructure: true,
      xmlAttributes: [{ name: 'lockStructure', value: '1' }],
    })
  })
})
