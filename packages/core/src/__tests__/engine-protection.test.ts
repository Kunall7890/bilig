import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
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

  it('allows content edits to unlocked cells on protected sheets while blocking locked cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-unlocked-cell-edits' })
    await engine.ready()
    engine.importSnapshot(protectedSheetWithUnlockedInputSnapshot())

    expect(() => engine.setCellValue('Input', 'B2', 25)).not.toThrow()
    expect(() => engine.setCellFormula('Input', 'B2', '40')).not.toThrow()
    expect(() => engine.clearCell('Input', 'B2')).not.toThrow()
    expect(() => engine.setCellValue('Input', 'A1', 1)).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.setCellValue('Input', 'C2', 50)).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.setCellFormat('Input', 'B2', '0.00')).toThrow(/Workbook protection blocks this change/)
  })

  it('blocks coordinate and direct fast-path content writes to protected cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-coordinate-cell-edits' })
    await engine.ready()
    engine.importSnapshot(protectedSheetWithUnlockedInputSnapshot())
    const sheetId = engine.workbook.getSheet('Input')!.id

    expect(() => engine.setCellValueAt(sheetId, 0, 0, 25)).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.setCellFormulaAt(sheetId, 0, 0, '40')).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.clearCellAt(sheetId, 0, 0)).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.applyCellMutationsAt([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 30 } }], 1)).toThrow(
      /Workbook protection blocks this change/,
    )
    expect(() =>
      engine.applyCellMutationsAtWithOptions([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 32 } }]),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() =>
      engine.applyCellMutationsAtWithOptions([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 35 } }], {
        captureUndo: false,
        source: 'local',
      }),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.applyOps([{ kind: 'setCellValue', sheetName: 'Input', address: 'A1', value: 40 }])).toThrow(
      /Workbook protection blocks this change/,
    )
    expect(() => engine.applyOps([{ kind: 'setCellValue', sheetName: 'Input', address: 'A1', value: 45 }], { trusted: true })).toThrow(
      /Workbook protection blocks this change/,
    )
    expect(() => engine.setCellValueAt(sheetId, 1, 1, 25)).not.toThrow()
    expect(() =>
      engine.applyCellMutationsAtWithOptions([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 50 } }], {
        captureUndo: false,
        source: 'restore',
      }),
    ).not.toThrow()
    expect(engine.getCellValue('Input', 'A1')).toMatchObject({ value: 50 })
  })

  it('keeps explicit protected ranges locked even when sheet cell style is unlocked', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'protection-unlocked-cell-explicit-range' })
    await engine.ready()
    const snapshot = protectedSheetWithUnlockedInputSnapshot()
    snapshot.sheets[0].metadata!.protectedRanges = [
      { id: 'protect-input', range: { sheetName: 'Input', startAddress: 'B2', endAddress: 'B2' } },
    ]
    engine.importSnapshot(snapshot)

    expect(() => engine.setCellValue('Input', 'B2', 25)).toThrow(/Workbook protection blocks this change/)
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

function protectedSheetWithUnlockedInputSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Protected input template',
      metadata: {
        styles: [
          {
            id: 'unlocked-input',
            protection: { locked: false },
          },
          {
            id: 'locked-formula',
            protection: { locked: true, hidden: true },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Input',
        order: 0,
        metadata: {
          sheetProtection: { sheetName: 'Input' },
          styleRanges: [
            { range: { sheetName: 'Input', startAddress: 'B2', endAddress: 'B2' }, styleId: 'unlocked-input' },
            { range: { sheetName: 'Input', startAddress: 'C2', endAddress: 'C2' }, styleId: 'locked-formula' },
          ],
        },
        cells: [
          { address: 'A1', value: 5 },
          { address: 'B2', value: 10 },
          { address: 'C2', formula: 'B2*2' },
          { address: 'D1', formula: 'A1*2' },
        ],
      },
    ],
  }
}
