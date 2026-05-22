import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine data validations', () => {
  it('roundtrips data validation metadata through snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-roundtrip' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
      showDropdown: true,
      errorStyle: 'stop',
      errorTitle: 'Status required',
      errorMessage: 'Pick Draft or Final.',
    })

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets.find((sheet) => sheet.name === 'Sheet1')?.metadata?.validations).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
        errorStyle: 'stop',
        errorTitle: 'Status required',
        errorMessage: 'Pick Draft or Final.',
      },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'validation-roundtrip-restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
        errorStyle: 'stop',
        errorTitle: 'Status required',
        errorMessage: 'Pick Draft or Final.',
      },
    ])
  })

  it('rewrites and restores data validation ranges across structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-structural' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
    })

    engine.insertRows('Sheet1', 1, 1)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B3',
          endAddress: 'B5',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])

    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A3',
          endAddress: 'A5',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])

    expect(engine.undo()).toBe(true)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B3',
          endAddress: 'B5',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])

    expect(engine.undo()).toBe(true)
    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
      },
    ])
  })

  it('clips bottom-bounded data validation ranges during structural inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-structural-bottom-clip' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C159',
        endAddress: 'C1048576',
      },
      rule: {
        kind: 'textLength',
        operator: 'lessThanOrEqual',
        values: [100],
      },
      allowBlank: true,
    })

    engine.insertRows('Sheet1', 0, 1)

    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'C160',
          endAddress: 'C1048576',
        },
        rule: {
          kind: 'textLength',
          operator: 'lessThanOrEqual',
          values: [100],
        },
        allowBlank: true,
      },
    ])
  })

  it('rejects local list and scalar validation violations before committing value batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-local-enforcement' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'D2', 'Draft')
    engine.setCellValue('Sheet1', 'D3', 'Final')
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B2',
      },
      rule: {
        kind: 'list',
        source: {
          kind: 'range-ref',
          sheetName: 'Sheet1',
          startAddress: 'D2',
          endAddress: 'D3',
        },
      },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C2',
        endAddress: 'C2',
      },
      rule: {
        kind: 'decimal',
        operator: 'between',
        values: [0, 1],
      },
      allowBlank: false,
    })

    engine.setRangeValues(
      {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C2',
      },
      [['Draft', 0.25]],
    )

    expect(() =>
      engine.setRangeValues(
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C2',
        },
        [['Bogus', 0.75]],
      ),
    ).toThrow(/Excel data validation/)
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.String, value: 'Draft', stringId: expect.any(Number) })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0.25 })

    expect(() =>
      engine.setRangeValues(
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C2',
        },
        [['Final', 2]],
      ),
    ).toThrow(/Excel data validation/)
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.String, value: 'Draft', stringId: expect.any(Number) })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0.25 })

    engine.setRangeValues(
      {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C2',
      },
      [['Final', 0.75]],
    )
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.String, value: 'Final', stringId: expect.any(Number) })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0.75 })
  })

  it('enforces validations on direct coordinate and existing-cell mutation fast paths', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-local-fast-paths', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.setCellValue('Sheet1', 'B2', 'Draft')
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B2',
      },
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
    })
    const b2Index = engine.workbook.getCellIndex('Sheet1', 'B2')!

    expect(() =>
      engine.tryApplyExistingLiteralCellMutationAt({
        sheetId,
        row: 1,
        col: 1,
        cellIndex: b2Index,
        value: 'Bogus',
      }),
    ).toThrow(/Excel data validation/)
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.String, value: 'Draft', stringId: expect.any(Number) })

    engine.setCellValue('Sheet1', 'C2', 5)
    engine.setDataValidation({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C2',
        endAddress: 'C2',
      },
      rule: {
        kind: 'whole',
        operator: 'between',
        values: [1, 10],
      },
      allowBlank: false,
    })
    const c2Index = engine.workbook.getCellIndex('Sheet1', 'C2')!

    expect(() =>
      engine.tryApplyExistingNumericCellMutationAt({
        sheetId,
        row: 1,
        col: 2,
        cellIndex: c2Index,
        value: 11,
      }),
    ).toThrow(/Excel data validation/)
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 5 })

    expect(() => engine.setCellValueAt(sheetId, 1, 2, 0)).toThrow(/Excel data validation/)
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.setCellValueAt(sheetId, 1, 2, 7)).toEqual({ tag: ValueTag.Number, value: 7 })
  })

  it('loads existing invalid validated values from snapshots but rejects later invalid local writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-import-existing-invalid' })
    await engine.ready()
    engine.importSnapshot({
      version: 1,
      workbook: { name: 'validation-import-existing-invalid' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          metadata: {
            validations: [
              {
                range: {
                  sheetName: 'Sheet1',
                  startAddress: 'B2',
                  endAddress: 'B2',
                },
                rule: {
                  kind: 'list',
                  values: ['Draft', 'Final'],
                },
                allowBlank: false,
              },
            ],
          },
          cells: [{ address: 'B2', value: 'Legacy' }],
        },
      ],
    })

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.String, value: 'Legacy', stringId: expect.any(Number) })
    expect(() => engine.setCellValue('Sheet1', 'B2', 'Bogus')).toThrow(/Excel data validation/)
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.String, value: 'Legacy', stringId: expect.any(Number) })

    expect(engine.setCellValue('Sheet1', 'B2', 'Final')).toEqual({
      tag: ValueTag.String,
      value: 'Final',
      stringId: expect.any(Number),
    })
  })
})
