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

  it('enforces checkbox, any, and scalar validation rule variants', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-scalar-variants' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      rule: { kind: 'any' },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      rule: { kind: 'checkbox', checkedValue: 'yes', uncheckedValue: 'no' },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C1' },
      rule: { kind: 'whole', operator: 'greaterThan', values: [10] },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C2' },
      rule: { kind: 'decimal', operator: 'notBetween', values: [1, 2] },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'C3' },
      rule: { kind: 'textLength', operator: 'lessThanOrEqual', values: [3] },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'C4' },
      rule: { kind: 'date', operator: 'lessThan', values: ['2024-02-01'] },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'C5' },
      rule: { kind: 'time', operator: 'lessThanOrEqual', values: ['12:30'] },
      allowBlank: false,
    })

    expect(engine.setCellValue('Sheet1', 'A1', 'anything')).toEqual({
      tag: ValueTag.String,
      value: 'anything',
      stringId: expect.any(Number),
    })
    expect(engine.setCellValue('Sheet1', 'B1', 'yes')).toEqual({ tag: ValueTag.String, value: 'yes', stringId: expect.any(Number) })
    expect(engine.setCellValue('Sheet1', 'B1', 'no')).toEqual({ tag: ValueTag.String, value: 'no', stringId: expect.any(Number) })
    expect(() => engine.setCellValue('Sheet1', 'B1', true)).toThrow(/Excel data validation rejects Sheet1!B1 value true/)
    expect(engine.setCellValue('Sheet1', 'C1', 11)).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(() => engine.setCellValue('Sheet1', 'C1', 10.5)).toThrow(/Excel data validation/)
    expect(() => engine.setCellValue('Sheet1', 'C1', 'not-a-number')).toThrow(/Excel data validation/)
    expect(engine.setCellValue('Sheet1', 'C2', 3)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(() => engine.setCellValue('Sheet1', 'C2', 1.5)).toThrow(/Excel data validation/)
    expect(engine.setCellValue('Sheet1', 'C3', 'abc')).toEqual({ tag: ValueTag.String, value: 'abc', stringId: expect.any(Number) })
    expect(() => engine.setCellValue('Sheet1', 'C3', 'abcd')).toThrow(/Excel data validation/)
    expect(engine.setCellValue('Sheet1', 'C4', '2024-01-01')).toEqual({
      tag: ValueTag.String,
      value: '2024-01-01',
      stringId: expect.any(Number),
    })
    expect(() => engine.setCellValue('Sheet1', 'C4', 'not-a-date')).toThrow(/Excel data validation/)
    expect(engine.setCellValue('Sheet1', 'C5', '12:00')).toEqual({ tag: ValueTag.String, value: '12:00', stringId: expect.any(Number) })
    expect(() => engine.setCellValue('Sheet1', 'C5', '25:00')).toThrow(/Excel data validation/)
  })

  it('enforces list validations from cell refs, named ranges, and structured refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-list-sources' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'D1', 'CellChoice')
    engine.setCellValue('Sheet1', 'E1', 'RangeA')
    engine.setCellValue('Sheet1', 'E2', 'RangeB')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'G1', endAddress: 'G4' }, [['Status'], ['Open'], ['Closed'], ['Total']])
    engine.setTable({
      name: 'StatusTable',
      sheetName: 'Sheet1',
      startAddress: 'G1',
      endAddress: 'G4',
      columnNames: ['Status'],
      headerRow: true,
      totalsRow: true,
    })
    engine.setDefinedName('ChoiceRange', { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'E1', endAddress: 'E2' })
    engine.setDefinedName('ChoiceScalar', { kind: 'scalar', value: 'Solo' })
    engine.setDefinedName('ChoiceCell', { kind: 'cell-ref', sheetName: 'Sheet1', address: 'D1' })
    engine.setDefinedName('ChoiceStructured', { kind: 'structured-ref', tableName: 'StatusTable', columnName: 'Status' })
    engine.setDefinedName('ChoiceFormula', { kind: 'formula', formula: '=Sheet1!D1' })

    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      rule: { kind: 'list', source: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'D1' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'A2' },
      rule: { kind: 'list', source: { kind: 'named-range', name: 'ChoiceRange' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A3' },
      rule: { kind: 'list', source: { kind: 'named-range', name: 'ChoiceScalar' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'A4' },
      rule: { kind: 'list', source: { kind: 'named-range', name: 'ChoiceCell' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A5', endAddress: 'A5' },
      rule: { kind: 'list', source: { kind: 'structured-ref', tableName: 'StatusTable', columnName: 'Status' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A6', endAddress: 'A6' },
      rule: { kind: 'list', source: { kind: 'named-range', name: 'ChoiceStructured' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A7', endAddress: 'A7' },
      rule: { kind: 'list', source: { kind: 'named-range', name: 'ChoiceFormula' } },
      allowBlank: false,
    })

    expect(engine.setCellValue('Sheet1', 'A1', 'cellchoice')).toEqual({
      tag: ValueTag.String,
      value: 'cellchoice',
      stringId: expect.any(Number),
    })
    expect(engine.setCellValue('Sheet1', 'A2', 'RangeB')).toEqual({ tag: ValueTag.String, value: 'RangeB', stringId: expect.any(Number) })
    expect(engine.setCellValue('Sheet1', 'A3', 'Solo')).toEqual({ tag: ValueTag.String, value: 'Solo', stringId: expect.any(Number) })
    expect(engine.setCellValue('Sheet1', 'A4', 'CellChoice')).toEqual({
      tag: ValueTag.String,
      value: 'CellChoice',
      stringId: expect.any(Number),
    })
    expect(engine.setCellValue('Sheet1', 'A5', 'Closed')).toEqual({ tag: ValueTag.String, value: 'Closed', stringId: expect.any(Number) })
    expect(engine.setCellValue('Sheet1', 'A6', 'Open')).toEqual({ tag: ValueTag.String, value: 'Open', stringId: expect.any(Number) })
    expect(() => engine.setCellValue('Sheet1', 'A5', 'Total')).toThrow(/Excel data validation/)
    expect(() => engine.setCellValue('Sheet1', 'A7', 'CellChoice')).toThrow(/formula-backed/)
  })

  it('fails closed for unresolved list validation sources', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'validation-list-source-errors' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      rule: { kind: 'list', source: { kind: 'range-ref', sheetName: 'Missing', startAddress: 'A1', endAddress: 'A1' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'A2' },
      rule: { kind: 'list', source: { kind: 'structured-ref', tableName: 'MissingTable', columnName: 'Status' } },
      allowBlank: false,
    })
    engine.setTable({
      name: 'StatusTable',
      sheetName: 'Sheet1',
      startAddress: 'D1',
      endAddress: 'D2',
      columnNames: ['Status'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A3' },
      rule: { kind: 'list', source: { kind: 'structured-ref', tableName: 'StatusTable', columnName: 'MissingColumn' } },
      allowBlank: false,
    })
    engine.setDataValidation({
      range: { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'A4' },
      rule: { kind: 'list', source: { kind: 'named-range', name: 'MissingName' } },
      allowBlank: false,
    })

    expect(() => engine.setCellValue('Sheet1', 'A1', 'x')).toThrow(/sheet not found/)
    expect(() => engine.setCellValue('Sheet1', 'A2', 'x')).toThrow(/table not found/)
    expect(() => engine.setCellValue('Sheet1', 'A3', 'x')).toThrow(/table column not found/)
    expect(() => engine.setCellValue('Sheet1', 'A4', 'x')).toThrow(/named range not found/)
  })
})
