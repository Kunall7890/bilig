import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type WorkbookConditionalFormatSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import {
  applyActionAndCaptureResult,
  createEngineSeedSnapshot,
  exportReplaySnapshot,
  normalizeSnapshotForSemanticComparison,
  type CoreAction,
} from './engine-fuzz-helpers.js'

describe('engine fuzz regressions', () => {
  it('falls back when range values overwrite direct scalar delta targets', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'range-overwrites-direct-scalar-delta-target-regression',
      replicaId: 'range-overwrites-direct-scalar-delta-target-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A2', 'B2-A1')
    engine.insertColumns('Sheet1', 0, 2)
    engine.setCellFormula('Sheet1', 'C1', 'B2+A1')

    expect(() =>
      engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }, [
        [0, null],
        [0, 'north'],
      ]),
    ).not.toThrow()

    expect(engine.exportSnapshot().sheets[0]?.cells).toEqual([
      { address: 'C1', formula: 'B2+A1' },
      { address: 'B2', value: 0 },
      { address: 'C2', value: null },
      { address: 'B3', value: 0 },
      { address: 'C3', value: 'north' },
    ])
  })

  it('keeps coalesced style history aligned before structural row inserts', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'style-history-row-insert-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'style-history-row-insert-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const actions: CoreAction[] = [
      {
        kind: 'style',
        range: { sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'C4' },
        patch: { fill: { backgroundColor: '#dbeafe' } },
      },
      { kind: 'insertRows', start: 3, count: 1 },
    ]

    engine.setRangeStyle(actions[0].range, actions[0].patch)
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })
    expect(engine.undo()).toBe(true)
    engine.insertRows('Sheet1', 3, 1)

    const expected = await exportReplaySnapshot(seedSnapshot, actions)
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(normalizeSnapshotForSemanticComparison(expected))
  })

  it('does not mutate source cells when moveRange is blocked by protection', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'move-protection-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C5' }, [
      ['Id', 'Status', 'Amount'],
      [1, 'Draft', 10],
      [2, 'Final', 20],
      [3, 'Review', 30],
      [4, 'Final', 40],
    ])
    engine.setRangeProtection({
      id: 'protect-a1-c4',
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C4' },
      hideFormulas: true,
    })

    const before = engine.exportSnapshot()
    expect(() =>
      engine.moveRange(
        { sheetName: 'Sheet1', startAddress: 'A5', endAddress: 'A6' },
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
      ),
    ).toThrow(/Failed to execute local transaction/)
    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('restores pivot materialization dimensions after undoing source clears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'pivot-undo-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C5' }, [
      ['Region', 'Quarter', 'Sales'],
      ['East', 'Q1', 10],
      ['West', 'Q1', 6],
      ['East', 'Q2', 12],
      ['West', 'Q2', 9],
    ])
    engine.setTable({
      name: 'QuarterlySales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C5',
      columnNames: ['Region', 'Quarter', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setPivotTable('Pivot', 'B2', {
      name: 'QuarterlyPivot',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C5' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getPivotTable('Pivot', 'B2')?.rows).toBe(3)
    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' })
    expect(engine.getPivotTable('Pivot', 'B2')?.rows).toBe(1)
    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')?.rows).toBe(3)
  })

  it('keeps pivot dimensions aligned after copy, row delete, and undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('pivot-analytics', 'pivot-copy-delete-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'pivot-copy-delete-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(seedSnapshot)

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
    )
    const pivotAfterCopy = engine.getPivotTable('Pivot', 'B2')
    expect(pivotAfterCopy).toMatchObject({ rows: 1, cols: 1 })

    engine.deleteRows('Sheet1', 0, 1)
    expect(engine.undo()).toBe(true)

    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      rows: pivotAfterCopy?.rows,
      cols: pivotAfterCopy?.cols,
    })
  })

  it('keeps pivot source clears aligned across move undo and redo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('pivot-analytics', 'pivot-clear-move-redo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'pivot-clear-move-redo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertRows('Sheet1', 0, 1)
    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'E3', endAddress: 'E4' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
    )
    const pivotAfterMove = engine.getPivotTable('Pivot', 'B2')

    expect(pivotAfterMove).toMatchObject({ rows: 1, cols: 1 })
    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({ rows: 3, cols: 2 })
    expect(engine.redo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      rows: pivotAfterMove?.rows,
      cols: pivotAfterMove?.cols,
    })
  })

  it('treats structural deletes on blank sheets as history no-ops', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'blank-structural-noop-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    const snapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'blank-structural-noop' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    const before = engine.exportSnapshot()
    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.exportSnapshot()).toEqual(before)
    expect(engine.undo()).toBe(false)
  })

  it('restores sparse style metadata after undoing coalesced structural deletes', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'structural-style-undo-regression')
    const engine = new SpreadsheetEngine({ workbookName: seedSnapshot.workbook.name, replicaId: 'structural-style-undo-regression' })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'C5' }, { fill: { backgroundColor: '#dbeafe' } })
    engine.deleteRows('Sheet1', 3, 1)
    engine.deleteColumns('Sheet1', 0, 1)

    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(seedSnapshot)
  })

  it('restores sparse style range shapes after undoing a partial clear', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'sparse-style-clear-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'sparse-style-clear-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'D4' }, { fill: { backgroundColor: '#dbeafe' } })
    const styledSnapshot = engine.exportSnapshot()

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'D3', endAddress: 'D4' })

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(styledSnapshot)
  })

  it('restores explicit formats on formula cells deleted by structural undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'sparse-formula-format-delete-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'sparse-formula-format-delete-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const moveAction = {
      kind: 'move',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      target: { sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'C5' },
    } satisfies CoreAction
    const formulaAction = {
      kind: 'formula',
      address: 'C5',
      formula: 'A1+A1',
    } satisfies CoreAction

    engine.moveRange(moveAction.source, moveAction.target)
    engine.setCellFormula('Sheet1', formulaAction.address, formulaAction.formula)
    expect(engine.getCell('Sheet1', 'C5').format).toBe('0.00')

    engine.deleteRows('Sheet1', 3, 2)

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'C5').format).toBe('0.00')
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(await exportReplaySnapshot(seedSnapshot, [moveAction, formulaAction])),
    )
  })

  it('rebinds structurally rewritten formulas when dependency addresses shift after prior ref errors', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-ref-error-rebind-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'C3+D4')
    engine.deleteRows('Sheet1', 2, 1)
    engine.insertColumns('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('#REF!+E3')

    engine.deleteColumns('Sheet1', 4, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('#REF!+#REF!')
    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'B1').formula).toBe('#REF!+E3')
  })

  it('exports structurally moved formula sources at the clamped column target', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'clamped-column-move-formula-source-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'E5' }, [
      [1, 2, 3, 4, 5],
      [11, 12, 13, 14, 15],
      [21, 22, 23, 24, 25],
      [31, 32, 33, 34, 35],
      [41, 42, 43, 44, 45],
    ])
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    engine.moveColumns('Sheet1', 0, 1, 5)

    expect(engine.getCell('Sheet1', 'A6').formula).toBe('E1+A1')
    expect(engine.getCell('Sheet1', 'A6').value).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCell('Sheet1', 'E6').formula).toBe('SUM(E1:E5)')
    expect(engine.getCell('Sheet1', 'E6').value).toEqual({ tag: ValueTag.Number, value: 105 })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCell('Sheet1', 'A6').formula).toBe('SUM(A1:A5)')
    expect(engine.getCell('Sheet1', 'B6').formula).toBe('A1+B1')

    expect(engine.redo()).toBe(true)
    expect(engine.getCell('Sheet1', 'A6').formula).toBe('E1+A1')
    expect(engine.getCell('Sheet1', 'E6').formula).toBe('SUM(E1:E5)')

    const restored = new SpreadsheetEngine({ workbookName: 'clamped-column-move-formula-source-restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A6').formula).toBe('E1+A1')
    expect(restored.getCell('Sheet1', 'A6').value).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(restored.getCell('Sheet1', 'E6').formula).toBe('SUM(E1:E5)')
    expect(restored.getCell('Sheet1', 'E6').value).toEqual({ tag: ValueTag.Number, value: 105 })
  })

  it('does not record history for duplicate formula writes', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('blank', 'duplicate-formula-history-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'duplicate-formula-history-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.setCellFormula('Sheet1', 'E3', 'F6/C3')
    const afterFirstWrite = engine.exportSnapshot()

    engine.setCellFormula('Sheet1', 'E3', 'F6/C3')
    expect(engine.exportSnapshot()).toEqual(afterFirstWrite)

    const sheetId = seedSnapshot.sheets[0]?.id
    if (sheetId === undefined) {
      throw new Error('Expected blank seed to include Sheet1 id')
    }
    engine.setCellFormulaAt(sheetId, 2, 4, 'F6/C3')
    expect(engine.exportSnapshot()).toEqual(afterFirstWrite)

    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(false)
  })

  it('rejects coordinate formula writes for unknown sheet ids', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'unknown-sheet-id-write-regression' })
    await engine.ready()

    expect(() => engine.setCellFormulaAt(999_999, 0, 0, 'A1')).toThrow('Unknown sheet id: 999999')
  })

  it('does not record history for duplicate conditional format writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'duplicate-conditional-format-history-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const format = {
      id: 'duplicate-cf',
      range: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B2' },
      rule: { kind: 'cellIs', operator: 'greaterThan', values: [10] },
      style: { fill: { backgroundColor: '#ff0000' } },
      priority: 1,
    } satisfies WorkbookConditionalFormatSnapshot

    engine.setConditionalFormat(format)
    const afterFirstWrite = engine.exportSnapshot()

    engine.setConditionalFormat(format)
    expect(engine.exportSnapshot()).toEqual(afterFirstWrite)
  })

  it('restores formula graphs after undoing mixed row inserts and column deletes', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'structural-formula-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'structural-formula-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertRows('Sheet1', 0, 1)
    engine.deleteColumns('Sheet1', 0, 1)
    engine.insertRows('Sheet1', 0, 1)

    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(seedSnapshot)
  })

  it('captures family-deferred formula sources before structural delete undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('cross-sheet-graph', 'family-deferred-formula-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'family-deferred-formula-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'C3' })
    engine.insertColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'D2').formula).toBe('C2*Summary!B2')

    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('B2*Summary!B2')

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'D2').formula).toBe('C2*Summary!B2')

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'C2').formula).toBe('B2*Summary!B2')

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(seedSnapshot)
  })

  it('rewrites non-family symbolic formulas after consecutive family-deferred column inserts', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('cross-sheet-graph', 'mixed-family-symbolic-column-insert-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'mixed-family-symbolic-column-insert-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertColumns('Sheet1', 0, 1)
    engine.insertColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'E3').formula).toBe('D3*(1+TaxCell)')
    engine.insertRows('Sheet1', 0, 1)

    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(seedSnapshot)
  })

  it('keeps named-range aggregate formulas symbolic across structural column undo', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('named-structures', 'structural-named-range-symbolic-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'structural-named-range-symbolic-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertColumns('Sheet1', 2, 1)
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 33 })

    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A3',
      },
    })
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 30 })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 33 })

    expect(engine.undo()).toBe(true)
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(normalizeSnapshotForSemanticComparison(seedSnapshot))
  })

  it('imports snapshots with structurally invalidated formula dependencies as ref errors', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'invalid-range-dependency-import-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'invalid-range-dependency-import-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.deleteColumns('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'A1', 'A1+A1')
    engine.deleteRows('Sheet1', 0, 1)
    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
    )
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, [['north']])
    engine.insertRows('Sheet1', 0, 1)
    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
    )

    const snapshot = engine.exportSnapshot()
    const restored = new SpreadsheetEngine({
      workbookName: snapshot.workbook.name,
      replicaId: 'invalid-range-dependency-import-restored',
    })
    await restored.ready()

    expect(() => restored.importSnapshot(structuredClone(snapshot))).not.toThrow()
  })

  it('restores structurally rewritten formula templates after refs collapse to #REF', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'structural-template-ref-restore-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'structural-template-ref-restore-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.deleteColumns('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'A1', 'A1+A1')
    engine.deleteColumns('Sheet1', 0, 1)
    engine.deleteRows('Sheet1', 0, 1)
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
    )

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets[0]?.cells).toContainEqual({ address: 'C2', formula: 'SUM(#REF!)' })

    const restored = new SpreadsheetEngine({
      workbookName: snapshot.workbook.name,
      replicaId: 'structural-template-ref-restore-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.exportSnapshot()).toEqual(snapshot)
  })

  it('clears target formats when moving empty cells over formatted blanks and keeps undo aligned', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'move-empty-format-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B2' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
    )

    expect(engine.getCell('Sheet1', 'A1').format).toBe('0.00')

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'B4', endAddress: 'C4' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
    )

    expect(engine.getCell('Sheet1', 'A1').format).toBeUndefined()
    expect(engine.exportSnapshot().sheets[0]?.cells ?? []).toEqual([])

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'A1').format).toBe('0.00')
  })

  it('restores explicit formats on deleted formula cells during structural undo', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'delete-formula-format-undo-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'B4', endAddress: 'C4' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
      { sheetName: 'Sheet1', startAddress: 'E4', endAddress: 'F5' },
    )
    engine.deleteRows('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'E4', 'A1+A1')

    expect(engine.getCell('Sheet1', 'E4').format).toBe('0.00')

    engine.deleteRows('Sheet1', 2, 2)

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'E4')).toMatchObject({
      formula: 'A1+A1',
      format: '0.00',
    })
  })

  it('propagates cycle errors to dependent formulas after direct formula writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cycle-dependent-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'SUM(B1:B1)')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(B1:C2)')
    engine.setCellFormula('Sheet1', 'A2', 'SUM(B1:C2)')

    expect(engine.getCell('Sheet1', 'C1').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCell('Sheet1', 'A2').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('propagates cycle errors through range dependents after direct formula writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cycle-range-dependent-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 'text')
    engine.setCellValue('Sheet1', 'C2', 'text')
    engine.setCellFormula('Sheet1', 'B3', 'SUM(A1:C3)')
    engine.setCellFormula('Sheet1', 'E4', 'SUM(B1:C4)')

    expect(engine.getCell('Sheet1', 'B3').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCell('Sheet1', 'E4').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('preserves cycle errors for self-referential range formulas after CSV roundtrip import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-cycle-roundtrip-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, [[false]])
    engine.setCellFormula('Sheet1', 'A2', 'A1+A4')
    engine.setCellFormula('Sheet1', 'A3', 'SUM(A1:A4)')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'A4' }, [['text:lB<`x']])
    engine.setCellFormula('Sheet1', 'A5', 'A1+A1')

    expect(engine.getCell('Sheet1', 'A3').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })

    const restored = new SpreadsheetEngine({ workbookName: 'csv-cycle-roundtrip-regression-restored' })
    await restored.ready()
    restored.createSheet('Sheet1')
    restored.importSheetCsv('Sheet1', engine.exportSheetCsv('Sheet1'))

    expect(restored.getCell('Sheet1', 'A3').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('settles imported range dependents of cycle formulas in one CSV import transaction', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-cycle-dependent-import-regression' })
    await engine.ready()

    engine.importSheetCsv(
      'Sheet1',
      [
        'text:&.Ape4.n,FALSE,=SUM(A1:D3),FALSE,=SUM(A1:D3)',
        '=SUM(A1:A1),401015777,TRUE,FALSE,-524808229',
        '"=IF(A1>0,""text:yes"",""text:no"")",=A1+A1,"=IF(A1>0,""text:yes"",""text:no"")",text:ps<n,text:#',
      ].join('\n'),
    )

    expect(engine.getCell('Sheet1', 'C1').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCell('Sheet1', 'E1').value).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('preserves shifted range sum precision after CSV roundtrip import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-shifted-sum-precision-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'SUM(A5:A5)')
    engine.setCellFormula('Sheet1', 'A2', 'A5*A5')
    engine.setCellFormula('Sheet1', 'A3', 'IF(A5>0,"text:yes","text:no")')
    engine.setCellFormula('Sheet1', 'A4', 'IF(A5>0,"text:yes","text:no")')
    engine.setCellValue('Sheet1', 'A5', 1429783918)

    const restored = new SpreadsheetEngine({ workbookName: 'csv-shifted-sum-precision-regression-restored' })
    await restored.ready()
    restored.createSheet('Sheet1')
    restored.importSheetCsv('Sheet1', engine.exportSheetCsv('Sheet1'))

    expect(restored.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 1429783918,
    })
  })

  it('preserves narrow SUM range precision after CSV roundtrip import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-sum-prefix-precision-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'SUM(A5:A5)')
    engine.setCellFormula('Sheet1', 'A2', 'A5+A5')
    engine.setCellFormula('Sheet1', 'A3', 'A5+A5')
    engine.setCellFormula('Sheet1', 'A4', 'A5*A5')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A5', endAddress: 'A5' }, [[663897248]])

    const restored = new SpreadsheetEngine({ workbookName: 'csv-sum-prefix-precision-regression-restored' })
    await restored.ready()
    restored.createSheet('Sheet1')
    restored.importSheetCsv('Sheet1', engine.exportSheetCsv('Sheet1'))

    expect(restored.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 663897248,
    })
  })

  it('serializes structurally shifted direct aggregate formulas with their preserved value dependencies', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-structural-source-regression',
      replicaId: 'direct-aggregate-structural-source-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, row * 10 + 1)
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')

    engine.insertRows('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'A7').formula).toBe('SUM(A2:A6)')
    expect(engine.getCellValue('Sheet1', 'A7')).toEqual({ tag: ValueTag.Number, value: 105 })

    const restored = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-structural-source-regression-restored',
      replicaId: 'direct-aggregate-structural-source-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A7').formula).toBe('SUM(A2:A6)')
    expect(restored.getCellValue('Sheet1', 'A7')).toEqual({ tag: ValueTag.Number, value: 105 })
  })

  it('does not preserve stale direct aggregate values when row moves create a self-reference', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-move-self-reference-regression',
      replicaId: 'direct-aggregate-move-self-reference-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, row * 10 + 1)
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')

    engine.moveRows('Sheet1', 0, 1, 5)

    expect(engine.getCell('Sheet1', 'A5').formula).toBe('SUM(A1:A6)')
    expect(engine.getCellValue('Sheet1', 'A5')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })

    const restored = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-move-self-reference-regression-restored',
      replicaId: 'direct-aggregate-move-self-reference-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCellValue('Sheet1', 'A5')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'A6').formula).toBe('SUM(A1:A5)')
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 105 })
  })

  it('restores self-referential range formulas as cycles after undoing structural deletes', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-delete-undo-self-reference-regression',
      replicaId: 'direct-aggregate-delete-undo-self-reference-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, row * 10 + 1)
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')

    engine.moveRows('Sheet1', 0, 1, 5)
    engine.insertRows('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'A6').formula).toBe('SUM(A2:A7)')
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })

    engine.deleteRows('Sheet1', 5, 1)
    expect(engine.undo()).toBe(true)

    expect(engine.getCell('Sheet1', 'A6').formula).toBe('SUM(A2:A7)')
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })
  })

  it('restores formulas with deferred move rewrites after undoing structural deletes', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'deferred-move-rewrite-delete-undo-regression',
      replicaId: 'deferred-move-rewrite-delete-undo-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      for (let col = 0; col < 3; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row + 1}`, row * 10 + col + 1)
      }
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    engine.moveColumns('Sheet1', 0, 1, 2)
    expect(engine.getCell('Sheet1', 'A6').formula).toBe('C1+A1')
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 3 })

    engine.deleteColumns('Sheet1', 2, 1)
    expect(engine.getCell('Sheet1', 'A6').formula).toBe('#REF!+A1')
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'A6').formula).toBe('C1+A1')
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCell('Sheet1', 'C6').formula).toBe('SUM(C1:C5)')
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 105 })
  })

  it('rewrites direct aggregate ranges when a moved row shifts the referenced interval boundary', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-move-boundary-regression',
      replicaId: 'direct-aggregate-move-boundary-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, row * 10 + 1)
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')

    engine.insertRows('Sheet1', 0, 1)
    engine.moveRows('Sheet1', 0, 1, 1)

    expect(engine.getCell('Sheet1', 'A7').formula).toBe('SUM(A1:A6)')
    expect(engine.getCellValue('Sheet1', 'A7')).toEqual({ tag: ValueTag.Number, value: 105 })

    const restored = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-move-boundary-regression-restored',
      replicaId: 'direct-aggregate-move-boundary-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A7').formula).toBe('SUM(A1:A6)')
    expect(restored.getCellValue('Sheet1', 'A7')).toEqual({ tag: ValueTag.Number, value: 105 })
  })

  it('restores deferred structural formula sources without trusting stale template ids', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'deferred-structural-template-restore-regression',
      replicaId: 'deferred-structural-template-restore-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      for (let col = 0; col < 5; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row + 1}`, row * 10 + col + 1)
      }
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    const applyWithUndoRedo = (apply: () => void): void => {
      apply()
      expect(engine.undo()).toBe(true)
      expect(engine.redo()).toBe(true)
    }
    applyWithUndoRedo(() => engine.setCellFormula('Sheet1', 'A3', 'A1+B1'))
    applyWithUndoRedo(() => engine.moveRows('Sheet1', 1, 1, 0))
    applyWithUndoRedo(() => engine.moveRows('Sheet1', 2, 1, 5))
    applyWithUndoRedo(() => engine.insertColumns('Sheet1', 0, 1))

    expect(engine.getCell('Sheet1', 'B6').formula).toBe('B2+C2')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })

    const restored = new SpreadsheetEngine({
      workbookName: 'deferred-structural-template-restore-regression-restored',
      replicaId: 'deferred-structural-template-restore-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'B6').formula).toBe('B2+C2')
    expect(restored.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })
  })

  it('does not rewrite formulas for sparse moves from outside materialized columns', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'sparse-outside-column-move-regression',
      replicaId: 'sparse-outside-column-move-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      for (let col = 0; col < 5; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row + 1}`, row * 10 + col + 1)
      }
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    engine.moveColumns('Sheet1', 5, 1, 0)

    expect(engine.getCell('Sheet1', 'A6').formula).toBe('SUM(A1:A5)')
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(engine.getCell('Sheet1', 'B6').formula).toBe('A1+B1')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })

    const restored = new SpreadsheetEngine({
      workbookName: 'sparse-outside-column-move-regression-restored',
      replicaId: 'sparse-outside-column-move-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A6').formula).toBe('SUM(A1:A5)')
    expect(restored.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(restored.getCell('Sheet1', 'B6').formula).toBe('A1+B1')
    expect(restored.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })
  })

  it('updates rectangular direct aggregate values when structural row deletes remove part of the range', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'rectangular-direct-aggregate-delete-regression',
      replicaId: 'rectangular-direct-aggregate-delete-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [
      [1, 2],
      [11, 12],
    ])
    engine.setCellFormula('Sheet1', 'A3', 'SUM(A1:B2)')

    engine.deleteRows('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'A2').formula).toBe('SUM(A1:B1)')
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 23 })

    const restored = new SpreadsheetEngine({
      workbookName: 'rectangular-direct-aggregate-delete-regression-restored',
      replicaId: 'rectangular-direct-aggregate-delete-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A2').formula).toBe('SUM(A1:B1)')
    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 23 })
  })

  it('recalculates direct aggregate moves when a range expands across moved cells', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-move-expanded-range-regression',
      replicaId: 'direct-aggregate-move-expanded-range-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C2' }, [
      [1, 2, 3],
      [11, 12, 13],
    ])
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B2)')

    engine.moveColumns('Sheet1', 0, 1, 2)

    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(A1:C2)')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 42 })

    const restored = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-move-expanded-range-regression-restored',
      replicaId: 'direct-aggregate-move-expanded-range-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'D1').formula).toBe('SUM(A1:C2)')
    expect(restored.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 42 })
  })

  it('restores snapshots when a formula runtime image contains a stale template id', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'stale-template-runtime-image-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'stale-template-runtime-image-primary',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const actions: CoreAction[] = [
      { kind: 'deleteColumns', start: 0, count: 1 },
      { kind: 'formula', address: 'A1', formula: 'A1+A1' },
      { kind: 'deleteRows', start: 0, count: 1 },
      {
        kind: 'format',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        format: '0.00',
      },
      {
        kind: 'clear',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      },
      {
        kind: 'style',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        patch: { fill: { backgroundColor: '#dbeafe' } },
      },
    ]
    actions.forEach((action) => applyActionAndCaptureResult(engine, action))

    const snapshot = engine.exportSnapshot()
    const restored = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'stale-template-runtime-image-restored',
    })
    await restored.ready()

    expect(() => restored.importSnapshot(structuredClone(snapshot))).not.toThrow()
    expect(normalizeSnapshotForSemanticComparison(restored.exportSnapshot())).toEqual(normalizeSnapshotForSemanticComparison(snapshot))
  })

  it('normalizes sparse style metadata by covered cells after undo replay', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('sparse-format', 'history-style-run-normalization-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'history-style-run-normalization',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const applied: CoreAction[] = []
    const applyAccepted = (action: CoreAction) => {
      const result = applyActionAndCaptureResult(engine, action)
      if (result.accepted) {
        applied.push(action)
      }
    }

    applyAccepted({ kind: 'insertRows', start: 0, count: 1 })
    applyAccepted({
      kind: 'style',
      range: { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'E3' },
      patch: { alignment: { horizontal: 'right', wrap: true } },
    })
    applyAccepted({ kind: 'insertRows', start: 0, count: 1 })
    applyAccepted({ kind: 'deleteColumns', start: 0, count: 1 })

    expect(engine.undo()).toBe(true)
    expect(applied.pop()?.kind).toBe('deleteColumns')

    applyAccepted({
      kind: 'style',
      range: { sheetName: 'Sheet1', startAddress: 'F3', endAddress: 'F3' },
      patch: { alignment: { horizontal: 'right', wrap: true } },
    })

    const expectedSnapshot = await exportReplaySnapshot(seedSnapshot, applied)

    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(expectedSnapshot),
    )
  })

  it('rewrites rectangular direct aggregate formulas when deleting covered columns', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-delete-column-regression',
      replicaId: 'direct-aggregate-delete-column-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C2' }, [
      [1, 2, 3],
      [11, 12, 13],
    ])
    engine.setCellFormula('Sheet1', 'A3', 'SUM(A1:B2)')

    engine.deleteColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'A3').formula).toBe('SUM(A1:A2)')
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 12 })

    const restored = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-delete-column-regression-restored',
      replicaId: 'direct-aggregate-delete-column-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A3').formula).toBe('SUM(A1:A2)')
    expect(restored.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 12 })
  })

  it('recalculates direct aggregate row deletes when surviving range members are formulas', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-delete-row-surviving-formula-regression',
      replicaId: 'direct-aggregate-delete-row-surviving-formula-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B5' }, [
      [1, 2],
      [11, 12],
      [21, 22],
      [31, 32],
      [41, 42],
    ])
    engine.setCellFormula('Sheet1', 'A3', 'SUM(A1:B2)')
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')

    engine.deleteRows('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'A2').formula).toBe('SUM(A1:B1)')
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 23 })
    expect(engine.getCell('Sheet1', 'A5').formula).toBe('SUM(A1:A4)')
    expect(engine.getCellValue('Sheet1', 'A5')).toEqual({ tag: ValueTag.Number, value: 106 })

    const restored = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-delete-row-surviving-formula-regression-restored',
      replicaId: 'direct-aggregate-delete-row-surviving-formula-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A5').formula).toBe('SUM(A1:A4)')
    expect(restored.getCellValue('Sheet1', 'A5')).toEqual({ tag: ValueTag.Number, value: 106 })
  })

  it('does not preserve stale direct aggregate values when deleting the full referenced row range', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-delete-full-row-range-regression',
      replicaId: 'direct-aggregate-delete-full-row-range-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [
      [1, 2],
      [11, 12],
    ])
    engine.setCellFormula('Sheet1', 'A3', 'SUM(A1:B2)')

    engine.deleteRows('Sheet1', 0, 2)

    expect(engine.getCell('Sheet1', 'A1').formula).toBe('SUM(#REF!)')
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })

    const restored = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-delete-full-row-range-regression-restored',
      replicaId: 'direct-aggregate-delete-full-row-range-regression-restored',
    })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A1').formula).toBe('SUM(#REF!)')
    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('recomputes direct aggregate values when undo restores a replaced formula', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-aggregate-formula-replacement-undo-regression',
      replicaId: 'direct-aggregate-formula-replacement-undo-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'E5' }, [
      [1, 2, 3, 4, 5],
      [11, 12, 13, 14, 15],
      [21, 22, 23, 24, 25],
      [31, 32, 33, 34, 35],
      [41, 42, 43, 44, 45],
    ])
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    engine.insertColumns('Sheet1', 0, 1)
    engine.moveColumns('Sheet1', 1, 1, 4)
    expect(engine.getCell('Sheet1', 'E6').formula).toBe('SUM(E1:E5)')
    expect(engine.getCellValue('Sheet1', 'E6')).toEqual({ tag: ValueTag.Number, value: 105 })

    engine.setCellFormula('Sheet1', 'E6', 'A1+B1')
    expect(engine.getCellValue('Sheet1', 'E6')).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'E6').formula).toBe('SUM(E1:E5)')
    expect(engine.getCellValue('Sheet1', 'E6')).toEqual({ tag: ValueTag.Number, value: 105 })

    expect(engine.redo()).toBe(true)
    expect(engine.getCell('Sheet1', 'E6').formula).toBe('A1+B1')
    expect(engine.getCellValue('Sheet1', 'E6')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('replays row-move formula rewrites exactly through redo history', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'row-move-formula-redo-history-regression',
      replicaId: 'row-move-formula-redo-history-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'E5' }, [
      [1, 2, 3, 4, 5],
      [11, 12, 13, 14, 15],
      [21, 22, 23, 24, 25],
      [31, 32, 33, 34, 35],
      [41, 42, 43, 44, 45],
    ])
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    engine.deleteRows('Sheet1', 1, 1)
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')
    engine.moveRows('Sheet1', 0, 1, 1)

    expect(engine.getCell('Sheet1', 'B6').formula).toBe('A3+B3')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 63 })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'B6').formula).toBe('A1+B1')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })

    expect(engine.redo()).toBe(true)
    expect(engine.getCell('Sheet1', 'B6').formula).toBe('A3+B3')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 63 })
  })

  it('restores formula sources when undoing column inserts after replayed row deletes', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'column-insert-formula-undo-history-regression',
      replicaId: 'column-insert-formula-undo-history-regression',
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'E5' }, [
      [1, 2, 3, 4, 5],
      [11, 12, 13, 14, 15],
      [21, 22, 23, 24, 25],
      [31, 32, 33, 34, 35],
      [41, 42, 43, 44, 45],
    ])
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    engine.deleteRows('Sheet1', 1, 1)
    expect(engine.undo()).toBe(true)
    expect(engine.redo()).toBe(true)

    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')
    expect(engine.undo()).toBe(true)
    expect(engine.redo()).toBe(true)

    engine.insertColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'C6').formula).toBe('B1+C1')
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 3 })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'B6').formula).toBe('A1+B1')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })

    expect(engine.redo()).toBe(true)
    expect(engine.getCell('Sheet1', 'C6').formula).toBe('B1+C1')
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 3 })
  })
})
