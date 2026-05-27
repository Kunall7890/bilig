import { isDeepStrictEqual } from 'node:util'
import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import type { EngineOpBatch } from '@bilig/workbook'
import { SpreadsheetEngine } from '../engine.js'
import { makeCellKey } from '../workbook-store.js'
import { runProperty } from '@bilig/test-fuzz'
import { createEngineSeedSnapshot } from './engine-fuzz-helpers.js'
import {
  applyAction,
  assertSnapshotInvariants,
  correctnessActionArbitrary,
  createBaselineSnapshot,
  redoAll,
  sheetName,
  toRangeRef,
  undoAll,
} from './engine-correctness-helpers.js'

describe('engine correctness', () => {
  it('clears sparse style and format metadata when undoing structural edit sequences', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-undo-sparse-ranges')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-undo-sparse-ranges',
      replicaId: 'correctness-undo-sparse-ranges',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.insertColumns(sheetName, 0, 1)
    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })
    engine.deleteColumns(sheetName, 0, 1)
    engine.setRangeNumberFormat({ sheetName, startAddress: 'A1', endAddress: 'A1' }, '0.00')

    expect(undoAll(engine, 16)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('prunes materialized empty cells when undo restores a blank address', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-undo-empty-cell-prune')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-undo-empty-cell-prune',
      replicaId: 'correctness-undo-empty-cell-prune',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'A1' }, [[false]])
    engine.insertRows(sheetName, 0, 1)
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'A1' }, [[0]])
    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })

    expect(undoAll(engine, 16)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('replays explicit blanks after styled cell value undo and redo', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-redo-explicit-blank')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-redo-explicit-blank',
      replicaId: 'correctness-redo-explicit-blank',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })
    engine.setRangeNumberFormat({ sheetName, startAddress: 'A1', endAddress: 'A1' }, '0.00')
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'A1' }, [[false]])
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'A1' }, [[null]])

    const finalSnapshot = engine.exportSnapshot()
    expect(finalSnapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: null }])

    const undoCount = undoAll(engine, 16)
    expect(undoCount).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)

    const redoCount = redoAll(engine, undoCount + 2)
    expect(redoCount).toBe(undoCount)
    expect(engine.exportSnapshot()).toEqual(finalSnapshot)
  })

  it('replays sparse inserted columns exactly across undo and redo', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-redo-column-identity')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-redo-column-identity',
      replicaId: 'correctness-redo-column-identity',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula(sheetName, 'A1', 'A1+A1')
    engine.insertColumns(sheetName, 0, 1)
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'A1' }, [[null]])
    engine.setRangeNumberFormat({ sheetName, startAddress: 'A1', endAddress: 'A1' }, '0.00')

    const finalSnapshot = engine.exportSnapshot()
    expect(finalSnapshot.sheets[0]?.metadata?.columns).toBeUndefined()

    const undoCount = undoAll(engine, 16)
    expect(undoCount).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)

    const redoCount = redoAll(engine, undoCount + 2)
    expect(redoCount).toBe(undoCount)
    expect(engine.exportSnapshot()).toEqual(finalSnapshot)
  })

  it('does not leave empty cells behind when fill replays a blank source cell', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-fill-empty-prune')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-fill-empty-prune',
      replicaId: 'correctness-fill-empty-prune',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula(sheetName, 'B1', 'A1+A1')
    engine.deleteColumns(sheetName, 0, 1)
    engine.fillRange({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { sheetName, startAddress: 'A1', endAddress: 'A1' })
    engine.clearRange({ sheetName, startAddress: 'A1', endAddress: 'A1' })

    expect(undoAll(engine, 16)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('prunes translated dependency cells after structural formula rebuild undo', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-structural-dependency-prune')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-dependency-prune',
      replicaId: 'correctness-structural-dependency-prune',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula(sheetName, 'B1', 'A1+D4')
    engine.deleteRows(sheetName, 1, 1)
    engine.insertRows(sheetName, 0, 1)
    engine.insertRows(sheetName, 0, 1)
    engine.deleteRows(sheetName, 3, 2)

    expect(undoAll(engine, 20)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('does not preserve temporary null dependency placeholders as authored blanks during undo', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-undo-temporary-blank-prune')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-undo-temporary-blank-prune',
      replicaId: 'correctness-undo-temporary-blank-prune',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula(sheetName, 'A1', 'C3+A1')
    engine.setRangeValues({ sheetName, startAddress: 'B3', endAddress: 'C3' }, [[0, false]])
    engine.deleteRows(sheetName, 0, 1)
    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })
    engine.setCellFormula(sheetName, 'A1', 'A1+A1')
    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })

    expect(undoAll(engine, 24)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('does not reify inherited number formats into explicit cells during structural undo', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-structural-format-prune')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-format-prune',
      replicaId: 'correctness-structural-format-prune',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setRangeValues({ sheetName, startAddress: 'A3', endAddress: 'A3' }, [[0]])
    engine.setRangeNumberFormat({ sheetName, startAddress: 'A3', endAddress: 'A3' }, '0.00')
    engine.clearRange({ sheetName, startAddress: 'A1', endAddress: 'A1' })
    engine.deleteRows(sheetName, 0, 1)
    engine.deleteColumns(sheetName, 0, 1)
    engine.clearRange({ sheetName, startAddress: 'A1', endAddress: 'A1' })

    expect(undoAll(engine, 24)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('does not materialize inherited format-range placeholders during snapshot roundtrip', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-snapshot-format-range',
      replicaId: 'correctness-snapshot-format-range',
    })
    await engine.ready()
    engine.createSheet(sheetName)

    engine.setCellFormula(sheetName, 'A1', 'C4+C3')
    engine.setRangeNumberFormat({ sheetName, startAddress: 'B4', endAddress: 'C4' }, '0.00')
    engine.clearRange({ sheetName, startAddress: 'B4', endAddress: 'C4' })

    const snapshot = engine.exportSnapshot()
    const restored = new SpreadsheetEngine({
      workbookName: snapshot.workbook.name,
      replicaId: 'correctness-snapshot-format-range-restored',
    })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.exportSnapshot()).toEqual(snapshot)
  })

  it('preserves rewritten error formulas across structural delete undo', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-formula-error-undo',
      replicaId: 'correctness-structural-formula-error-undo',
    })
    await engine.ready()
    engine.createSheet(sheetName)

    engine.setCellFormula(sheetName, 'A1', 'A1+D4')
    engine.deleteRows(sheetName, 2, 2)

    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: 'correctness-structural-formula-error-undo' },
      sheets: [
        {
          id: 1,
          name: sheetName,
          order: 0,
          cells: [{ address: 'A1', formula: 'A1+#REF!' }],
        },
      ],
    })

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: 'correctness-structural-formula-error-undo' },
      sheets: [
        {
          id: 1,
          name: sheetName,
          order: 0,
          cells: [{ address: 'A1', formula: 'A1+D4' }],
        },
      ],
    })
  })

  it('restores metadata-only number formats after structural delete undo', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-format-only-undo',
      replicaId: 'correctness-structural-format-only-undo',
    })
    await engine.ready()
    engine.createSheet(sheetName)

    engine.setRangeNumberFormat({ sheetName, startAddress: 'A1', endAddress: 'A1' }, '0.00')
    const initialSnapshot = engine.exportSnapshot()

    engine.deleteColumns(sheetName, 0, 1)
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('restores named-range structures after insert-column undo', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-named-range-undo',
      replicaId: 'correctness-structural-named-range-undo',
    })
    await engine.ready()
    engine.createSheet(sheetName)
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'B3' }, [
      ['Qty', 'Amount'],
      [1, 10],
      [2, 20],
    ])
    engine.setDefinedName('SalesRange', {
      kind: 'range-ref',
      sheetName,
      startAddress: 'A1',
      endAddress: 'B3',
    })
    engine.setTable({
      name: 'Sales',
      sheetName,
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula(sheetName, 'C1', 'SUM(SalesRange)')

    const initialSnapshot = engine.exportSnapshot()

    engine.insertColumns(sheetName, 0, 1)
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName,
        startAddress: 'B1',
        endAddress: 'C3',
      },
    })
    expect(engine.getTable('Sales')).toEqual({
      name: 'Sales',
      sheetName,
      startAddress: 'B1',
      endAddress: 'C3',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue(sheetName, 'D1')).toMatchObject({ tag: 1, value: 33 })

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName,
        startAddress: 'A1',
        endAddress: 'B3',
      },
    })
    expect(engine.getTable('Sales')).toEqual({
      name: 'Sales',
      sheetName,
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    expect(engine.getCell(sheetName, 'C1').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue(sheetName, 'C1')).toMatchObject({ tag: 1, value: 33 })
  })

  it('restores shifted relative references after insert-column undo', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-relative-reference-undo',
      replicaId: 'correctness-structural-relative-reference-undo',
    })
    await engine.ready()
    engine.createSheet(sheetName)
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'B2' }, [
      [4, 2],
      [3, 7],
    ])
    engine.setCellFormula(sheetName, 'C1', 'A1+B2')
    engine.setCellFormula(sheetName, 'D2', 'C1*A2')
    engine.setCellFormula(sheetName, 'E3', 'SUM(A1:B2)')

    const initialSnapshot = engine.exportSnapshot()
    const sheetId = engine.workbook.getSheet(sheetName)!.id

    engine.insertColumns(sheetName, 1, 2)
    expect(engine.workbook.cellKeyToIndex.get(makeCellKey(sheetId, 0, 4))).toBe(engine.workbook.getCellIndex(sheetName, 'E1'))
    expect(engine.getCell(sheetName, 'E1').formula).toBe('A1+D2')
    expect(engine.getCell(sheetName, 'F2').formula).toBe('E1*A2')
    expect(engine.getCell(sheetName, 'G3').formula).toBe('SUM(A1:D2)')

    expect(engine.undo()).toBe(true)
    expect(engine.workbook.cellKeyToIndex.get(makeCellKey(sheetId, 0, 2))).toBe(engine.workbook.getCellIndex(sheetName, 'C1'))
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
    expect(engine.getCell(sheetName, 'C1').formula).toBe('A1+B2')
    expect(engine.getCell(sheetName, 'D2').formula).toBe('C1*A2')
    expect(engine.getCell(sheetName, 'E3').formula).toBe('SUM(A1:B2)')
  })

  it('restores direct aggregate formula sources after delete-row undo overlaps the aggregate range', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-direct-aggregate-delete-row-undo',
      replicaId: 'correctness-structural-direct-aggregate-delete-row-undo',
    })
    await engine.ready()
    engine.createSheet(sheetName)
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'B2' }, [
      [4, 2],
      [3, 7],
    ])
    engine.setCellFormula(sheetName, 'E3', 'SUM(A1:B2)')

    const initialSnapshot = engine.exportSnapshot()

    engine.deleteRows(sheetName, 0, 1)
    expect(engine.getCell(sheetName, 'E2').formula).toBe('SUM(A1:B1)')

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
    expect(engine.getCell(sheetName, 'E3').formula).toBe('SUM(A1:B2)')
  })

  it('materializes deferred structural formula sources before consecutive axis edits', async () => {
    const initialSnapshot = await createEngineSeedSnapshot('formula-graph', 'correctness-consecutive-structural-formula-source')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-consecutive-structural-formula-source',
      replicaId: 'correctness-consecutive-structural-formula-source',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.fillRange({ sheetName, startAddress: 'C1', endAddress: 'C1' }, { sheetName, startAddress: 'D1', endAddress: 'D1' })
    engine.insertColumns(sheetName, 1, 1)
    expect(engine.getCell(sheetName, 'E1').formula).toBe('C1+D2')

    engine.insertRows(sheetName, 0, 1)
    expect(engine.getCell(sheetName, 'E2').formula).toBe('C2+D3')

    expect(engine.undo()).toBe(true)
    expect(engine.getCell(sheetName, 'E1').formula).toBe('C1+D2')
  })

  it('restores copied formula sources after undoing a row insert', async () => {
    const initialSnapshot = await createEngineSeedSnapshot('formula-graph', 'correctness-copied-formula-row-insert-undo')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-copied-formula-row-insert-undo',
      replicaId: 'correctness-copied-formula-row-insert-undo',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.copyRange({ sheetName, startAddress: 'D1', endAddress: 'D2' }, { sheetName, startAddress: 'D3', endAddress: 'D4' })
    expect(engine.getCell(sheetName, 'D4').formula).toBe('C3*A4')
    const afterCopySnapshot = engine.exportSnapshot()

    engine.insertRows(sheetName, 1, 1)
    expect(engine.getCell(sheetName, 'D5').formula).toBe('C3*A5')

    expect(engine.undo()).toBe(true)
    expect(engine.getCell(sheetName, 'D4').formula).toBe('C3*A4')
    expect(engine.exportSnapshot()).toEqual(afterCopySnapshot)
  })

  it('restores named-range structures after delete-row undo', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-named-range-delete-row-undo',
      replicaId: 'correctness-structural-named-range-delete-row-undo',
    })
    await engine.ready()
    engine.createSheet(sheetName)
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'B3' }, [
      ['Qty', 'Amount'],
      [1, 10],
      [2, 20],
    ])
    engine.setDefinedName('SalesRange', {
      kind: 'range-ref',
      sheetName,
      startAddress: 'A1',
      endAddress: 'B3',
    })
    engine.setTable({
      name: 'Sales',
      sheetName,
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula(sheetName, 'C1', 'SUM(SalesRange)')

    const initialSnapshot = engine.exportSnapshot()

    engine.deleteRows(sheetName, 2, 2)
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName,
        startAddress: 'A1',
        endAddress: 'B2',
      },
    })
    expect(engine.getTable('Sales')).toEqual({
      name: 'Sales',
      sheetName,
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName,
        startAddress: 'A1',
        endAddress: 'B3',
      },
    })
    expect(engine.getTable('Sales')).toEqual({
      name: 'Sales',
      sheetName,
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    expect(engine.getCell(sheetName, 'C1').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue(sheetName, 'C1')).toMatchObject({ tag: 1, value: 33 })
  })

  it('restores filter, sort, and validation metadata after delete-row undo', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-sheet-metadata-delete-row-undo',
      replicaId: 'correctness-structural-sheet-metadata-delete-row-undo',
    })
    await engine.ready()
    engine.createSheet(sheetName)
    engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'C3' }, [
      ['Id', 'Status', 'Amount'],
      [1, 'Draft', 10],
      [2, 'Final', 20],
    ])
    engine.setFilter(sheetName, { sheetName, startAddress: 'A1', endAddress: 'C3' })
    engine.setSort(sheetName, { sheetName, startAddress: 'A1', endAddress: 'C3' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setDataValidation({
      range: { sheetName, startAddress: 'B2', endAddress: 'B3' },
      rule: { kind: 'list', values: ['Draft', 'Final'] },
      allowBlank: false,
      showDropdown: true,
    })

    const initialSnapshot = engine.exportSnapshot()

    engine.deleteRows(sheetName, 0, 1)
    expect(engine.getFilters(sheetName)).toEqual([
      {
        sheetName,
        range: { sheetName, startAddress: 'A1', endAddress: 'C2' },
      },
    ])

    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
    expect(engine.getFilters(sheetName)).toEqual([
      {
        sheetName,
        range: { sheetName, startAddress: 'A1', endAddress: 'C3' },
      },
    ])
    expect(engine.getSorts(sheetName)).toEqual([
      {
        sheetName,
        range: { sheetName, startAddress: 'A1', endAddress: 'C3' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])
    expect(engine.getDataValidations(sheetName)).toEqual([
      {
        range: { sheetName, startAddress: 'B2', endAddress: 'B3' },
        rule: { kind: 'list', values: ['Draft', 'Final'] },
        allowBlank: false,
        showDropdown: true,
      },
    ])
  })

  it('prunes orphaned explicit formats after structural undo restores', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-undo-format-orphan-prune')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-undo-format-orphan-prune',
      replicaId: 'correctness-undo-format-orphan-prune',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setRangeNumberFormat({ sheetName, startAddress: 'A1', endAddress: 'B2' }, '0.00')
    engine.insertColumns(sheetName, 0, 1)
    engine.insertColumns(sheetName, 0, 1)
    engine.setCellFormula(sheetName, 'A1', 'D4+A1')
    engine.deleteColumns(sheetName, 1, 1)
    engine.fillRange({ sheetName, startAddress: 'B1', endAddress: 'C2' }, { sheetName, startAddress: 'B4', endAddress: 'C5' })
    engine.setRangeNumberFormat({ sheetName, startAddress: 'A1', endAddress: 'A1' }, '0.00')
    engine.insertColumns(sheetName, 0, 1)
    engine.deleteRows(sheetName, 3, 1)
    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })

    expect(undoAll(engine, 24)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('does not make inherited range formats explicit while undoing row deletes', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-structural-delete-inherited-format')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-delete-inherited-format',
      replicaId: 'correctness-structural-delete-inherited-format',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setRangeValues({ sheetName, startAddress: 'E5', endAddress: 'E6' }, [[0], ['north']])
    engine.setRangeNumberFormat({ sheetName, startAddress: 'D5', endAddress: 'E6' }, '0.00')
    engine.deleteRows(sheetName, 4, 2)

    expect(undoAll(engine, 8)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('coalesces adjacent style ranges before structural insert replay', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-style-range-coalesce',
      replicaId: 'correctness-style-range-coalesce',
    })
    await engine.ready()
    engine.createSheet(sheetName)
    engine.setRangeValues(toRangeRef(0, 0, 3, 0), [[1], [2], [3], [4]])

    engine.setRangeStyle(toRangeRef(2, 0, 3, 0), { fill: { backgroundColor: '#dbeafe' } })
    engine.setRangeStyle(toRangeRef(2, 0, 2, 0), { font: { bold: true } })

    expect(engine.undo()).toBe(true)
    engine.insertRows(sheetName, 3, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.styleRanges).toEqual([
      {
        range: {
          sheetName,
          startAddress: 'A3',
          endAddress: 'A5',
        },
        styleId: expect.any(String),
      },
    ])
  })

  it('drops orphaned formula dependents during structural undo replay', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-orphan-formula-undo')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-orphan-formula-undo',
      replicaId: 'correctness-orphan-formula-undo',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula(sheetName, 'A1', 'A1+B2')
    engine.setCellFormula(sheetName, 'B3', 'A1+A1')
    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })
    engine.deleteRows(sheetName, 1, 1)
    engine.insertRows(sheetName, 0, 1)
    engine.clearRange({ sheetName, startAddress: 'A1', endAddress: 'A1' })

    expect(undoAll(engine, 24)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
    expect(engine.getDependents(sheetName, 'A1')).toEqual({
      directPrecedents: [],
      directDependents: [],
    })
  })

  it('retargets cross-sheet formulas correctly after structural undo followed by another structural edit', async () => {
    const initialSnapshot = await createEngineSeedSnapshot('cross-sheet-graph', 'correctness-structural-undo-retarget')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-structural-undo-retarget',
      replicaId: 'correctness-structural-undo-retarget',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.insertRows('Sheet1', 0, 2)
    expect(engine.undo()).toBe(true)
    engine.deleteRows('Sheet1', 0, 2)

    expect(
      engine
        .exportSnapshot()
        .sheets.find((sheet) => sheet.name === 'Summary')
        ?.cells.find((cell) => cell.address === 'D1'),
    ).toEqual({
      address: 'D1',
      formula: 'SUM(Sheet1!B1:B1)',
    })
  })

  it('does not record undo history for empty fill and structural no-op sequences', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-empty-noop-history')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-empty-noop-history',
      replicaId: 'correctness-empty-noop-history',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.clearRange({ sheetName, startAddress: 'A1', endAddress: 'A1' })
    engine.fillRange({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { sheetName, startAddress: 'A1', endAddress: 'A1' })
    engine.clearRange({ sheetName, startAddress: 'A1', endAddress: 'A1' })
    engine.deleteColumns(sheetName, 0, 1)

    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
    expect(engine.undo()).toBe(false)
  })

  it('does not leave authored blank residue when undoing a formula after structural and null-write replay', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-formula-authored-blank-undo')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-formula-authored-blank-undo',
      replicaId: 'correctness-formula-authored-blank-undo',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula(sheetName, 'B4', 'A1+A1')
    engine.setRangeStyle({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { fill: { backgroundColor: '#dbeafe' } })
    engine.fillRange({ sheetName, startAddress: 'A1', endAddress: 'A1' }, { sheetName, startAddress: 'A1', endAddress: 'A1' })
    engine.insertColumns(sheetName, 0, 1)
    engine.setRangeValues({ sheetName, startAddress: 'B3', endAddress: 'C4' }, [
      [false, 'north'],
      ['north', null],
    ])

    expect(undoAll(engine, 16)).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('records undo history when inserting before an authored blank shifted from a formula cell', async () => {
    const initialSnapshot = await createBaselineSnapshot('correctness-insert-authored-blank-history')
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-insert-authored-blank-history',
      replicaId: 'correctness-insert-authored-blank-history',
    })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula(sheetName, 'E1', 'A1+A1')
    engine.setRangeValues({ sheetName, startAddress: 'D1', endAddress: 'E2' }, [
      [null, null],
      [false, null],
    ])
    engine.deleteColumns(sheetName, 0, 1)
    engine.insertColumns(sheetName, 3, 1)
    engine.clearRange({ sheetName, startAddress: 'A1', endAddress: 'A1' })
    engine.deleteRows(sheetName, 1, 1)

    const finalSnapshot = engine.exportSnapshot()
    const undoCount = undoAll(engine, 24)
    expect(undoCount).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)

    const redoCount = redoAll(engine, undoCount + 2)
    expect(redoCount).toBe(undoCount)
    expect(engine.exportSnapshot()).toEqual(finalSnapshot)
  })

  it('does not record extra undo history when clearing an explicit blank cell', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'correctness-explicit-blank-clear-history',
      replicaId: 'correctness-explicit-blank-clear-history',
    })
    await engine.ready()
    engine.importSnapshot({
      version: 1,
      workbook: { name: 'correctness-explicit-blank-clear-history' },
      sheets: [{ id: 1, name: sheetName, order: 0, cells: [{ address: 'B4', value: 'Review' }] }],
    })

    engine.setCellValue(sheetName, 'B4', null)
    engine.clearRange({ sheetName, startAddress: 'B4', endAddress: 'B4' })
    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(true)
    expect(engine.undo()).toBe(false)
  })

  it('reverses random local edit streams through undo and redo', async () => {
    await runProperty({
      suite: 'core/undo-redo-reversibility',
      arbitrary: fc.array(correctnessActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        const initialSnapshot = await createBaselineSnapshot('correctness-undo-redo')
        const engine = new SpreadsheetEngine({
          workbookName: 'correctness-undo-redo',
          replicaId: 'correctness-undo-redo',
        })
        await engine.ready()
        engine.importSnapshot(initialSnapshot)

        let observedSemanticChange = false
        for (const action of actions) {
          applyAction(engine, action)
          const currentSnapshot = engine.exportSnapshot()
          assertSnapshotInvariants(currentSnapshot)
          observedSemanticChange ||= !isDeepStrictEqual(currentSnapshot, initialSnapshot)
        }

        const finalSnapshot = engine.exportSnapshot()
        const undoCount = undoAll(engine, actions.length * 4)
        if (observedSemanticChange) {
          expect(undoCount).toBeGreaterThan(0)
        }
        expect(engine.exportSnapshot()).toEqual(initialSnapshot)

        const redoCount = redoAll(engine, undoCount + 2)
        expect(redoCount).toBe(undoCount)
        expect(engine.exportSnapshot()).toEqual(finalSnapshot)
      },
    })
  })

  it('replays captured local batches into an equivalent replica state', async () => {
    await runProperty({
      suite: 'core/local-batch-replay-parity',
      arbitrary: fc.array(correctnessActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        const initialSnapshot = await createBaselineSnapshot('correctness-replay')
        const primary = new SpreadsheetEngine({
          workbookName: 'correctness-replay',
          replicaId: 'primary',
        })
        const replica = new SpreadsheetEngine({
          workbookName: 'correctness-replay',
          replicaId: 'replica',
        })
        await Promise.all([primary.ready(), replica.ready()])

        const outbound: EngineOpBatch[] = []
        primary.subscribeBatches((batch) => outbound.push(batch))

        primary.importSnapshot(initialSnapshot)
        replica.importSnapshot(initialSnapshot)

        let appliedBatches = 0
        expect(replica.exportSnapshot()).toEqual(primary.exportSnapshot())

        for (const action of actions) {
          applyAction(primary, action)
          while (appliedBatches < outbound.length) {
            const nextBatch = outbound[appliedBatches]
            if (!nextBatch) {
              throw new Error(`Missing outbound batch at index ${appliedBatches}`)
            }
            replica.applyRemoteBatch(nextBatch)
            appliedBatches += 1
          }
          const primarySnapshot = primary.exportSnapshot()
          assertSnapshotInvariants(primarySnapshot)
          expect(replica.exportSnapshot()).toEqual(primarySnapshot)
        }
      },
    })
  })
})
