import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { createEngineSeedSnapshot, exportReplaySnapshot, normalizeSnapshotForSemanticComparison } from './engine-fuzz-helpers.js'

describe('structural fuzz regressions', () => {
  it('recalculates expanded direct aggregates after row moves', async () => {
    const engine = await createSeededEngine('direct-aggregate-row-move-regression')

    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:B2)')
    engine.moveRows('Sheet1', 2, 1, 1)

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 69 })
  })

  it('recalculates dependents of precomputed aggregate values after row deletes', async () => {
    const engine = await createSeededEngine('precomputed-aggregate-dependent-delete-regression')

    engine.moveRows('Sheet1', 5, 1, 0)
    engine.setCellFormula('Sheet1', 'B1', 'IF(A1>0,A1,0)')
    engine.deleteRows('Sheet1', 1, 1)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 104 })
  })

  it('restores formula dependency bindings after undoing column inserts', async () => {
    const engine = await createSeededEngine('column-insert-undo-binding-regression')

    engine.setCellFormula('Sheet1', 'C1', 'A1&"-"&B1')
    engine.insertColumns('Sheet1', 0, 1)
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({ tag: ValueTag.String, value: '1-2' })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('A1&"-"&B1')
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({ tag: ValueTag.String, value: '1-2' })

    engine.setCellValue('Sheet1', 'A1', 9)
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({ tag: ValueTag.String, value: '9-2' })
    engine.setCellValue('Sheet1', 'B1', 8)
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({ tag: ValueTag.String, value: '9-8' })
  })

  it('uses rewritten structural references on the JS fallback path before wasm is ready', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'deferred-structural-js-fallback-regression' })
    engine.createSheet('Sheet1')
    seedInitialGrid(engine)

    engine.setCellFormula('Sheet1', 'B1', 'IF(A1>0,A1,0)')
    engine.setCellFormula('Sheet1', 'C1', 'A1&"-"&B1')

    expect(engine.wasm.ready).toBe(false)
    engine.moveColumns('Sheet1', 0, 1, 2)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('C1&"-"&A1')
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({ tag: ValueTag.String, value: '1-1' })
    expect(engine.wasm.ready).toBe(false)
  })

  it('does not clear restored range formats while undoing structural deletes', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('pivot-analytics', 'format-range-delete-column-undo-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'format-range-delete-column-undo-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    const formatAction = {
      kind: 'format' as const,
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      format: '0.00',
    }
    engine.setRangeNumberFormat(formatAction.range, formatAction.format)
    engine.deleteColumns('Sheet1', 0, 1)

    expect(engine.undo()).toBe(true)
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(await exportReplaySnapshot(seedSnapshot, [formatAction])),
    )
  })

  it('keeps structural reference errors dominant over induced self cycles', async () => {
    const engine = await createSeededEngine('structural-ref-error-cycle-dominance-regression')

    engine.setCellFormula('Sheet1', 'B1', 'A1+B1')
    engine.deleteColumns('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'A1').formula).toBe('#REF!+A1')
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('recalculates direct scalar formulas when structural rewrites change dependency cells', async () => {
    const engine = await createSeededEngine('direct-scalar-structural-dependency-change-regression')

    engine.deleteRows('Sheet1', 1, 1)
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')
    engine.insertRows('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'B7').formula).toBe('A3+B3')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 43 })
  })

  it('replays direct scalar insert rewrites through redo history', async () => {
    const engine = await createSeededEngine('direct-scalar-structural-insert-redo-regression')

    engine.deleteRows('Sheet1', 1, 1)
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')
    engine.insertRows('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'B7').formula).toBe('A3+B3')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 43 })

    expect(engine.undo()).toBe(true)
    expect(engine.getCell('Sheet1', 'B6').formula).toBe('A1+B1')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })

    expect(engine.redo()).toBe(true)
    expect(engine.getCell('Sheet1', 'B7').formula).toBe('A3+B3')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 43 })
  })

  it('does not reuse stale template rewrites for copied formulas after undo', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'copied-formula-template-rewrite-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'copied-formula-template-rewrite-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertRows('Sheet1', 1, 2)
    engine.deleteRows('Sheet1', 0, 1)
    expect(engine.undo()).toBe(true)
    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C2' },
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
    )
    engine.deleteRows('Sheet1', 2, 2)

    expect(engine.getCell('Sheet1', 'C2').formula).toBe('A2+B3')
    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(
        await exportReplaySnapshot(seedSnapshot, [
          { kind: 'insertRows', start: 1, count: 2 },
          {
            kind: 'copy',
            source: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C2' },
            target: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
          },
          { kind: 'deleteRows', start: 2, count: 2 },
        ]),
      ),
    )
  })

  it('does not apply stale direct formula results after move overwrites a target formula', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'move-over-target-formula-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'move-over-target-formula-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D2' },
    )

    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCell('Sheet1', 'D2').formula).toBeUndefined()
  })

  it('keeps move-over-target-formula history aligned after structural undo', async () => {
    const seedSnapshot = await createEngineSeedSnapshot('formula-graph', 'move-over-target-formula-history-regression')
    const engine = new SpreadsheetEngine({
      workbookName: seedSnapshot.workbook.name,
      replicaId: 'move-over-target-formula-history-regression',
    })
    await engine.ready()
    engine.importSnapshot(structuredClone(seedSnapshot))

    engine.insertColumns('Sheet1', 0, 1)
    expect(engine.undo()).toBe(true)
    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D2' },
    )

    expect(normalizeSnapshotForSemanticComparison(engine.exportSnapshot())).toEqual(
      normalizeSnapshotForSemanticComparison(
        await exportReplaySnapshot(seedSnapshot, [
          {
            kind: 'move',
            source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
            target: { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D2' },
          },
        ]),
      ),
    )
  })

  it('keeps structural reference errors dominant after undoing column moves', async () => {
    const engine = await createSeededEngine('structural-column-move-undo-cycle-regression')

    engine.setCellFormula('Sheet1', 'A1', 'A1+B1')
    expect(engine.undo()).toBe(true)
    expect(engine.redo()).toBe(true)

    engine.insertColumns('Sheet1', 0, 1)
    expect(engine.undo()).toBe(true)
    expect(engine.redo()).toBe(true)

    engine.deleteColumns('Sheet1', 2, 1)
    expect(engine.undo()).toBe(true)
    expect(engine.redo()).toBe(true)

    expect(engine.getCell('Sheet1', 'B6').formula).toBe('SUM(B1:B5)')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })

    engine.moveColumns('Sheet1', 0, 1, 1)
    expect(engine.undo()).toBe(true)

    expect(engine.getCell('Sheet1', 'B6').formula).toBe('SUM(B1:B5)')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })
})

async function createSeededEngine(workbookName: string): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName })
  await engine.ready()
  engine.createSheet('Sheet1')
  seedInitialGrid(engine)
  return engine
}

function seedInitialGrid(engine: SpreadsheetEngine): void {
  for (let row = 0; row < 5; row += 1) {
    for (let column = 0; column < 5; column += 1) {
      engine.setCellValue('Sheet1', formatAddress(row, column), row * 10 + column + 1)
    }
  }
  engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
  engine.setCellFormula('Sheet1', 'B6', 'A1+B1')
}

function formatAddress(row: number, column: number): string {
  let dividend = column + 1
  let columnName = ''
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26
    columnName = String.fromCharCode(65 + modulo) + columnName
    dividend = Math.floor((dividend - modulo) / 26)
  }
  return `${columnName}${row + 1}`
}
