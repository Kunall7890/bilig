import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  WorkPaper,
  type RawCellContent,
} from '../index.js'

const sheetName = 'Sheet1'
const rowCountToCompare = 7
const colCountToCompare = 7

type ParityAction =
  | { readonly kind: 'set-cell'; readonly row: number; readonly col: number; readonly value: RawCellContent }
  | { readonly kind: 'clear-cell'; readonly row: number; readonly col: number }
  | { readonly kind: 'insert-rows'; readonly start: number; readonly count: number }
  | { readonly kind: 'delete-rows'; readonly start: number; readonly count: number }
  | { readonly kind: 'insert-columns'; readonly start: number; readonly count: number }
  | { readonly kind: 'delete-columns'; readonly start: number; readonly count: number }
  | { readonly kind: 'move-rows'; readonly start: number; readonly count: number; readonly target: number }
  | { readonly kind: 'move-columns'; readonly start: number; readonly count: number; readonly target: number }

interface ParityHarness {
  readonly workPaper: WorkPaper
  readonly engine: SpreadsheetEngine
  readonly sheetId: number
}

describe('WorkPaper/core parity fuzz', () => {
  it('replays row-move self-reference persistence counterexample', async () => {
    const harness = await createParityHarness()
    applyParityAction(harness, { kind: 'move-rows', start: 0, count: 1, target: 5 })
    assertCoreParity(harness)
    assertRestoredWorkPaperParity(harness)
    await assertRestoredEngineParity(harness)

    applyParityAction(harness, { kind: 'delete-rows', start: 0, count: 1 })
    assertCoreParity(harness)
    assertRestoredWorkPaperParity(harness)
    await assertRestoredEngineParity(harness)
  })

  it('recalculates expanded direct aggregates after row moves before save/load', async () => {
    const harness = await createParityHarness()
    applyParityAction(harness, { kind: 'set-cell', row: 0, col: 2, value: '=SUM(A1:B2)' })
    applyParityAction(harness, { kind: 'move-rows', start: 2, count: 1, target: 1 })

    assertCoreParity(harness)
    expect(harness.workPaper.getCellValue({ sheet: harness.sheetId, row: 0, col: 2 })).toEqual({ tag: ValueTag.Number, value: 69 })
    assertRestoredWorkPaperParity(harness)
    await assertRestoredEngineParity(harness)
  })

  it('recalculates dependents of precomputed aggregate values after row deletes', async () => {
    const harness = await createParityHarness()
    applyParityAction(harness, { kind: 'move-rows', start: 5, count: 1, target: 0 })
    applyParityAction(harness, { kind: 'set-cell', row: 0, col: 1, value: '=IF(A1>0,A1,0)' })
    applyParityAction(harness, { kind: 'delete-rows', start: 1, count: 1 })

    assertCoreParity(harness)
    expect(harness.workPaper.getCellValue({ sheet: harness.sheetId, row: 0, col: 1 })).toEqual({ tag: ValueTag.Number, value: 104 })
    assertRestoredWorkPaperParity(harness)
    await assertRestoredEngineParity(harness)
  })

  it('restores formula bindings when undoing column inserts', async () => {
    const harness = await createParityHarness()
    const beforeParity = readParityDigest(harness)
    applyParityAction(harness, { kind: 'set-cell', row: 0, col: 2, value: '=A1&"-"&B1' })
    const afterFormula = readParityDigest(harness)
    applyParityAction(harness, { kind: 'insert-columns', start: 0, count: 1 })

    harness.workPaper.undo()
    expect(harness.engine.undo()).toBe(true)
    expect(readParityDigest(harness)).toEqual(afterFormula)
    assertCoreParity(harness)

    harness.workPaper.undo()
    expect(harness.engine.undo()).toBe(true)
    expect(readParityDigest(harness)).toEqual(beforeParity)
    assertCoreParity(harness)
  })

  it('keeps headless WorkPaper edits, structural operations, undo/redo, and save/load in core parity', async () => {
    await runProperty({
      suite: 'headless/core-parity/work-paper-engine-action-sequence',
      arbitrary: fc.array(parityActionArbitrary, { minLength: 1, maxLength: 24 }),
      predicate: async (actions) => {
        const harness = await createParityHarness()
        assertCoreParity(harness)

        await actions.reduce(async (previous, action) => {
          await previous
          const beforeParity = readParityDigest(harness)
          applyParityAction(harness, action)
          assertCoreParity(harness)
          assertRestoredWorkPaperParity(harness)
          await assertRestoredEngineParity(harness)

          const afterParity = readParityDigest(harness)
          if (!sameJson(beforeParity, afterParity)) {
            harness.workPaper.undo()
            expect(harness.engine.undo()).toBe(true)
            expect(readParityDigest(harness)).toEqual(beforeParity)
            assertCoreParity(harness)

            harness.workPaper.redo()
            expect(harness.engine.redo()).toBe(true)
            expect(readParityDigest(harness)).toEqual(afterParity)
            assertCoreParity(harness)
          }
        }, Promise.resolve())
      },
      parameters: { numRuns: 160, interruptAfterTimeLimit: 120_000 },
    })
  }, 300_000)
})

const literalArbitrary: fc.Arbitrary<LiteralInput> = fc.oneof(
  fc.constant(null),
  fc.boolean(),
  fc.integer({ min: -1_000, max: 1_000 }),
  fc.constantFrom('north', 'south', 'total', 'line item', ''),
)

const rawCellContentArbitrary: fc.Arbitrary<RawCellContent> = fc.oneof(
  literalArbitrary,
  fc.constantFrom('=A1+B1', '=SUM(A1:B2)', '=IF(A1>0,A1,0)', '=A1&"-"&B1'),
)

const parityActionArbitrary: fc.Arbitrary<ParityAction> = fc.oneof(
  fc
    .record({
      row: fc.integer({ min: 0, max: 6 }),
      col: fc.integer({ min: 0, max: 6 }),
      value: rawCellContentArbitrary,
    })
    .map(({ row, col, value }) => ({ kind: 'set-cell', row, col, value })),
  fc
    .record({
      row: fc.integer({ min: 0, max: 6 }),
      col: fc.integer({ min: 0, max: 6 }),
    })
    .map(({ row, col }) => ({ kind: 'clear-cell', row, col })),
  axisEditArbitrary('insert-rows'),
  axisEditArbitrary('delete-rows'),
  axisEditArbitrary('insert-columns'),
  axisEditArbitrary('delete-columns'),
  axisMoveArbitrary('move-rows'),
  axisMoveArbitrary('move-columns'),
)

function axisEditArbitrary(
  kind: Extract<ParityAction['kind'], 'insert-rows' | 'delete-rows' | 'insert-columns' | 'delete-columns'>,
): fc.Arbitrary<ParityAction> {
  return fc
    .record({
      start: fc.integer({ min: 0, max: 5 }),
      count: fc.integer({ min: 1, max: 2 }),
    })
    .map(({ start, count }) => ({ kind, start, count }))
}

function axisMoveArbitrary(kind: Extract<ParityAction['kind'], 'move-rows' | 'move-columns'>): fc.Arbitrary<ParityAction> {
  return fc
    .record({
      start: fc.integer({ min: 0, max: 5 }),
      count: fc.constant(1),
      target: fc.integer({ min: 0, max: 6 }),
    })
    .map(({ start, count, target }) => ({ kind, start, count, target }))
}

async function createParityHarness(): Promise<ParityHarness> {
  const workPaper = WorkPaper.buildEmpty({
    parseDateTime: () => undefined,
    functionPlugins: [],
  })
  workPaper.addSheet(sheetName)
  const sheetId = workPaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error('Expected WorkPaper sheet to exist')
  }

  const engine = new SpreadsheetEngine({
    workbookName: 'headless-core-parity',
    replicaId: 'headless-core-parity',
    trackReplicaVersions: false,
  })
  await engine.ready()
  engine.createSheet(sheetName)

  const harness = { workPaper, engine, sheetId }
  seedInitialGrid(harness)
  return harness
}

function seedInitialGrid(harness: ParityHarness): void {
  for (let row = 0; row < 5; row += 1) {
    for (let col = 0; col < 5; col += 1) {
      const value = row * 10 + col + 1
      setCellOnBoth(harness, row, col, value)
    }
  }
  setCellOnBoth(harness, 5, 0, '=SUM(A1:A5)')
  setCellOnBoth(harness, 5, 1, '=A1+B1')
}

function applyParityAction(harness: ParityHarness, action: ParityAction): void {
  switch (action.kind) {
    case 'set-cell':
      setCellOnBoth(harness, action.row, action.col, action.value)
      return
    case 'clear-cell':
      setCellOnBoth(harness, action.row, action.col, null)
      return
    case 'insert-rows':
      harness.workPaper.addRows(harness.sheetId, action.start, action.count)
      harness.engine.insertRows(sheetName, action.start, action.count)
      return
    case 'delete-rows':
      harness.workPaper.removeRows(harness.sheetId, action.start, action.count)
      harness.engine.deleteRows(sheetName, action.start, action.count)
      return
    case 'insert-columns':
      harness.workPaper.addColumns(harness.sheetId, action.start, action.count)
      harness.engine.insertColumns(sheetName, action.start, action.count)
      return
    case 'delete-columns':
      harness.workPaper.removeColumns(harness.sheetId, action.start, action.count)
      harness.engine.deleteColumns(sheetName, action.start, action.count)
      return
    case 'move-rows':
      harness.workPaper.moveRows(harness.sheetId, action.start, action.count, action.target)
      harness.engine.moveRows(sheetName, action.start, action.count, action.target)
      return
    case 'move-columns':
      harness.workPaper.moveColumns(harness.sheetId, action.start, action.count, action.target)
      harness.engine.moveColumns(sheetName, action.start, action.count, action.target)
      return
    default:
      assertNever(action)
  }
}

function setCellOnBoth(harness: ParityHarness, row: number, col: number, value: RawCellContent): void {
  harness.workPaper.setCellContents({ sheet: harness.sheetId, row, col }, value)
  const address = formatAddress(row, col)
  if (typeof value === 'string' && value.startsWith('=')) {
    harness.engine.setCellFormula(sheetName, address, value.slice(1))
    return
  }
  harness.engine.setCellValue(sheetName, address, value)
}

function assertCoreParity(harness: ParityHarness): void {
  expect(readParityDigest(harness)).toEqual(readEngineDigest(harness.engine))
}

function assertRestoredWorkPaperParity(harness: ParityHarness): void {
  const document = exportWorkPaperDocument(harness.workPaper, { includeConfig: true })
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkPaperDocument(document)))
  const restoredSheetId = restored.getSheetId(sheetName)
  if (restoredSheetId === undefined) {
    throw new Error('Expected restored WorkPaper sheet to exist')
  }
  expect(readWorkPaperDigest(restored, restoredSheetId)).toEqual(readParityDigest(harness))
}

async function assertRestoredEngineParity(harness: ParityHarness): Promise<void> {
  const snapshot = harness.engine.exportSnapshot()
  const restored = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: 'headless-core-parity-restored',
  })
  await restored.ready()
  restored.importSnapshot(snapshot)
  expect(readEngineDigest(restored)).toEqual(readParityDigest(harness))
}

function readParityDigest(harness: ParityHarness): unknown {
  return readWorkPaperDigest(harness.workPaper, harness.sheetId)
}

function readWorkPaperDigest(workPaper: WorkPaper, sheetId: number): unknown {
  const cells = []
  for (let row = 0; row < rowCountToCompare; row += 1) {
    for (let col = 0; col < colCountToCompare; col += 1) {
      const address = { sheet: sheetId, row, col }
      cells.push({
        address: formatAddress(row, col),
        formula: normalizeSerializedFormula(workPaper.getCellSerialized(address)),
        value: normalizeCellValue(workPaper.getCellValue(address)),
      })
    }
  }
  return cells
}

function readEngineDigest(engine: SpreadsheetEngine): unknown {
  const cells = []
  for (let row = 0; row < rowCountToCompare; row += 1) {
    for (let col = 0; col < colCountToCompare; col += 1) {
      const address = formatAddress(row, col)
      cells.push({
        address,
        formula: normalizeSerializedFormula(readEngineSerialized(engine, address)),
        value: normalizeCellValue(engine.getCellValue(sheetName, address)),
      })
    }
  }
  return cells
}

function readEngineSerialized(engine: SpreadsheetEngine, address: string): RawCellContent {
  const cell = engine.getCell(sheetName, address)
  if (cell.formula) {
    return `=${cell.formula}`
  }
  return cell.input ?? null
}

function normalizeSerializedFormula(value: RawCellContent): string | null {
  return typeof value === 'string' && value.startsWith('=') ? value : null
}

function normalizeCellValue(value: CellValue): unknown {
  switch (value.tag) {
    case ValueTag.Empty:
      return { tag: value.tag }
    case ValueTag.Number:
      return { tag: value.tag, value: Object.is(value.value, -0) ? 0 : value.value }
    case ValueTag.Boolean:
      return { tag: value.tag, value: value.value }
    case ValueTag.String:
      return { tag: value.tag, value: value.value }
    case ValueTag.Error:
      return { tag: value.tag, code: value.code }
    default:
      assertNever(value)
  }
}

function formatAddress(row: number, col: number): string {
  let current = col + 1
  let label = ''
  while (current > 0) {
    current -= 1
    label = String.fromCharCode(65 + (current % 26)) + label
    current = Math.floor(current / 26)
  }
  return `${label}${row + 1}`
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${JSON.stringify(value)}`)
}
