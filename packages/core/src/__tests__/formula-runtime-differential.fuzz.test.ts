import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { ErrorCode, FormulaMode, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import { compileFormula, formulaCompatibilityRegistry } from '../../../formula/src/index.js'
import { SpreadsheetEngine } from '../engine.js'
import { fastPathFormulaArbitrary } from '../../../formula/src/__tests__/formula-fuzz-helpers.js'
import { runProperty } from '@bilig/test-fuzz'

const sheetName = 'Sheet1'
const lookupKeys = ['apple', 'banana', 'pear', 'plum', 'quince', 'rice'] as const
const lookupFormulaKindArbitrary = fc.constantFrom('VLOOKUP', 'HLOOKUP', 'XLOOKUP' as const)
const lookupReturnValueArbitrary = fc.oneof(fc.integer({ min: -100, max: 100 }), fc.constant(null))
const volatileInventoryFormulaPattern = /\b(?:NOW|RAND|TODAY)\s*\(/iu
const FUZZ_TEST_TIMEOUT_MS = 30_000

const inventoryWasmFormulaCases = formulaCompatibilityRegistry
  .filter(
    (entry) =>
      entry.status === 'implemented-wasm-production' &&
      entry.wasmStatus === 'production' &&
      !volatileInventoryFormulaPattern.test(entry.formula) &&
      compileFormula(entry.formula).mode === FormulaMode.WasmFastPath,
  )
  .map((entry) => ({ id: entry.id, formula: entry.formula }))

const inventoryWasmFormulaCaseArbitrary = fc.constantFrom(...inventoryWasmFormulaCases)
const runtimeLiteralArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -20, max: 20 }),
  fc.boolean(),
  fc.constantFrom('', 'x', 'pear', '42'),
  fc.constant(null),
)
const smallNumberArbitrary = fc.integer({ min: -20, max: 20 })
const positiveNumberArbitrary = fc.integer({ min: 1, max: 20 })
const scalarDefinedNameFormulaCaseArbitrary = fc.constantFrom<ScalarDefinedNameFormulaCase>(
  { formula: 'TaxRate*$A$1', updatedName: 'TaxRate' },
  { formula: 'taxrate*$A$1+FeeRate', updatedName: 'FeeRate' },
  { formula: 'MissingRate*$A$1', updatedName: 'MissingRate' },
)

interface LookupFormulaCase {
  readonly formula: string
  readonly range: { startAddress: string; endAddress: string }
  readonly values: LiteralInput[][]
  readonly expected: CellValue
}

interface ScalarDefinedNameFormulaCase {
  readonly formula: string
  readonly updatedName: 'TaxRate' | 'FeeRate' | 'MissingRate'
}

const fastPathLookupFormulaCaseArbitrary = fc
  .uniqueArray(fc.constantFrom(...lookupKeys), { minLength: 2, maxLength: 5 })
  .chain((keys) =>
    fc.record({
      kind: lookupFormulaKindArbitrary,
      keys: fc.constant(keys),
      returns: fc.array(lookupReturnValueArbitrary, {
        minLength: keys.length,
        maxLength: keys.length,
      }),
      query: fc.oneof(fc.constantFrom(...keys), fc.constant('missing')),
    }),
  )
  .map(({ kind, keys, returns, query }) => buildLookupFormulaCase(kind, keys, returns, query))

function buildNumericGrid(rows: number, cols: number): number[][] {
  return Array.from({ length: rows }, (_rowValue, row) => Array.from({ length: cols }, (_colValue, col) => row * cols + col + 1))
}

describe('formula runtime differential fuzz', () => {
  it(
    'keeps GROUPBY and PIVOTBY WASM header labels in JS text parity',
    async () => {
      const engine = new SpreadsheetEngine({
        workbookName: 'grouped-array-header-label-parity',
        replicaId: 'grouped-array-header-label-parity',
      })
      await engine.ready()
      seedInventoryDifferentialWorkbook(engine)

      engine.setCellFormula(sheetName, 'H1', 'GROUPBY(A1:A5,C1:C5,SUM,3,1)')
      engine.setCellFormula(sheetName, 'H8', 'PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)')

      expect(engine.explainCell(sheetName, 'H1').mode).toBe(FormulaMode.WasmFastPath)
      expect(engine.explainCell(sheetName, 'H8').mode).toBe(FormulaMode.WasmFastPath)
      expect(engine.recalculateDifferential().drift).toEqual([])

      expectCellSemantics(engine.getCellValue(sheetName, 'H1'), stringValue('1'))
      expectCellSemantics(engine.getCellValue(sheetName, 'I1'), stringValue('3'))
      expectCellSemantics(engine.getCellValue(sheetName, 'H8'), stringValue('1'))
      expectCellSemantics(engine.getCellValue(sheetName, 'I8'), stringValue(''))
      expectCellSemantics(engine.getCellValue(sheetName, 'J8'), stringValue('4'))
      expectCellSemantics(engine.getCellValue(sheetName, 'K8'), stringValue('5'))
      expectCellSemantics(engine.getCellValue(sheetName, 'L8'), stringValue('6'))

      engine.setCellValue(sheetName, 'A1', null)
      expect(engine.recalculateDifferential().drift).toEqual([])
      expectCellSemantics(engine.getCellValue(sheetName, 'H1'), stringValue('Row Field 1'))
      expectCellSemantics(engine.getCellValue(sheetName, 'H8'), stringValue('Row Field 1'))

      engine.setCellValue(sheetName, 'A1', '')
      expect(engine.recalculateDifferential().drift).toEqual([])
      expectCellSemantics(engine.getCellValue(sheetName, 'H1'), stringValue('Row Field 1'))
      expectCellSemantics(engine.getCellValue(sheetName, 'H8'), stringValue('Row Field 1'))
    },
    FUZZ_TEST_TIMEOUT_MS,
  )

  it(
    'keeps statistical WASM range booleans in JS parity',
    async () => {
      const engine = new SpreadsheetEngine({
        workbookName: 'statistical-boolean-range-parity',
        replicaId: 'statistical-boolean-range-parity',
      })
      await engine.ready()
      seedInventoryDifferentialWorkbook(engine)

      engine.setCellFormula(sheetName, 'H8', 'STDEV(A1:A4)')
      engine.setCellFormula(sheetName, 'I8', 'VAR(A1:A4)')

      expect(engine.explainCell(sheetName, 'H8').mode).toBe(FormulaMode.WasmFastPath)
      expect(engine.explainCell(sheetName, 'I8').mode).toBe(FormulaMode.WasmFastPath)
      expect(engine.recalculateDifferential().drift).toEqual([])

      engine.setCellValue(sheetName, 'A1', false)

      expect(engine.recalculateDifferential().drift).toEqual([])
      expectNumberCellClose(engine.getCellValue(sheetName, 'H8'), 1.707825127659933)
      expectNumberCellClose(engine.getCellValue(sheetName, 'I8'), 2.9166666666666665)
    },
    FUZZ_TEST_TIMEOUT_MS,
  )

  it(
    'keeps generated fast-path formulas in JS and wasm parity',
    async () => {
      await runProperty({
        suite: 'core/formula-runtime/generated-differential',
        arbitrary: fastPathFormulaArbitrary,
        predicate: async (formula) => {
          const engine = new SpreadsheetEngine({
            workbookName: `fuzz-formula-diff-${formula.length}`,
            replicaId: 'fuzz-formula-diff',
          })
          await engine.ready()
          engine.createSheet(sheetName)
          engine.setRangeValues({ sheetName, startAddress: 'A1', endAddress: 'F6' }, buildNumericGrid(6, 6))
          engine.setCellFormula(sheetName, 'G1', formula)

          const explanation = engine.explainCell(sheetName, 'G1')
          expect(explanation.mode).toBe(FormulaMode.WasmFastPath)

          const differential = engine.recalculateDifferential()
          expect(differential.drift).toEqual([])

          const snapshot = engine.exportSnapshot()
          const restored = new SpreadsheetEngine({
            workbookName: snapshot.workbook.name,
            replicaId: 'fuzz-formula-diff-restored',
          })
          await restored.ready()
          restored.importSnapshot(snapshot)

          expect(restored.getCellValue(sheetName, 'G1')).toEqual(engine.getCellValue(sheetName, 'G1'))
        },
      })
    },
    FUZZ_TEST_TIMEOUT_MS,
  )

  it(
    'keeps inventory-declared WASM formulas in JS and wasm parity on mixed sparse inputs',
    async () => {
      await runProperty({
        suite: 'core/formula-runtime/inventory-wasm-differential',
        arbitrary: fc.record({
          formulaCase: inventoryWasmFormulaCaseArbitrary,
          updateValue: runtimeLiteralArbitrary,
        }),
        parameters: { numRuns: 90, interruptAfterTimeLimit: 20_000 },
        predicate: async ({ formulaCase, updateValue }) => {
          const engine = new SpreadsheetEngine({
            workbookName: `fuzz-inventory-diff-${formulaCase.id}`,
            replicaId: 'fuzz-inventory-diff',
          })
          await engine.ready()
          seedInventoryDifferentialWorkbook(engine)
          engine.setCellFormula(sheetName, 'H8', formulaCase.formula)

          expect(engine.recalculateDifferential().drift).toEqual([])

          engine.setCellValue(sheetName, 'A1', updateValue)
          expect(engine.recalculateDifferential().drift).toEqual([])
        },
      })
    },
    FUZZ_TEST_TIMEOUT_MS,
  )

  it(
    'keeps exact lookup fast-path formulas in JS and wasm parity',
    async () => {
      await runProperty({
        suite: 'core/formula-runtime/generated-lookup-differential',
        arbitrary: fastPathLookupFormulaCaseArbitrary,
        parameters: { numRuns: 75, interruptAfterTimeLimit: 15_000 },
        predicate: async (lookupCase) => {
          const engine = new SpreadsheetEngine({
            workbookName: `fuzz-lookup-diff-${lookupCase.formula.length}`,
            replicaId: 'fuzz-lookup-diff',
          })
          await engine.ready()
          engine.createSheet(sheetName)
          engine.setRangeValues({ sheetName, ...lookupCase.range }, lookupCase.values)
          engine.setCellFormula(sheetName, 'G1', lookupCase.formula)

          expect(engine.explainCell(sheetName, 'G1').mode).toBe(FormulaMode.WasmFastPath)
          expectCellSemantics(engine.getCellValue(sheetName, 'G1'), lookupCase.expected)

          const differential = engine.recalculateDifferential()
          expect(differential.drift).toEqual([])

          const snapshot = engine.exportSnapshot()
          const restored = new SpreadsheetEngine({
            workbookName: snapshot.workbook.name,
            replicaId: 'fuzz-lookup-diff-restored',
          })
          await restored.ready()
          restored.importSnapshot(snapshot)

          expectCellSemantics(restored.getCellValue(sheetName, 'G1'), lookupCase.expected)
        },
      })
    },
    FUZZ_TEST_TIMEOUT_MS,
  )

  it(
    'keeps scalar defined-name formulas differential-clean across dependency invalidation',
    async () => {
      await runProperty({
        suite: 'core/formula-runtime/scalar-defined-name-invalidation',
        arbitrary: fc.record({
          formulaCase: scalarDefinedNameFormulaCaseArbitrary,
          sourceValue: smallNumberArbitrary,
          nextSourceValue: smallNumberArbitrary,
          taxRate: smallNumberArbitrary,
          feeRate: smallNumberArbitrary,
          nextNameValue: smallNumberArbitrary,
        }),
        parameters: { numRuns: 90, interruptAfterTimeLimit: 20_000 },
        predicate: async ({ formulaCase, sourceValue, nextSourceValue, taxRate, feeRate, nextNameValue }) => {
          const engine = new SpreadsheetEngine({
            workbookName: `fuzz-defined-name-diff-${formulaCase.updatedName}`,
            replicaId: 'fuzz-defined-name-diff',
          })
          await engine.ready()
          engine.createSheet(sheetName)
          engine.setCellValue(sheetName, 'A1', sourceValue)
          engine.setDefinedName('TaxRate', { kind: 'scalar', value: taxRate })
          engine.setDefinedName('FeeRate', { kind: 'scalar', value: feeRate })
          engine.setCellFormula(sheetName, 'G2', formulaCase.formula)

          expect(engine.explainCell(sheetName, 'G2').mode).toBe(FormulaMode.WasmFastPath)
          expect(engine.recalculateDifferential().drift).toEqual([])

          engine.setCellValue(sheetName, 'A1', nextSourceValue)
          expect(engine.recalculateDifferential().drift).toEqual([])

          engine.setDefinedName(formulaCase.updatedName, { kind: 'scalar', value: nextNameValue })
          expectCellSemantics(
            engine.getCellValue(sheetName, 'G2'),
            expectedDefinedNameValue(formulaCase, nextSourceValue, taxRate, feeRate, nextNameValue),
          )
          expect(engine.recalculateDifferential().drift).toEqual([])
        },
      })
    },
    FUZZ_TEST_TIMEOUT_MS,
  )

  it(
    'keeps quoted-sheet structural rewrites differential-clean on sparse WASM ranges',
    async () => {
      await runProperty({
        suite: 'core/formula-runtime/quoted-sheet-structural-differential',
        arbitrary: fc.record({
          insertAt: fc.integer({ min: 0, max: 3 }),
          firstValue: positiveNumberArbitrary,
          middleValue: positiveNumberArbitrary,
          tailValue: positiveNumberArbitrary,
          insertedValue: positiveNumberArbitrary,
          flag: fc.boolean(),
        }),
        parameters: { numRuns: 90, interruptAfterTimeLimit: 20_000 },
        predicate: async ({ insertAt, firstValue, middleValue, tailValue, insertedValue, flag }) => {
          const engine = new SpreadsheetEngine({
            workbookName: 'fuzz-quoted-structural-diff',
            replicaId: 'fuzz-quoted-structural-diff',
          })
          await engine.ready()
          engine.createSheet('Input Sheet')
          engine.createSheet('Summary')
          engine.setCellValue('Input Sheet', 'A1', firstValue)
          engine.setCellValue('Input Sheet', 'A3', middleValue)
          engine.setCellValue('Input Sheet', 'A6', tailValue)
          engine.setCellValue('Input Sheet', 'B5', 'text')
          engine.setCellValue('Input Sheet', 'C1', flag)
          engine.setCellFormula('Input Sheet', 'D1', '1/0')
          engine.setCellFormula(
            'Summary',
            'G1',
            "SUM('Input Sheet'!$A$1:$A$6)+COUNTBLANK('Input Sheet'!B1:B6)+N('Input Sheet'!$C$1)+IFERROR('Input Sheet'!D1,0)",
          )

          expect(engine.explainCell('Summary', 'G1').mode).toBe(FormulaMode.WasmFastPath)
          expect(engine.recalculateDifferential().drift).toEqual([])

          engine.insertRows('Input Sheet', insertAt, 1)
          engine.setCellValue('Input Sheet', `A${insertAt + 1}`, insertedValue)
          expect(engine.recalculateDifferential().drift).toEqual([])

          engine.deleteRows('Input Sheet', insertAt, 1)
          expect(engine.recalculateDifferential().drift).toEqual([])
        },
      })
    },
    FUZZ_TEST_TIMEOUT_MS,
  )
})

function expectCellSemantics(actual: CellValue, expected: CellValue): void {
  if (expected.tag === ValueTag.String) {
    expect(actual).toMatchObject({ tag: ValueTag.String, value: expected.value })
    return
  }
  expect(actual).toEqual(expected)
}

function expectNumberCellClose(actual: CellValue, expected: number): void {
  expect(actual.tag).toBe(ValueTag.Number)
  if (actual.tag !== ValueTag.Number) {
    return
  }
  expect(actual.value).toBeCloseTo(expected, 12)
}

function buildLookupFormulaCase(
  kind: 'VLOOKUP' | 'HLOOKUP' | 'XLOOKUP',
  keys: readonly (typeof lookupKeys)[number][],
  returns: readonly (number | null)[],
  query: (typeof lookupKeys)[number] | 'missing',
): LookupFormulaCase {
  const matchIndex = keys.findIndex((key) => key === query)
  const expected = matchIndex === -1 ? lookupMissValue(kind) : lookupReturnValue(returns[matchIndex] ?? null)
  const quotedQuery = quoteFormulaString(query)
  if (kind === 'HLOOKUP') {
    const endAddress = `${columnName(keys.length - 1)}2`
    const keyRow: LiteralInput[] = [...keys]
    return {
      formula: `HLOOKUP(${quotedQuery},A1:${endAddress},2,FALSE)`,
      range: { startAddress: 'A1', endAddress },
      values: [keyRow, [...returns]],
      expected,
    }
  }
  const endAddress = `B${keys.length}`
  const values = keys.map((key, index) => [key, returns[index] ?? null])
  return {
    formula:
      kind === 'VLOOKUP'
        ? `VLOOKUP(${quotedQuery},A1:${endAddress},2,FALSE)`
        : `XLOOKUP(${quotedQuery},A1:A${keys.length},B1:B${keys.length},"missing")`,
    range: { startAddress: 'A1', endAddress },
    values,
    expected,
  }
}

function lookupMissValue(kind: 'VLOOKUP' | 'HLOOKUP' | 'XLOOKUP'): CellValue {
  return kind === 'XLOOKUP' ? stringValue('missing') : { tag: ValueTag.Error, code: ErrorCode.NA }
}

function lookupReturnValue(value: number | null): CellValue {
  return { tag: ValueTag.Number, value: value ?? 0 }
}

function expectedDefinedNameValue(
  formulaCase: ScalarDefinedNameFormulaCase,
  sourceValue: number,
  taxRate: number,
  feeRate: number,
  nextNameValue: number,
): CellValue {
  switch (formulaCase.updatedName) {
    case 'TaxRate':
      return { tag: ValueTag.Number, value: nextNameValue * sourceValue }
    case 'FeeRate':
      return { tag: ValueTag.Number, value: taxRate * sourceValue + nextNameValue }
    case 'MissingRate':
      return { tag: ValueTag.Number, value: nextNameValue * sourceValue }
  }
}

function seedInventoryDifferentialWorkbook(engine: SpreadsheetEngine): void {
  engine.createSheet(sheetName)
  engine.createSheet('Sheet2')
  engine.createSheet('Data')
  engine.createSheet('My Sheet')
  seedSheet(engine, sheetName)
  seedSheet(engine, 'Sheet2')
  seedSheet(engine, 'Data')
  seedSheet(engine, 'My Sheet')
}

function seedSheet(engine: SpreadsheetEngine, name: string): void {
  engine.setRangeValues({ sheetName: name, startAddress: 'A1', endAddress: 'F8' }, [
    [1, 2, 3, 4, 5, 6],
    [2, null, 4, true, 6, 7],
    [3, 4, 'x', 6, 7, 8],
    [4, 5, 6, 7, null, 9],
    [5, 6, 7, 8, 9, true],
    [6, 7, 8, 9, 10, 'tail'],
    [7, null, 9, 10, 11, 12],
    [8, 9, 10, 11, 12, 13],
  ])
}

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function quoteFormulaString(value: string): string {
  return `"${value.replace(/"/g, '""')}"`
}

function columnName(index: number): string {
  return String.fromCharCode('A'.charCodeAt(0) + index)
}
