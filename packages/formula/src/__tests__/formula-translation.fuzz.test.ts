import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { formulaInventory } from '../generated/formula-inventory.js'
import { parseFormula } from '../parser.js'
import { rewriteFormulaForStructuralTransform, serializeFormula, translateFormulaReferences } from '../translation.js'
import { validFormulaArbitrary } from './formula-fuzz-helpers.js'
import { runProperty } from '@bilig/test-fuzz'

const inventoryFunctionNameArbitrary = fc.constantFrom(
  ...formulaInventory.filter((entry) => entry.registeredInCodebase && entry.protocolSupportsWasm).map((entry) => entry.name),
)

const richSheetNameArbitrary = fc.constantFrom('Input Sheet', "O'Brien Data", 'Data.2026')

const inventoryRichReferenceFormulaArbitrary = fc.record({
  firstName: inventoryFunctionNameArbitrary,
  secondName: inventoryFunctionNameArbitrary,
  targetSheetName: richSheetNameArbitrary,
  otherSheetName: richSheetNameArbitrary,
})

describe('formula translation fuzz', () => {
  it('reverses translated references back to the canonical formula', async () => {
    await runProperty({
      suite: 'formula/translation/reference-reversal',
      arbitrary: fc
        .record({
          formula: validFormulaArbitrary,
          rowDelta: fc.integer({ min: 0, max: 4 }),
          colDelta: fc.integer({ min: 0, max: 2 }),
        })
        .filter((value) => value.rowDelta !== 0 || value.colDelta !== 0),
      predicate: ({ formula, rowDelta, colDelta }) => {
        const canonical = serializeFormula(parseFormula(formula))
        const translated = translateFormulaReferences(canonical, rowDelta, colDelta)
        const restored = translateFormulaReferences(translated, -rowDelta, -colDelta)
        expect(restored).toBe(canonical)
      },
    })
  })

  it('reverses inventory-derived translated references with quoted sheets and absolute anchors', async () => {
    await runProperty({
      suite: 'formula/translation/inventory-rich-reference-reversal',
      arbitrary: inventoryRichReferenceFormulaArbitrary,
      parameters: { numRuns: 140 },
      predicate: ({ firstName, secondName, targetSheetName, otherSheetName }) => {
        const formula = `${firstName}(${quoteSheetName(targetSheetName)}!$A$1,${quoteSheetName(
          targetSheetName,
        )}!B$2:C3,${quoteSheetName(otherSheetName)}!$D:$E,TaxRate,#REF!,TRUE,"x")+${secondName}($A1,B$2)`
        const canonical = serializeFormula(parseFormula(formula))
        const translated = translateFormulaReferences(canonical, 3, 2)
        const restored = translateFormulaReferences(translated, -3, -2)
        expect(restored).toBe(canonical)
      },
    })
  })

  it('reverses insert transforms through matching delete transforms', async () => {
    await runProperty({
      suite: 'formula/translation/structural-insert-delete-reversal',
      arbitrary: fc.record({
        formula: validFormulaArbitrary,
        axis: fc.constantFrom<'row' | 'column'>('row', 'column'),
        start: fc.integer({ min: 0, max: 4 }),
        count: fc.integer({ min: 1, max: 2 }),
      }),
      predicate: ({ formula, axis, start, count }) => {
        const canonical = serializeFormula(parseFormula(formula))
        const inserted = rewriteFormulaForStructuralTransform(canonical, 'Sheet1', 'Sheet1', {
          kind: 'insert',
          axis,
          start,
          count,
        })
        const restored = rewriteFormulaForStructuralTransform(inserted, 'Sheet1', 'Sheet1', {
          kind: 'delete',
          axis,
          start,
          count,
        })
        expect(restored).toBe(canonical)
      },
    })
  })

  it('reverses inventory-derived structural transforms on targeted quoted sheets only', async () => {
    await runProperty({
      suite: 'formula/translation/inventory-structural-quoted-sheet-reversal',
      arbitrary: inventoryRichReferenceFormulaArbitrary
        .filter((value) => value.targetSheetName !== value.otherSheetName)
        .chain((value) =>
          fc.record({
            firstName: fc.constant(value.firstName),
            secondName: fc.constant(value.secondName),
            targetSheetName: fc.constant(value.targetSheetName),
            otherSheetName: fc.constant(value.otherSheetName),
            axis: fc.constantFrom<'row' | 'column'>('row', 'column'),
            start: fc.integer({ min: 0, max: 2 }),
            count: fc.constantFrom(1, 2),
          }),
        ),
      parameters: { numRuns: 120 },
      predicate: ({ firstName, secondName, targetSheetName, otherSheetName, axis, start, count }) => {
        const targetPrefix = `${quoteSheetName(targetSheetName)}!`
        const otherPrefix = `${quoteSheetName(otherSheetName)}!`
        const formula = `${firstName}(${targetPrefix}$A$1:$C$5,${targetPrefix}1:4,${targetPrefix}$A:$C,${otherPrefix}A1:B2,TaxRate)+${secondName}(${otherPrefix}$D$4,#N/A)`
        const canonical = serializeFormula(parseFormula(formula))
        const inserted = rewriteFormulaForStructuralTransform(canonical, 'Sheet1', targetSheetName, {
          kind: 'insert',
          axis,
          start,
          count,
        })
        const restored = rewriteFormulaForStructuralTransform(inserted, 'Sheet1', targetSheetName, {
          kind: 'delete',
          axis,
          start,
          count,
        })
        expect(restored).toBe(canonical)
      },
    })
  })
})

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}
