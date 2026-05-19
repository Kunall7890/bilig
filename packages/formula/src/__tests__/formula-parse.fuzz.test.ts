import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { formulaCompatibilityRegistry } from '../compatibility.js'
import { formulaInventory } from '../generated/formula-inventory.js'
import { parseFormula } from '../parser.js'
import { serializeFormula } from '../translation.js'
import { invalidFormulaArbitrary, validFormulaArbitrary } from './formula-fuzz-helpers.js'
import { runProperty } from '@bilig/test-fuzz'

const inventoryFunctionNameArbitrary = fc.constantFrom(
  ...formulaInventory.filter((entry) => entry.registeredInCodebase && entry.runtimeStatus === 'implemented').map((entry) => entry.name),
)

const inventoryFormulaArbitrary = fc.oneof(
  fc.constantFrom(...formulaCompatibilityRegistry.map((entry) => entry.formula)),
  inventoryFunctionNameArbitrary.chain((name) =>
    fc.constantFrom(
      `=${name}($A$1,'My Sheet'!$B2,Data!C$3,#N/A,TRUE,"alpha",,)`,
      `=${name}('My Sheet'!$A$1:$B$3,Sheet1!$1:$3,Data!$C:$D,TaxRate)`,
      `=${name}({1,TRUE,#N/A;"",FALSE,3},'My Sheet'!$A$1#,Sales[Amount])`,
    ),
  ),
)

describe('formula parse fuzz', () => {
  it('canonicalizes valid formulas through parse and serialize', async () => {
    await runProperty({
      suite: 'formula/parse/canonicalization',
      arbitrary: validFormulaArbitrary,
      predicate: (formula) => {
        const canonical = serializeFormula(parseFormula(formula))
        expect(serializeFormula(parseFormula(canonical))).toBe(canonical)
      },
    })
  })

  it('canonicalizes inventory-derived formulas with rich reference syntax', async () => {
    await runProperty({
      suite: 'formula/parse/inventory-rich-syntax',
      arbitrary: inventoryFormulaArbitrary,
      parameters: { numRuns: 180 },
      predicate: (formula) => {
        const canonical = serializeFormula(parseFormula(formula))
        expect(serializeFormula(parseFormula(canonical))).toBe(canonical)
      },
    })
  })

  it('rejects malformed formulas without crashing the parser', async () => {
    await runProperty({
      suite: 'formula/parse/invalid-input',
      arbitrary: invalidFormulaArbitrary,
      predicate: (formula) => {
        expect(() => parseFormula(formula)).toThrow(/.+/)
      },
    })
  })
})
