import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { formulaInventory } from '../generated/formula-inventory.js'
import { parseFormula } from '../parser.js'
import { renameFormulaSheetReferences, serializeFormula } from '../translation.js'
import { renameScopedFormulaArbitrary, sheetNameArbitrary } from './formula-fuzz-helpers.js'
import { runProperty } from '@bilig/test-fuzz'

const inventoryFunctionNameArbitrary = fc.constantFrom(
  ...formulaInventory.filter((entry) => entry.registeredInCodebase && entry.protocolSupportsWasm).map((entry) => entry.name),
)

const richSheetNameArbitrary = fc.constantFrom('Old Sheet', "O'Brien Data", 'Data.2026', 'Archive Sheet')

describe('formula rename fuzz', () => {
  it('roundtrips quoted and unquoted sheet renames', async () => {
    await runProperty({
      suite: 'formula/rename/sheet-roundtrip',
      arbitrary: fc
        .record({
          oldSheetName: sheetNameArbitrary,
          newSheetName: sheetNameArbitrary,
        })
        .filter((value) => value.oldSheetName !== value.newSheetName)
        .chain(({ oldSheetName, newSheetName }) =>
          renameScopedFormulaArbitrary(oldSheetName, newSheetName).map((formula) => ({
            formula,
            oldSheetName,
            newSheetName,
          })),
        ),
      predicate: ({ formula, oldSheetName, newSheetName }) => {
        const canonical = serializeFormula(parseFormula(formula))
        const renamed = renameFormulaSheetReferences(canonical, oldSheetName, newSheetName)
        const restored = renameFormulaSheetReferences(renamed, newSheetName, oldSheetName)
        expect(restored).toBe(canonical)
      },
    })
  })

  it('roundtrips inventory formulas without renaming strings or defined names', async () => {
    await runProperty({
      suite: 'formula/rename/inventory-string-and-name-safety',
      arbitrary: fc
        .record({
          functionName: inventoryFunctionNameArbitrary,
          oldSheetName: richSheetNameArbitrary,
          newSheetName: richSheetNameArbitrary,
        })
        .filter((value) => value.oldSheetName !== value.newSheetName),
      parameters: { numRuns: 120 },
      predicate: ({ functionName, oldSheetName, newSheetName }) => {
        const oldPrefix = `${quoteSheetName(oldSheetName)}!`
        const formula = `${functionName}(${oldPrefix}$A$1,${oldPrefix}B2:C3,OldSheetTaxRate,"${escapeFormulaString(
          oldSheetName,
        )}")+${oldPrefix}$D$4+SUM(${oldPrefix}$A:$C)`
        const canonical = serializeFormula(parseFormula(formula))
        const renamed = renameFormulaSheetReferences(canonical, oldSheetName, newSheetName)
        const restored = renameFormulaSheetReferences(renamed, newSheetName, oldSheetName)

        expect(renamed).toContain(`"${escapeFormulaString(oldSheetName)}"`)
        expect(renamed).toContain('OldSheetTaxRate')
        expect(restored).toBe(canonical)
      },
    })
  })
})

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function escapeFormulaString(value: string): string {
  return value.replaceAll('"', '""')
}
