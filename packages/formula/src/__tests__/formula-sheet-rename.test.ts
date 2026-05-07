import { describe, expect, it } from 'vitest'
import { compileFormulaAst } from '../compiler.js'
import { parseFormula } from '../parser.js'
import {
  renameCompiledFormulaSheetReferenceMetadataInPlace,
  renameCompiledFormulaSheetReferences,
  renameFormulaSheetReferences,
} from '../formula-sheet-rename.js'

function compile(source: string) {
  return compileFormulaAst(source, parseFormula(source))
}

describe('formula sheet rename', () => {
  it('renames parsed formulas, compiled references, and metadata in place', () => {
    expect(renameFormulaSheetReferences("'Old Sheet'!A1+SUM('Old Sheet'!B1:B3)", 'Old Sheet', 'New Sheet')).toBe(
      "'New Sheet'!A1+SUM('New Sheet'!B1:B3)",
    )

    const renamed = renameCompiledFormulaSheetReferences(compile("'Old Sheet'!A1+'Old Sheet'!B1"), 'Old Sheet', 'New Sheet')
    expect(renamed.reusedProgram).toBe(true)
    expect(renamed.source).toBe("'New Sheet'!A1+'New Sheet'!B1")
    expect(renamed.compiled.symbolicRefs).toEqual(["'New Sheet'!A1", "'New Sheet'!B1"])

    const compiled = compile("'Old Sheet'!A1+'Old Sheet'!B1")
    expect(renameCompiledFormulaSheetReferenceMetadataInPlace(compiled, 'Old Sheet', 'New Sheet')).toBe(true)
    expect(compiled.symbolicRefs).toEqual(["'New Sheet'!A1", "'New Sheet'!B1"])
    expect(compiled.astMatchesSource).toBe(false)
  })
})
