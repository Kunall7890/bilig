import { describe, expect, it } from 'vitest'

import { ErrorCode, ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const EXTERNAL_FORMULA = "'[1]THB data'!BK150/100"
const FORMULA_COUNT = 33

describe('engine external workbook formula caches', () => {
  it('preserves cached values for imported external workbook formulas during full recalculation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'external-workbook-formula-cache' })
    await engine.ready()

    engine.importSnapshot(buildExternalWorkbookFormulaSnapshot({ preserveCachedValues: true }))
    engine.recalculateNow()

    expect(engine.getCellValue('Regional network', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.25 })
    expect(engine.getCellValue('Regional network', `A${FORMULA_COUNT}`)).toEqual({
      tag: ValueTag.Number,
      value: 0.25 + FORMULA_COUNT - 1,
    })
  })

  it('preserves cached values for parser-hostile imported external workbook formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'external-workbook-formula-parser-hostile-cache' })
    await engine.ready()

    engine.importSnapshot(buildParserHostileExternalWorkbookFormulaSnapshot())

    expect(engine.getCellValue('Opened', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'ACCORD MORTGAGES LIMITED',
    })

    engine.recalculateNow()

    expect(engine.getCellValue('Opened', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'ACCORD MORTGAGES LIMITED',
    })
  })

  it('imports uncached external workbook formulas without using the direct-scalar fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'external-workbook-formula-uncached' })
    await engine.ready()

    expect(() => engine.importSnapshot(buildExternalWorkbookFormulaSnapshot({ preserveCachedValues: false }))).not.toThrow()
    engine.recalculateNow()

    expect(engine.getCellValue('Regional network', 'A1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })
})

function buildExternalWorkbookFormulaSnapshot(options: { readonly preserveCachedValues: boolean }): WorkbookSnapshot {
  const cells = Array.from({ length: FORMULA_COUNT }, (_, index) => {
    const row = index
    const address = `A${index + 1}`
    return {
      address,
      row,
      col: 0,
      formula: EXTERNAL_FORMULA,
      ...(options.preserveCachedValues ? { value: 0.25 + index } : {}),
    }
  })

  return {
    version: 1,
    workbook: {
      name: 'Imported external workbook formulas',
      metadata: {
        externalWorkbookReferences: [
          {
            bookIndex: 1,
            target: 'file:///tmp/Data for website.xlsx',
            targetMode: 'External',
            workbookName: 'Data for website.xlsx',
            sheetNames: ['THB data', 'Regional network'],
          },
        ],
        unsupportedFormulaDependencies: cells.map((cell) => ({
          kind: 'external-workbook-reference',
          sheetName: 'Regional network',
          address: cell.address,
          formula: EXTERNAL_FORMULA,
          importedFormula: EXTERNAL_FORMULA,
          linkedWorkbooks: [
            {
              bookIndex: 1,
              target: 'file:///tmp/Data for website.xlsx',
              targetMode: 'External',
              workbookName: 'Data for website.xlsx',
              sheetNames: ['THB data', 'Regional network'],
            },
          ],
          cachedValuesUsed: options.preserveCachedValues,
          cachedFormulaValuePreserved: options.preserveCachedValues,
          cachedExternalReferenceValuesUsed: false,
          resolvedExternalReferenceCount: 0,
          unresolvedExternalReferenceCount: 1,
          reason:
            'Formula depends on an external workbook reference; cached linked values are preserved but linked workbooks are not recalculated during import.',
        })),
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Regional network',
        order: 0,
        cells,
      },
    ],
  }
}

function buildParserHostileExternalWorkbookFormulaSnapshot(): WorkbookSnapshot {
  const formula = 'IFERROR(VLOOKUP(A2,[1]Closed!$A:$B,2,FALSE),"NO GROUP")'
  return {
    version: 1,
    workbook: {
      name: 'Imported parser-hostile external workbook formulas',
      metadata: {
        externalWorkbookReferences: [
          {
            bookIndex: 1,
            target: 'file:///tmp/firm-level-complaints-data-2021-h1.xlsx',
            targetMode: 'External',
            workbookName: 'firm-level-complaints-data-2021-h1.xlsx',
            sheetNames: ['Closed'],
          },
        ],
        unsupportedFormulaDependencies: [
          {
            kind: 'external-workbook-reference',
            sheetName: 'Opened',
            address: 'B2',
            formula,
            importedFormula: formula,
            linkedWorkbooks: [
              {
                bookIndex: 1,
                target: 'file:///tmp/firm-level-complaints-data-2021-h1.xlsx',
                targetMode: 'External',
                workbookName: 'firm-level-complaints-data-2021-h1.xlsx',
                sheetNames: ['Closed'],
              },
            ],
            cachedValuesUsed: true,
            cachedFormulaValuePreserved: true,
            cachedExternalReferenceValuesUsed: false,
            resolvedExternalReferenceCount: 0,
            unresolvedExternalReferenceCount: 0,
            reason:
              'Formula depends on an external workbook reference; cached linked values are preserved but linked workbooks are not recalculated during import.',
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Opened',
        order: 0,
        cells: [
          { address: 'A2', row: 1, col: 0, value: 'Accord Mortgages Limited' },
          {
            address: 'B2',
            row: 1,
            col: 1,
            formula,
            value: 'ACCORD MORTGAGES LIMITED',
          },
        ],
      },
    ],
  }
}
