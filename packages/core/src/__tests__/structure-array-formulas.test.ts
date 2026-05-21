import { SpreadsheetEngine } from '../engine.js'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

describe('structural array formula metadata', () => {
  it('restores and exports imported native array formula metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'array-formula-metadata-restore' })
    await engine.ready()
    engine.importSnapshot(arrayFormulaSnapshot())

    expect(engine.exportSnapshot().sheets[0]?.metadata?.arrayFormulas).toEqual(arrayFormulaMetadata())
  })

  it('retargets native array formula metadata through structural row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'array-formula-row-insert' })
    await engine.ready()
    engine.importSnapshot(arrayFormulaSnapshot())

    engine.insertRows('Forecast', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.arrayFormulas).toEqual({
      formulas: [{ address: 'D3', formulaXml: '<f t="array" ref="D3:D5">TRANSPOSE(A3:C3)</f>' }],
    })
  })

  it('fails closed instead of exporting stale native array metadata through destructive edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'array-formula-delete-through-range' })
    await engine.ready()
    engine.importSnapshot(arrayFormulaSnapshot())

    engine.deleteRows('Forecast', 2, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.arrayFormulas).toBeUndefined()
  })

  it('preserves no-op native array formula XML byte-for-byte after the formula range', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'array-formula-no-op-native-xml' })
    await engine.ready()
    engine.importSnapshot(futureFunctionArrayFormulaSnapshot())

    engine.insertRows('Forecast', 10, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.arrayFormulas).toEqual(futureFunctionArrayFormulaMetadata())
  })
})

function arrayFormulaSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Array formula metadata' },
    sheets: [
      {
        id: 1,
        name: 'Forecast',
        order: 0,
        cells: [
          { address: 'A2', value: 10 },
          { address: 'B2', value: 20 },
          { address: 'C2', value: 30 },
          { address: 'D2', formula: '=TRANSPOSE(A2:C2)', value: 10 },
          { address: 'D3', value: 20 },
          { address: 'D4', value: 30 },
        ],
        metadata: { arrayFormulas: arrayFormulaMetadata() },
      },
    ],
  }
}

function arrayFormulaMetadata() {
  return {
    formulas: [{ address: 'D2', formulaXml: '<f t="array" ref="D2:D4">TRANSPOSE(A2:C2)</f>' }],
  }
}

function futureFunctionArrayFormulaSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Array formula future function metadata' },
    sheets: [
      {
        id: 1,
        name: 'Forecast',
        order: 0,
        cells: [
          { address: 'E2', value: 100 },
          { address: 'F2', value: 101 },
          { address: 'E3', value: 102 },
          { address: 'F3', value: 103 },
        ],
        metadata: { arrayFormulas: futureFunctionArrayFormulaMetadata() },
      },
    ],
  }
}

function futureFunctionArrayFormulaMetadata() {
  return {
    formulas: [
      {
        address: 'E2',
        formulaXml: '<f t="array" aca="1" ref="E2:F3" ca="1">_xlfn.STOCKHISTORY(&quot;MSFT&quot;,&quot;12/1/2020&quot;,,2)</f>',
      },
    ],
  }
}
