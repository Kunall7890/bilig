import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

describe('precision-as-displayed calculation', () => {
  it('uses displayed precision for formula outputs when fullPrecision is disabled', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'precision-as-displayed' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCalculationSettings({ fullPrecision: false })

    engine.setCellValue('Sheet1', 'A1', 1.234)
    engine.setCellFormat('Sheet1', 'A1', '0.0')
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormat('Sheet1', 'B1', '0.0')
    engine.setCellFormula('Sheet1', 'C1', 'B1*10')
    engine.recalculateNow()

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1.234 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2.5 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 25 })
  })

  it('uses rounded percent-formatted formula outputs in dependent formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'precision-as-displayed-percent' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCalculationSettings({ fullPrecision: false })

    engine.setCellValue('Sheet1', 'A1', 0.12345)
    engine.setCellFormula('Sheet1', 'B1', 'A1')
    engine.setCellFormat('Sheet1', 'B1', '0.0%')
    engine.setCellFormula('Sheet1', 'C1', 'B1*100')
    engine.recalculateNow()

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.12345 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0.123 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12.3 })
  })
})
