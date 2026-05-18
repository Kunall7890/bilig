import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { importXlsx } from '../index.js'

describe('MS-OI29500 formula context audit import', () => {
  it('records worksheet, defined-name, CF, DV, shared-formula, and cached-value provenance', () => {
    const imported = importXlsx(buildFormulaContextWorkbookBytes(), 'ms-oi29500-formula-context.xlsx')
    const formulaAudit = imported.snapshot.workbook.metadata?.formulaAudit

    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'B2', formula: 'A2*2', value: 20 }),
        expect.objectContaining({ address: 'B3', formula: 'A3*2', value: 30 }),
      ]),
    )
    expect(formulaAudit).toMatchObject({
      formulas: expect.arrayContaining([
        expect.objectContaining({
          context: 'worksheet-cell',
          sheetName: 'Model',
          address: 'B2',
          formula: 'A2*2',
          formulaType: 'shared',
          sharedIndex: '7',
          ref: 'B2:B3',
          cachedValue: 20,
          cacheStatus: 'trustedCached',
          clause: '18.3.1.40',
        }),
        expect.objectContaining({
          context: 'worksheet-cell',
          sheetName: 'Model',
          address: 'B3',
          formula: 'A3*2',
          formulaType: 'shared',
          sharedIndex: '7',
          cachedValue: 30,
          cacheStatus: 'trustedCached',
        }),
        expect.objectContaining({
          context: 'defined-name',
          name: 'NameFormula',
          formula: 'R1C1+1',
        }),
        expect.objectContaining({
          context: 'conditional-format',
          sheetName: 'Model',
          formula: 'A1:A2',
        }),
        expect.objectContaining({
          context: 'data-validation',
          sheetName: 'Model',
          formula: 'R1C1',
        }),
      ]),
      diagnostics: expect.arrayContaining([
        expect.objectContaining({
          code: 'r1c1-reference',
          context: 'defined-name',
          clause: '3.2.3.1',
        }),
        expect.objectContaining({
          code: 'conditional-format-range-reference',
          context: 'conditional-format',
          clause: '3.2.3.1',
        }),
        expect.objectContaining({
          code: 'data-validation-r1c1-reference',
          context: 'data-validation',
          clause: '3.2.3.1',
        }),
      ]),
    })
  })
})

function buildFormulaContextWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Input', 'Double'],
    [10, { f: 'A2*2', v: 20 }],
    [15, { f: 'A3*2', v: 30 }],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')
  workbook.Workbook = {
    Names: [
      {
        Name: 'NameFormula',
        Ref: 'R1C1+1',
      },
    ],
  }

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sourceSheetXml
      .replace(/<c r="B([23])" t="str">/gu, '<c r="B$1">')
      .replace('<f>A2*2</f><v>20</v>', '<f t="shared" si="7" ref="B2:B3">A2*2</f><v>20</v>')
      .replace('<f>A3*2</f><v>30</v>', '<f t="shared" si="7"/><v>30</v>')
      .replace(
        '</worksheet>',
        [
          '<conditionalFormatting sqref="A2:A3"><cfRule type="expression" priority="1"><formula>A1:A2</formula></cfRule></conditionalFormatting>',
          '<dataValidations count="1"><dataValidation type="custom" sqref="A2:A3"><formula1>R1C1</formula1></dataValidation></dataValidations>',
          '</worksheet>',
        ].join(''),
      ),
  )
  return zipSync(zip)
}
