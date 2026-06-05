import { readXlsxZipEntries } from '@bilig/xlsx'
import { describe, expect, it } from 'vitest'
import { writeBiligXlsxFixtureWorkbook } from '../bilig-xlsx-fixture-writer.ts'

const textDecoder = new TextDecoder()

describe('@bilig/xlsx fixture writer', () => {
  it('writes scalar values and formula strings through @bilig/xlsx', () => {
    const zip = readXlsxZipEntries(
      writeBiligXlsxFixtureWorkbook({
        sheetName: 'Cases',
        rows: [[10, true, 'label', '=A1*2']],
      }),
    )
    const sheetXml = textDecoder.decode(zip['xl/worksheets/sheet1.xml'])

    expect(sheetXml).toContain('<c r="A1"><v>10</v></c>')
    expect(sheetXml).toContain('<c r="B1" t="b"><v>1</v></c>')
    expect(sheetXml).toContain('<c r="C1" t="inlineStr"><is><t>label</t></is></c>')
    expect(sheetXml).toContain('<c r="D1"><f>A1*2</f></c>')
  })

  it('can blank formula strings for Excel automation fixtures', () => {
    const zip = readXlsxZipEntries(
      writeBiligXlsxFixtureWorkbook({
        sheetName: 'Cases',
        rows: [[10, '=A1*2']],
        formulaStrings: 'as-blanks',
      }),
    )
    const sheetXml = textDecoder.decode(zip['xl/worksheets/sheet1.xml'])

    expect(sheetXml).toContain('<c r="A1"><v>10</v></c>')
    expect(sheetXml).not.toContain('r="B1"')
    expect(sheetXml).not.toContain('<f>')
  })
})
