import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { importXlsx } from '../index.js'

describe('worksheet relationship path import', () => {
  it('keeps sheet-local metadata on workbook relationship targets after tab reorder', () => {
    const imported = importXlsx(buildReorderedWorksheetPathWorkbook(), 'worksheet-path-reordered.xlsx')

    expect(imported.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Detail', 'Summary'])
    expect(imported.snapshot.sheets[0]?.metadata).toMatchObject({
      sheetPr: { xml: '<sheetPr codeName="DetailCode"><outlinePr summaryBelow="0"/></sheetPr>' },
      tabColor: { rgb: 'FF0000FF' },
      freezePane: { rows: 2, cols: 1, topLeftCell: 'B3', activePane: 'bottomRight' },
    })
    expect(imported.snapshot.sheets[0]?.metadata?.ignoredErrors?.xml).toContain('sqref="B2"')
    expect(imported.snapshot.sheets[0]?.metadata?.ignoredErrors?.xml).toContain('evalError="1"')
    expect(imported.snapshot.sheets[0]?.metadata?.cellMetadataRefs).toEqual([expect.objectContaining({ address: 'B1', cm: '7', vm: '8' })])
    expect(imported.snapshot.sheets[0]?.metadata?.arrayFormulas?.formulas).toEqual([
      expect.objectContaining({ address: 'C1', formulaXml: '<f t="array" ref="C1:C2">ROW(A1:A2)</f>' }),
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.dataTableFormulas?.formulas).toEqual([
      expect.objectContaining({ address: 'D1', formulaXml: '<f t="dataTable" ref="D1:E2" r1="A1"/>' }),
    ])

    expect(imported.snapshot.sheets[1]?.metadata).toMatchObject({
      sheetPr: { xml: '<sheetPr codeName="SummaryCode"><outlinePr summaryRight="0"/></sheetPr>' },
      tabColor: { rgb: 'FFFF0000' },
      freezePane: { rows: 1, cols: 0, topLeftCell: 'A2', activePane: 'bottomLeft' },
    })
    expect(imported.snapshot.sheets[1]?.metadata?.ignoredErrors?.xml).toContain('sqref="A2"')
    expect(imported.snapshot.sheets[1]?.metadata?.ignoredErrors?.xml).toContain('numberStoredAsText="1"')
    expect(imported.snapshot.sheets[1]?.metadata?.cellMetadataRefs).toEqual([expect.objectContaining({ address: 'B1', cm: '3', vm: '4' })])
  })
})

function buildReorderedWorksheetPathWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['summary', 'metadata', 1, 2, 3], ['001']]), 'Summary')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['detail', 'metadata', 1, 2, 3], ['002']]), 'Detail')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  updateWorksheet(zip, 1, {
    sheetPr: '<sheetPr codeName="SummaryCode"><tabColor rgb="FFFF0000"/><outlinePr summaryRight="0"/></sheetPr>',
    sheetViews:
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>',
    ignoredErrors: '<ignoredErrors><ignoredError sqref="A2" numberStoredAsText="1"/></ignoredErrors>',
    cellMetadata: { address: 'B1', cm: '3', vm: '4' },
  })
  updateWorksheet(zip, 2, {
    sheetPr: '<sheetPr codeName="DetailCode"><tabColor rgb="FF0000FF"/><outlinePr summaryBelow="0"/></sheetPr>',
    sheetViews:
      '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/><selection pane="bottomRight"/></sheetView></sheetViews>',
    ignoredErrors: '<ignoredErrors><ignoredError sqref="B2" evalError="1"/></ignoredErrors>',
    cellMetadata: { address: 'B1', cm: '7', vm: '8' },
    cellReplacements: new Map([
      ['C1', '<c r="C1"><f t="array" ref="C1:C2">ROW(A1:A2)</f><v>1</v></c>'],
      ['D1', '<c r="D1"><f t="dataTable" ref="D1:E2" r1="A1"/><v>2</v></c>'],
    ]),
  })

  reorderWorkbookSheets(zip, ['Detail', 'Summary'])
  return zipSync(zip)
}

function updateWorksheet(
  zip: Record<string, Uint8Array>,
  sheetIndex: number,
  input: {
    readonly sheetPr: string
    readonly sheetViews: string
    readonly ignoredErrors: string
    readonly cellMetadata: { readonly address: string; readonly cm: string; readonly vm: string }
    readonly cellReplacements?: ReadonlyMap<string, string>
  },
): void {
  const path = `xl/worksheets/sheet${String(sheetIndex)}.xml`
  let xml = strFromU8(zip[path] ?? new Uint8Array())
  xml = xml.replace(/<sheetPr\b[^>]*(?:\/>|>[\s\S]*?<\/sheetPr>)/u, '')
  xml = xml.replace(/<sheetViews\b[^>]*(?:\/>|>[\s\S]*?<\/sheetViews>)/u, '')
  xml = xml.replace(/<ignoredErrors\b[^>]*(?:\/>|>[\s\S]*?<\/ignoredErrors>)/gu, '')
  xml = xml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1>${input.sheetPr}${input.sheetViews}`)
  xml = xml.replace('</worksheet>', `${input.ignoredErrors}</worksheet>`)
  xml = addCellMetadata(xml, input.cellMetadata.address, input.cellMetadata.cm, input.cellMetadata.vm)
  for (const [address, cellXml] of input.cellReplacements ?? []) {
    xml = replaceCellXml(xml, address, cellXml)
  }
  zip[path] = strToU8(xml)
}

function addCellMetadata(sheetXml: string, address: string, cm: string, vm: string): string {
  const addressPattern = escapeRegExp(address)
  return sheetXml.replace(
    new RegExp(`<c\\b([^>]*\\br="${addressPattern}"[^>]*)>`, 'u'),
    (_match, attributes: string) => `<c${attributes} cm="${cm}" vm="${vm}">`,
  )
}

function replaceCellXml(sheetXml: string, address: string, cellXml: string): string {
  const addressPattern = escapeRegExp(address)
  return sheetXml.replace(new RegExp(`<c\\b[^>]*\\br="${addressPattern}"[^>]*(?:/>|>[\\s\\S]*?<\\/c>)`, 'u'), cellXml)
}

function reorderWorkbookSheets(zip: Record<string, Uint8Array>, sheetNames: readonly string[]): void {
  const path = 'xl/workbook.xml'
  const workbookXml = strFromU8(zip[path] ?? new Uint8Array())
  const sheets = [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*\/>/gu)].map((match) => match[0])
  const sheetsByName = new Map(sheets.map((sheetXml) => [readXmlAttribute(sheetXml, 'name'), sheetXml]))
  const reorderedSheets = sheetNames.map((sheetName) => {
    const sheetXml = sheetsByName.get(sheetName)
    if (!sheetXml) {
      throw new Error(`Missing sheet ${sheetName}`)
    }
    return sheetXml
  })
  zip[path] = strToU8(
    workbookXml.replace(
      /<((?:[A-Za-z_][\w.-]*:)?sheets)\b[^>]*>[\s\S]*?<\/\1>/u,
      (source, tagName: string) => `<${tagName}>${reorderedSheets.join('')}</${tagName}>`,
    ),
  )
}

function readXmlAttribute(xml: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(xml)?.[1]
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
