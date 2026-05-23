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
    expect(imported.snapshot.sheets[0]?.metadata?.filters).toEqual([
      expect.objectContaining({ sheetName: 'Detail', startAddress: 'B1', endAddress: 'C2' }),
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.sorts).toEqual([
      expect.objectContaining({
        range: { sheetName: 'Detail', startAddress: 'B1', endAddress: 'C2' },
        keys: [{ keyAddress: 'C2', direction: 'desc' }],
      }),
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.validations).toEqual([
      expect.objectContaining({
        range: { sheetName: 'Detail', startAddress: 'E1', endAddress: 'E1' },
        rule: { kind: 'whole', operator: 'greaterThan', values: [10] },
      }),
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.sheetProtection).toMatchObject({ sheetName: 'Detail' })
    expect(imported.snapshot.sheets[0]?.metadata?.sheetProtection?.xmlAttributes).toContainEqual({ name: 'sort', value: '1' })
    expect(imported.snapshot.sheets[0]?.metadata?.protectedRanges).toEqual([
      expect.objectContaining({
        id: 'DetailLock',
        range: { sheetName: 'Detail', startAddress: 'B2', endAddress: 'B2' },
      }),
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.conditionalFormats).toEqual([
      expect.objectContaining({
        range: { sheetName: 'Detail', startAddress: 'B2', endAddress: 'B2' },
        rule: { kind: 'cellIs', operator: 'lessThan', values: [10] },
      }),
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml).toContain('FF00AAFF')
    expect(imported.snapshot.sheets[0]?.metadata?.sparklines?.xml).toContain('Detail!B2:C2')

    expect(imported.snapshot.sheets[1]?.metadata).toMatchObject({
      sheetPr: { xml: '<sheetPr codeName="SummaryCode"><outlinePr summaryRight="0"/></sheetPr>' },
      tabColor: { rgb: 'FFFF0000' },
      freezePane: { rows: 1, cols: 0, topLeftCell: 'A2', activePane: 'bottomLeft' },
    })
    expect(imported.snapshot.sheets[1]?.metadata?.ignoredErrors?.xml).toContain('sqref="A2"')
    expect(imported.snapshot.sheets[1]?.metadata?.ignoredErrors?.xml).toContain('numberStoredAsText="1"')
    expect(imported.snapshot.sheets[1]?.metadata?.cellMetadataRefs).toEqual([expect.objectContaining({ address: 'B1', cm: '3', vm: '4' })])
    expect(imported.snapshot.sheets[1]?.metadata?.filters).toEqual([
      expect.objectContaining({ sheetName: 'Summary', startAddress: 'A1', endAddress: 'B2' }),
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.sorts).toEqual([
      expect.objectContaining({
        range: { sheetName: 'Summary', startAddress: 'A1', endAddress: 'B2' },
        keys: [{ keyAddress: 'B2', direction: 'asc' }],
      }),
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.validations).toEqual([
      expect.objectContaining({
        range: { sheetName: 'Summary', startAddress: 'E1', endAddress: 'E1' },
        rule: { kind: 'list', values: ['A', 'B'] },
      }),
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.sheetProtection).toMatchObject({ sheetName: 'Summary' })
    expect(imported.snapshot.sheets[1]?.metadata?.sheetProtection?.xmlAttributes).toContainEqual({ name: 'formatCells', value: '0' })
    expect(imported.snapshot.sheets[1]?.metadata?.protectedRanges).toEqual([
      expect.objectContaining({
        id: 'SummaryLock',
        range: { sheetName: 'Summary', startAddress: 'A2', endAddress: 'A2' },
      }),
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.conditionalFormats).toEqual([
      expect.objectContaining({
        range: { sheetName: 'Summary', startAddress: 'A2', endAddress: 'A2' },
        rule: { kind: 'cellIs', operator: 'greaterThan', values: [0] },
      }),
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.conditionalFormatArtifacts?.xml).toContain('FFFFAA00')
    expect(imported.snapshot.sheets[1]?.metadata?.sparklines?.xml).toContain('Summary!A2:B2')
  })
})

const sparklineExtensionUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const x14Namespace = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const xmNamespace = 'http://schemas.microsoft.com/office/excel/2006/main'

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
    sheetProtection: '<sheetProtection sheet="1" formatCells="0"/>',
    protectedRanges: '<protectedRanges><protectedRange name="SummaryLock" sqref="A2"/></protectedRanges>',
    autoFilter: '<autoFilter ref="A1:B2"><filterColumn colId="0"><filters><filter val="summary"/></filters></filterColumn></autoFilter>',
    sortState: '<sortState ref="A1:B2"><sortCondition ref="B2:B2"/></sortState>',
    dataValidations:
      '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="E1"><formula1>"A,B"</formula1></dataValidation></dataValidations>',
    conditionalFormats: [
      '<conditionalFormatting sqref="A2"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="A2"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFAA00"/><color rgb="FFFFCC66"/></colorScale></cfRule></conditionalFormatting>',
    ],
    sparklines: sparklineExtensionXml('Summary!A2:B2', 'F2'),
  })
  updateWorksheet(zip, 2, {
    sheetPr: '<sheetPr codeName="DetailCode"><tabColor rgb="FF0000FF"/><outlinePr summaryBelow="0"/></sheetPr>',
    sheetViews:
      '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/><selection pane="bottomRight"/></sheetView></sheetViews>',
    ignoredErrors: '<ignoredErrors><ignoredError sqref="B2" evalError="1"/></ignoredErrors>',
    cellMetadata: { address: 'B1', cm: '7', vm: '8' },
    sheetProtection: '<sheetProtection sheet="1" sort="1"/>',
    protectedRanges: '<protectedRanges><protectedRange name="DetailLock" sqref="B2"/></protectedRanges>',
    autoFilter: '<autoFilter ref="B1:C2"><filterColumn colId="0"><filters><filter val="detail"/></filters></filterColumn></autoFilter>',
    sortState: '<sortState ref="B1:C2"><sortCondition descending="1" ref="C2:C2"/></sortState>',
    dataValidations:
      '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="E1"><formula1>10</formula1></dataValidation></dataValidations>',
    conditionalFormats: [
      '<conditionalFormatting sqref="B2"><cfRule type="cellIs" priority="1" operator="lessThan"><formula>10</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="B2"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF00AAFF"/><color rgb="FF66CCFF"/></colorScale></cfRule></conditionalFormatting>',
    ],
    sparklines: sparklineExtensionXml('Detail!B2:C2', 'F2'),
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
    readonly sheetProtection?: string
    readonly protectedRanges?: string
    readonly autoFilter?: string
    readonly sortState?: string
    readonly dataValidations?: string
    readonly conditionalFormats?: readonly string[]
    readonly sparklines?: string
    readonly cellReplacements?: ReadonlyMap<string, string>
  },
): void {
  const path = `xl/worksheets/sheet${String(sheetIndex)}.xml`
  let xml = strFromU8(zip[path] ?? new Uint8Array())
  xml = xml.replace(/<sheetPr\b[^>]*(?:\/>|>[\s\S]*?<\/sheetPr>)/u, '')
  xml = xml.replace(/<sheetViews\b[^>]*(?:\/>|>[\s\S]*?<\/sheetViews>)/u, '')
  xml = xml.replace(/<ignoredErrors\b[^>]*(?:\/>|>[\s\S]*?<\/ignoredErrors>)/gu, '')
  xml = xml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1>${input.sheetPr}${input.sheetViews}`)
  xml = insertWorksheetPolicyXml(
    xml,
    [
      input.sheetProtection,
      input.protectedRanges,
      input.autoFilter,
      input.sortState,
      ...(input.conditionalFormats ?? []),
      input.dataValidations,
      input.ignoredErrors,
      input.sparklines ? `<extLst>${input.sparklines}</extLst>` : undefined,
    ].flatMap((part) => (part ? [part] : [])),
  )
  xml = addCellMetadata(xml, input.cellMetadata.address, input.cellMetadata.cm, input.cellMetadata.vm)
  for (const [address, cellXml] of input.cellReplacements ?? []) {
    xml = replaceCellXml(xml, address, cellXml)
  }
  zip[path] = strToU8(xml)
}

function sparklineExtensionXml(formula: string, sqref: string): string {
  return `<ext uri="${sparklineExtensionUri}" xmlns:x14="${x14Namespace}"><x14:sparklineGroups xmlns:xm="${xmNamespace}"><x14:sparklineGroup type="line"><x14:sparklines><x14:sparkline><xm:f>${formula}</xm:f><xm:sqref>${sqref}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`
}

function insertWorksheetPolicyXml(sheetXml: string, policyXml: readonly string[]): string {
  const insertIndex = sheetXml.search(/<pageMargins\b|<\/worksheet>/u)
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${policyXml.join('')}${sheetXml.slice(insertIndex)}`
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
