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
    expect(imported.snapshot.sheets[0]?.metadata?.commentThreads).toEqual([
      expect.objectContaining({
        address: 'A2',
        sheetName: 'Detail',
        comments: [expect.objectContaining({ authorDisplayName: 'DetailAudit', body: 'Detail relationship note' })],
      }),
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.legacyCommentVml?.vmlXml).toContain('DetailVML')

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
    expect(imported.snapshot.sheets[1]?.metadata?.commentThreads).toEqual([
      expect.objectContaining({
        address: 'A2',
        sheetName: 'Summary',
        comments: [expect.objectContaining({ authorDisplayName: 'SummaryAudit', body: 'Summary relationship note' })],
      }),
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.legacyCommentVml?.vmlXml).toContain('SummaryVML')

    expect(
      (imported.snapshot.workbook.metadata?.tables ?? []).map((table) => ({
        columnNames: table.columnNames,
        endAddress: table.endAddress,
        name: table.name,
        sheetName: table.sheetName,
        startAddress: table.startAddress,
      })),
    ).toEqual([
      { columnNames: ['detail', 'metadata'], endAddress: 'B2', name: 'DetailTable', sheetName: 'Detail', startAddress: 'A1' },
      { columnNames: ['summary', 'metadata'], endAddress: 'B2', name: 'SummaryTable', sheetName: 'Summary', startAddress: 'A1' },
    ])
  })
})

const sparklineExtensionUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const x14Namespace = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const xmNamespace = 'http://schemas.microsoft.com/office/excel/2006/main'

function buildReorderedWorksheetPathWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const summarySheet = XLSX.utils.aoa_to_sheet([['summary', 'metadata', 1, 2, 3], ['001']])
  const detailSheet = XLSX.utils.aoa_to_sheet([['detail', 'metadata', 1, 2, 3], ['002']])
  addCellComment(summarySheet, 'A2', 'SummaryAudit', 'Summary relationship note')
  addCellComment(detailSheet, 'A2', 'DetailAudit', 'Detail relationship note')
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary')
  XLSX.utils.book_append_sheet(workbook, detailSheet, 'Detail')

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
  addWorksheetTable(zip, 1, {
    columns: ['summary', 'metadata'],
    name: 'SummaryTable',
    ref: 'A1:B2',
    relationshipId: 'rId91',
    tablePath: 'xl/tables/table91.xml',
  })
  addWorksheetTable(zip, 2, {
    columns: ['detail', 'metadata'],
    name: 'DetailTable',
    ref: 'A1:B2',
    relationshipId: 'rId92',
    tablePath: 'xl/tables/table92.xml',
  })
  markLegacyCommentVml(zip, 1, 'SummaryVML')
  markLegacyCommentVml(zip, 2, 'DetailVML')

  reorderWorkbookSheets(zip, ['Detail', 'Summary'])
  return zipSync(zip)
}

function addCellComment(sheet: XLSX.WorkSheet, address: string, author: string, body: string): void {
  const cell = sheet[address]
  if (!cell) {
    throw new Error(`Missing fixture cell ${address}`)
  }
  cell.c = [{ a: author, t: body }]
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

function addWorksheetTable(
  zip: Record<string, Uint8Array>,
  sheetIndex: number,
  input: {
    readonly columns: readonly string[]
    readonly name: string
    readonly ref: string
    readonly relationshipId: string
    readonly tablePath: string
  },
): void {
  const sheetPath = `xl/worksheets/sheet${String(sheetIndex)}.xml`
  const relationshipPath = `xl/worksheets/_rels/sheet${String(sheetIndex)}.xml.rels`
  zip[sheetPath] = strToU8(insertWorksheetTablePart(strFromU8(zip[sheetPath] ?? new Uint8Array()), input.relationshipId))
  zip[relationshipPath] = strToU8(
    appendRelationship(strFromU8(zip[relationshipPath] ?? new Uint8Array()), {
      id: input.relationshipId,
      target: `../tables/${input.tablePath.slice(input.tablePath.lastIndexOf('/') + 1)}`,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
    }),
  )
  zip[input.tablePath] = strToU8(tableXml(input))
  zip['[Content_Types].xml'] = strToU8(
    appendContentTypeOverride(
      strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array()),
      `/${input.tablePath}`,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
    ),
  )
}

function insertWorksheetTablePart(sheetXml: string, relationshipId: string): string {
  const tablePart = `<tablePart r:id="${relationshipId}"/>`
  const withNamespace = ensureRelationshipNamespace(sheetXml)
  const tablePartsMatch = /<tableParts\b[^>]*\bcount="(\d+)"[^>]*>([\s\S]*?)<\/tableParts>/u.exec(withNamespace)
  if (tablePartsMatch) {
    const nextCount = Number(tablePartsMatch[1] ?? '0') + 1
    return withNamespace.replace(
      tablePartsMatch[0],
      `<tableParts count="${String(nextCount)}">${tablePartsMatch[2] ?? ''}${tablePart}</tableParts>`,
    )
  }
  const insertIndex = withNamespace.search(/<extLst\b|<\/worksheet>/u)
  if (insertIndex < 0) {
    return withNamespace
  }
  return `${withNamespace.slice(0, insertIndex)}<tableParts count="1">${tablePart}</tableParts>${withNamespace.slice(insertIndex)}`
}

function ensureRelationshipNamespace(xml: string): string {
  return /\sxmlns:r=/u.test(xml)
    ? xml
    : xml.replace(/<worksheet\b([^>]*)>/u, '<worksheet$1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
}

function appendRelationship(
  relationshipsXml: string,
  relationship: { readonly id: string; readonly target: string; readonly type: string },
): string {
  const entry = `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`
  if (!relationshipsXml.trim()) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entry}</Relationships>`
  }
  if (relationshipsXml.includes(`Id="${relationship.id}"`)) {
    return relationshipsXml
  }
  return relationshipsXml.replace('</Relationships>', `${entry}</Relationships>`)
}

function appendContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (contentTypesXml.includes(`PartName="${partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)
}

function tableXml(input: { readonly columns: readonly string[]; readonly name: string; readonly ref: string }): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${input.name}" displayName="${input.name}" ref="${input.ref}" headerRowCount="1" totalsRowShown="0">`,
    `<autoFilter ref="${input.ref}"/>`,
    `<tableColumns count="${String(input.columns.length)}">`,
    ...input.columns.map((column, index) => `<tableColumn id="${String(index + 1)}" name="${column}"/>`),
    '</tableColumns>',
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    '</table>',
  ].join('')
}

function markLegacyCommentVml(zip: Record<string, Uint8Array>, sheetIndex: number, marker: string): void {
  const sheetPath = `xl/worksheets/sheet${String(sheetIndex)}.xml`
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const relationshipId = readXmlAttribute(/<legacyDrawing\b[^>]*>/u.exec(sheetXml)?.[0] ?? '', 'id')
  if (!relationshipId) {
    throw new Error(`Missing legacy comment VML relationship for ${sheetPath}`)
  }
  const relationshipsPath = `xl/worksheets/_rels/sheet${String(sheetIndex)}.xml.rels`
  const target = readRelationshipTarget(strFromU8(zip[relationshipsPath] ?? new Uint8Array()), relationshipId, 'vmlDrawing')
  if (!target) {
    throw new Error(`Missing legacy comment VML target for ${sheetPath}`)
  }
  const vmlPath = resolveTargetPath(sheetPath, target)
  zip[vmlPath] = strToU8(`${strFromU8(zip[vmlPath] ?? new Uint8Array())}<!--${marker}-->`)
}

function readRelationshipTarget(relationshipsXml: string, relationshipId: string, relationshipTypeSuffix: string): string | undefined {
  const relationship = new RegExp(
    `<Relationship\\b(?=[^>]*\\bId="${escapeRegExp(relationshipId)}")(?=[^>]*\\bType="[^"]*${escapeRegExp(
      relationshipTypeSuffix,
    )}")[^>]*\\bTarget="([^"]*)"`,
    'u',
  ).exec(relationshipsXml)?.[1]
  return relationship
}

function resolveTargetPath(basePartPath: string, target: string): string {
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
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
