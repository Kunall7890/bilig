import { encodeCellRange, type XlsxCellRange } from './address.js'
import { zipSourcePreservingEntries } from './source-preserving-zip.js'
import { escapeXmlAttribute, escapeXmlText } from './xml.js'

export type SimpleXlsxCellValue = string | number | boolean

export interface SimpleXlsxCell {
  readonly address: string
  readonly row: number
  readonly col: number
  readonly value?: SimpleXlsxCellValue | null
  readonly formula?: string
  readonly styleId?: string
  readonly numberFormat?: string
}

export interface SimpleXlsxAxisEntry {
  readonly index: number
  readonly size?: number
  readonly hidden?: boolean
}

export interface SimpleXlsxMergeRange {
  readonly startAddress: string
  readonly endAddress: string
}

export interface SimpleXlsxAutoFilter {
  readonly startAddress: string
  readonly endAddress: string
}

export interface SimpleXlsxFont {
  readonly family?: string
  readonly size?: number
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly color?: string
}

export interface SimpleXlsxFill {
  readonly backgroundColor?: string
}

export interface SimpleXlsxAlignment {
  readonly horizontal?: string
  readonly vertical?: string
  readonly wrap?: boolean
  readonly indent?: number
  readonly shrinkToFit?: boolean
  readonly readingOrder?: number
  readonly textRotation?: number
  readonly justifyLastLine?: boolean
}

export interface SimpleXlsxProtection {
  readonly locked?: boolean
  readonly hidden?: boolean
}

export interface SimpleXlsxStyle {
  readonly id: string
  readonly fill?: SimpleXlsxFill
  readonly font?: SimpleXlsxFont
  readonly alignment?: SimpleXlsxAlignment
  readonly protection?: SimpleXlsxProtection
}

export interface SimpleXlsxDefinedName {
  readonly name: string
  readonly formula: string
  readonly localSheetIndex?: number
}

export interface SimpleXlsxSheet {
  readonly name: string
  readonly cells: readonly SimpleXlsxCell[]
  readonly rows?: readonly SimpleXlsxAxisEntry[]
  readonly columns?: readonly SimpleXlsxAxisEntry[]
  readonly merges?: readonly SimpleXlsxMergeRange[]
  readonly autoFilters?: readonly SimpleXlsxAutoFilter[]
  readonly dimension?: XlsxCellRange
}

export interface SimpleXlsxWorkbook {
  readonly sheets: readonly SimpleXlsxSheet[]
  readonly styles?: readonly SimpleXlsxStyle[]
  readonly definedNames?: readonly SimpleXlsxDefinedName[]
}

interface RegisteredStyle {
  readonly styleId?: string
  readonly numberFormat?: string
}

interface StyleRegistry {
  readonly styleIndexByKey: ReadonlyMap<string, number>
  readonly stylesXml: string
}

const spreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const relationshipNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const packageRelationshipNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
const contentTypesNs = 'http://schemas.openxmlformats.org/package/2006/content-types'
const customNumberFormatStartId = 164
const defaultFontXml = '<font><sz val="11"/><color theme="1"/><name val="Aptos"/><family val="2"/></font>'
const simpleWorkbookZipDosTimeParts = { time: 0, date: (1 << 5) | 1 }
const themeContentType = 'application/vnd.openxmlformats-officedocument.theme+xml'

function normalizeRgbColor(value: string | undefined): string | null {
  if (!value) {
    return null
  }
  const normalized = value.trim().replace(/^#/u, '')
  if (/^[0-9a-fA-F]{6}$/u.test(normalized)) {
    return `FF${normalized.toUpperCase()}`
  }
  if (/^[0-9a-fA-F]{8}$/u.test(normalized)) {
    return normalized.toUpperCase()
  }
  return null
}

function xmlDeclaration(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
}

function textElement(name: string, value: string): string {
  const preserveSpace = /^\s|\s$/u.test(value) || value.includes('\n') || value.includes('\t')
  const attributes = preserveSpace ? ' xml:space="preserve"' : ''
  return `<${name}${attributes}>${escapeXmlText(value)}</${name}>`
}

function formulaText(value: string): string {
  return escapeXmlText(value).replace(/"/gu, '&quot;')
}

function fontXml(font: SimpleXlsxFont | undefined): string {
  if (!font) {
    return defaultFontXml
  }
  const children: string[] = []
  if (font.size !== undefined && Number.isFinite(font.size) && font.size > 0) {
    children.push(`<sz val="${escapeXmlAttribute(String(font.size))}"/>`)
  } else {
    children.push('<sz val="11"/>')
  }
  const color = normalizeRgbColor(font.color)
  if (color) {
    children.push(`<color rgb="${color}"/>`)
  } else {
    children.push('<color theme="1"/>')
  }
  children.push(`<name val="${escapeXmlAttribute(font.family?.trim() || 'Aptos')}"/>`)
  children.push('<family val="2"/>')
  if (font.bold === true) {
    children.push('<b/>')
  }
  if (font.italic === true) {
    children.push('<i/>')
  }
  if (font.underline === true) {
    children.push('<u/>')
  }
  return `<font>${children.join('')}</font>`
}

function fillXml(fill: SimpleXlsxFill | undefined): string | null {
  const color = normalizeRgbColor(fill?.backgroundColor)
  return color ? `<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>` : null
}

function alignmentXml(alignment: SimpleXlsxAlignment | undefined): string {
  if (!alignment) {
    return ''
  }
  const attributes = [
    alignment.horizontal ? `horizontal="${escapeXmlAttribute(alignment.horizontal)}"` : null,
    alignment.vertical ? `vertical="${escapeXmlAttribute(alignment.vertical === 'middle' ? 'center' : alignment.vertical)}"` : null,
    alignment.wrap === true ? 'wrapText="1"' : null,
    alignment.indent !== undefined && alignment.indent >= 0 ? `indent="${String(alignment.indent)}"` : null,
    alignment.shrinkToFit === true ? 'shrinkToFit="1"' : null,
    alignment.readingOrder !== undefined ? `readingOrder="${String(alignment.readingOrder)}"` : null,
    alignment.textRotation !== undefined ? `textRotation="${String(alignment.textRotation)}"` : null,
    alignment.justifyLastLine === true ? 'justifyLastLine="1"' : null,
  ].filter((entry): entry is string => Boolean(entry))
  return attributes.length > 0 ? `<alignment ${attributes.join(' ')}/>` : ''
}

function protectionXml(protection: SimpleXlsxProtection | undefined): string {
  if (!protection) {
    return ''
  }
  const attributes = [
    protection.locked !== undefined ? `locked="${protection.locked ? '1' : '0'}"` : null,
    protection.hidden !== undefined ? `hidden="${protection.hidden ? '1' : '0'}"` : null,
  ].filter((entry): entry is string => Boolean(entry))
  return attributes.length > 0 ? `<protection ${attributes.join(' ')}/>` : '<protection/>'
}

function styleKey(style: RegisteredStyle): string {
  return `${style.styleId ?? ''}\u0000${style.numberFormat ?? ''}`
}

function createRegisteredStyle(styleId: string | undefined, numberFormat: string | undefined): RegisteredStyle {
  return {
    ...(styleId !== undefined ? { styleId } : {}),
    ...(numberFormat !== undefined ? { numberFormat } : {}),
  }
}

function collectRegisteredStyles(workbook: SimpleXlsxWorkbook): readonly RegisteredStyle[] {
  const keys = new Set<string>()
  const output: RegisteredStyle[] = []
  const add = (style: RegisteredStyle): void => {
    const key = styleKey(style)
    if (keys.has(key)) {
      return
    }
    keys.add(key)
    output.push(style)
  }
  add({})
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells) {
      if (cell.styleId || cell.numberFormat) {
        add(createRegisteredStyle(cell.styleId, cell.numberFormat))
      }
    }
  }
  return output
}

function buildStyleRegistry(workbook: SimpleXlsxWorkbook): StyleRegistry {
  const styleById = new Map((workbook.styles ?? []).map((style) => [style.id, style]))
  const registeredStyles = collectRegisteredStyles(workbook)
  const customFormats = [
    ...new Set(registeredStyles.map((style) => style.numberFormat).filter((format): format is string => Boolean(format))),
  ]
  const numberFormatIdByCode = new Map(customFormats.map((format, index) => [format, customNumberFormatStartId + index]))
  const fontXmls = [defaultFontXml]
  const fillXmls = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']
  const styleIndexByKey = new Map<string, number>()
  const xfs = registeredStyles.map((registeredStyle, index) => {
    styleIndexByKey.set(styleKey(registeredStyle), index)
    const style = registeredStyle.styleId ? styleById.get(registeredStyle.styleId) : undefined
    const styleFontXml = fontXml(style?.font)
    const fontId = style?.font ? fontXmls.push(styleFontXml) - 1 : 0
    const styleFillXml = fillXml(style?.fill)
    const fillId = styleFillXml ? fillXmls.push(styleFillXml) - 1 : 0
    const numberFormatId = registeredStyle.numberFormat ? (numberFormatIdByCode.get(registeredStyle.numberFormat) ?? 0) : 0
    const alignment = alignmentXml(style?.alignment)
    const protection = protectionXml(style?.protection)
    const children = `${alignment}${protection}`
    const attributes = [
      `numFmtId="${String(numberFormatId)}"`,
      `fontId="${String(fontId)}"`,
      `fillId="${String(fillId)}"`,
      'borderId="0"',
      'xfId="0"',
      fontId > 0 ? 'applyFont="1"' : null,
      fillId > 0 ? 'applyFill="1"' : null,
      numberFormatId > 0 ? 'applyNumberFormat="1"' : null,
      alignment ? 'applyAlignment="1"' : null,
      protection ? 'applyProtection="1"' : null,
    ].filter((entry): entry is string => Boolean(entry))
    return children ? `<xf ${attributes.join(' ')}>${children}</xf>` : `<xf ${attributes.join(' ')}/>`
  })
  const numFmtsXml =
    customFormats.length === 0
      ? ''
      : `<numFmts count="${String(customFormats.length)}">${customFormats
          .map(
            (format) => `<numFmt numFmtId="${String(numberFormatIdByCode.get(format) ?? 0)}" formatCode="${escapeXmlAttribute(format)}"/>`,
          )
          .join('')}</numFmts>`
  const stylesXml = [
    xmlDeclaration(),
    `<styleSheet xmlns="${spreadsheetNs}">`,
    numFmtsXml,
    `<fonts count="${String(fontXmls.length)}">${fontXmls.join('')}</fonts>`,
    `<fills count="${String(fillXmls.length)}">${fillXmls.join('')}</fills>`,
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    `<cellXfs count="${String(xfs.length)}">${xfs.join('')}</cellXfs>`,
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    '</styleSheet>',
  ].join('')
  return { styleIndexByKey, stylesXml }
}

function styleIndexForCell(cell: SimpleXlsxCell, registry: StyleRegistry): number | undefined {
  const index = registry.styleIndexByKey.get(styleKey(createRegisteredStyle(cell.styleId, cell.numberFormat)))
  return index && index > 0 ? index : undefined
}

function cellValueXml(
  value: SimpleXlsxCellValue | null | undefined,
  formula: string | undefined,
): { readonly type: string; readonly valueXml: string } {
  if (value === undefined || value === null) {
    return { type: '', valueXml: '' }
  }
  if (typeof value === 'number') {
    return { type: '', valueXml: `<v>${escapeXmlText(String(value))}</v>` }
  }
  if (typeof value === 'boolean') {
    return { type: 'b', valueXml: `<v>${value ? '1' : '0'}</v>` }
  }
  if (formula) {
    return { type: 'str', valueXml: `<v>${escapeXmlText(value)}</v>` }
  }
  return { type: 'inlineStr', valueXml: `<is>${textElement('t', value)}</is>` }
}

function cellXml(cell: SimpleXlsxCell, registry: StyleRegistry): string {
  const styleIndex = styleIndexForCell(cell, registry)
  const formula = cell.formula?.replace(/^=/u, '')
  const { type, valueXml } = cellValueXml(cell.value, formula)
  const attributes = [
    `r="${escapeXmlAttribute(cell.address)}"`,
    styleIndex !== undefined ? `s="${String(styleIndex)}"` : null,
    type ? `t="${type}"` : null,
  ].filter((entry): entry is string => Boolean(entry))
  const formulaXml = formula ? `<f>${formulaText(formula)}</f>` : ''
  const body = `${formulaXml}${valueXml}`
  return body ? `<c ${attributes.join(' ')}>${body}</c>` : `<c ${attributes.join(' ')}/>`
}

function rowXml(
  rowIndex: number,
  cells: readonly SimpleXlsxCell[],
  rowMetadata: SimpleXlsxAxisEntry | undefined,
  registry: StyleRegistry,
): string {
  const attributes = [`r="${String(rowIndex + 1)}"`]
  if (rowMetadata?.size !== undefined && Number.isFinite(rowMetadata.size) && rowMetadata.size > 0) {
    attributes.push(`ht="${String(rowMetadata.size)}"`, 'customHeight="1"')
  }
  if (rowMetadata?.hidden === true) {
    attributes.push('hidden="1"')
  }
  const body = cells.map((cell) => cellXml(cell, registry)).join('')
  return body ? `<row ${attributes.join(' ')}>${body}</row>` : `<row ${attributes.join(' ')}/>`
}

function worksheetDimension(sheet: SimpleXlsxSheet): string {
  if (sheet.dimension) {
    return encodeCellRange(sheet.dimension)
  }
  let startRow = Number.POSITIVE_INFINITY
  let startCol = Number.POSITIVE_INFINITY
  let endRow = 0
  let endCol = 0
  const visit = (address: { readonly row: number; readonly col: number }): void => {
    startRow = Math.min(startRow, address.row)
    startCol = Math.min(startCol, address.col)
    endRow = Math.max(endRow, address.row)
    endCol = Math.max(endCol, address.col)
  }
  for (const cell of sheet.cells) {
    visit(cell)
  }
  for (const merge of sheet.merges ?? []) {
    const range = encodeCellRangeAddress(merge.startAddress, merge.endAddress)
    visit({ row: range.s.r, col: range.s.c })
    visit({ row: range.e.r, col: range.e.c })
  }
  if (!Number.isFinite(startRow) || !Number.isFinite(startCol)) {
    return 'A1'
  }
  return encodeCellRange({ s: { r: startRow, c: startCol }, e: { r: endRow, c: endCol } })
}

function encodeCellRangeAddress(startAddress: string, endAddress: string): XlsxCellRange {
  const start = parseCellAddressFromA1(startAddress)
  const end = parseCellAddressFromA1(endAddress)
  return {
    s: { r: Math.min(start.row, end.row), c: Math.min(start.col, end.col) },
    e: { r: Math.max(start.row, end.row), c: Math.max(start.col, end.col) },
  }
}

function parseCellAddressFromA1(address: string): { readonly row: number; readonly col: number } {
  const match = /^([A-Z]+)([1-9][0-9]*)$/iu.exec(address.replaceAll('$', ''))
  if (!match) {
    throw new Error(`Invalid XLSX cell address: ${address}`)
  }
  let col = 0
  for (const character of match[1]!.toUpperCase()) {
    col = col * 26 + character.charCodeAt(0) - 64
  }
  return { row: Number(match[2]) - 1, col: col - 1 }
}

function columnsXml(columns: readonly SimpleXlsxAxisEntry[] | undefined): string {
  if (!columns || columns.length === 0) {
    return ''
  }
  const sorted = [...columns]
    .filter((column) => Number.isSafeInteger(column.index) && column.index >= 0)
    .toSorted((left, right) => left.index - right.index)
  const groups: string[] = []
  let index = 0
  while (index < sorted.length) {
    const first = sorted[index]!
    let last = first
    index += 1
    while (
      index < sorted.length &&
      sorted[index]!.index === last.index + 1 &&
      sorted[index]!.size === first.size &&
      sorted[index]!.hidden === first.hidden
    ) {
      last = sorted[index]!
      index += 1
    }
    const attributes = [`min="${String(first.index + 1)}"`, `max="${String(last.index + 1)}"`]
    if (first.size !== undefined && Number.isFinite(first.size) && first.size > 0) {
      attributes.push(`width="${String(Math.max(0.1, first.size / 7))}"`, 'customWidth="1"')
    }
    if (first.hidden === true) {
      attributes.push('hidden="1"')
    }
    groups.push(`<col ${attributes.join(' ')}/>`)
  }
  return groups.length > 0 ? `<cols>${groups.join('')}</cols>` : ''
}

function sheetDataXml(sheet: SimpleXlsxSheet, registry: StyleRegistry): string {
  const cellsByRow = new Map<number, SimpleXlsxCell[]>()
  for (const cell of sheet.cells) {
    let row = cellsByRow.get(cell.row)
    if (!row) {
      row = []
      cellsByRow.set(cell.row, row)
    }
    row.push(cell)
  }
  for (const row of cellsByRow.values()) {
    row.sort((left, right) => left.col - right.col)
  }
  const rowMetadataByIndex = new Map((sheet.rows ?? []).map((row) => [row.index, row]))
  const rowIndexes = [...new Set([...cellsByRow.keys(), ...rowMetadataByIndex.keys()])].toSorted((left, right) => left - right)
  return `<sheetData>${rowIndexes
    .map((rowIndex) => rowXml(rowIndex, cellsByRow.get(rowIndex) ?? [], rowMetadataByIndex.get(rowIndex), registry))
    .join('')}</sheetData>`
}

function mergeCellsXml(merges: readonly SimpleXlsxMergeRange[] | undefined): string {
  if (!merges || merges.length === 0) {
    return ''
  }
  const entries = merges.map((merge) => `<mergeCell ref="${escapeXmlAttribute(`${merge.startAddress}:${merge.endAddress}`)}"/>`)
  return `<mergeCells count="${String(entries.length)}">${entries.join('')}</mergeCells>`
}

function autoFilterXml(autoFilters: readonly SimpleXlsxAutoFilter[] | undefined): string {
  const filter = autoFilters?.[0]
  return filter ? `<autoFilter ref="${escapeXmlAttribute(`${filter.startAddress}:${filter.endAddress}`)}"/>` : ''
}

function worksheetXml(sheet: SimpleXlsxSheet, registry: StyleRegistry): string {
  return [
    xmlDeclaration(),
    `<worksheet xmlns="${spreadsheetNs}" xmlns:r="${relationshipNs}">`,
    `<dimension ref="${escapeXmlAttribute(worksheetDimension(sheet))}"/>`,
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
    '<sheetFormatPr defaultRowHeight="15"/>',
    columnsXml(sheet.columns),
    sheetDataXml(sheet, registry),
    mergeCellsXml(sheet.merges),
    autoFilterXml(sheet.autoFilters),
    '</worksheet>',
  ].join('')
}

function workbookXml(workbook: SimpleXlsxWorkbook): string {
  const sheetsXml = workbook.sheets
    .map(
      (sheet, index) => `<sheet name="${escapeXmlAttribute(sheet.name)}" sheetId="${String(index + 1)}" r:id="rId${String(index + 1)}"/>`,
    )
    .join('')
  const definedNamesXml = definedNamesXmlForWorkbook(workbook.definedNames)
  return [
    xmlDeclaration(),
    `<workbook xmlns="${spreadsheetNs}" xmlns:r="${relationshipNs}">`,
    '<workbookPr/>',
    `<sheets>${sheetsXml}</sheets>`,
    definedNamesXml,
    '</workbook>',
  ].join('')
}

function definedNamesXmlForWorkbook(definedNames: readonly SimpleXlsxDefinedName[] | undefined): string {
  if (!definedNames || definedNames.length === 0) {
    return ''
  }
  return `<definedNames>${definedNames
    .map((entry) => {
      const localSheetId =
        entry.localSheetIndex !== undefined && Number.isSafeInteger(entry.localSheetIndex) && entry.localSheetIndex >= 0
          ? ` localSheetId="${String(entry.localSheetIndex)}"`
          : ''
      return `<definedName name="${escapeXmlAttribute(entry.name)}"${localSheetId}>${escapeXmlText(entry.formula.replace(/^=/u, ''))}</definedName>`
    })
    .join('')}</definedNames>`
}

function workbookRelationshipsXml(sheetCount: number): string {
  const sheetRelationships = Array.from(
    { length: sheetCount },
    (_entry, index) =>
      `<Relationship Id="rId${String(index + 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${String(
        index + 1,
      )}.xml"/>`,
  ).join('')
  return [
    xmlDeclaration(),
    `<Relationships xmlns="${packageRelationshipNs}">`,
    sheetRelationships,
    `<Relationship Id="rId${String(sheetCount + 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    `<Relationship Id="rId${String(sheetCount + 2)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`,
    '</Relationships>',
  ].join('')
}

function rootRelationshipsXml(): string {
  return [
    xmlDeclaration(),
    `<Relationships xmlns="${packageRelationshipNs}">`,
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
    '</Relationships>',
  ].join('')
}

function contentTypesXml(sheetCount: number): string {
  const worksheets = Array.from(
    { length: sheetCount },
    (_entry, index) =>
      `<Override PartName="/xl/worksheets/sheet${String(index + 1)}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
  ).join('')
  return [
    xmlDeclaration(),
    `<Types xmlns="${contentTypesNs}">`,
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    `<Override PartName="/xl/theme/theme1.xml" ContentType="${themeContentType}"/>`,
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    worksheets,
    '</Types>',
  ].join('')
}

function corePropertiesXml(): string {
  return [
    xmlDeclaration(),
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    '<dc:creator>Bilig</dc:creator>',
    '<cp:lastModifiedBy>Bilig</cp:lastModifiedBy>',
    '</cp:coreProperties>',
  ].join('')
}

function appPropertiesXml(sheetCount: number): string {
  return [
    xmlDeclaration(),
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    '<Application>Bilig</Application>',
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${String(
      sheetCount,
    )}</vt:i4></vt:variant></vt:vector></HeadingPairs>`,
    '</Properties>',
  ].join('')
}

function themeXml(): string {
  return [
    xmlDeclaration(),
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Bilig Theme">',
    '<a:themeElements>',
    '<a:clrScheme name="Bilig">',
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>',
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>',
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>',
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>',
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>',
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>',
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>',
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>',
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>',
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>',
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>',
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>',
    '</a:clrScheme>',
    '<a:fontScheme name="Aptos">',
    '<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>',
    '<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>',
    '</a:fontScheme>',
    '<a:fmtScheme name="Bilig"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle/></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>',
    '</a:themeElements>',
    '</a:theme>',
  ].join('')
}

export function writeSimpleXlsxWorkbook(workbook: SimpleXlsxWorkbook): Uint8Array {
  if (workbook.sheets.length === 0) {
    throw new Error('Cannot write an XLSX workbook without worksheets')
  }
  const registry = buildStyleRegistry(workbook)
  const zip: Record<string, Uint8Array> = {
    '[Content_Types].xml': new TextEncoder().encode(contentTypesXml(workbook.sheets.length)),
    '_rels/.rels': new TextEncoder().encode(rootRelationshipsXml()),
    'docProps/app.xml': new TextEncoder().encode(appPropertiesXml(workbook.sheets.length)),
    'docProps/core.xml': new TextEncoder().encode(corePropertiesXml()),
    'xl/workbook.xml': new TextEncoder().encode(workbookXml(workbook)),
    'xl/_rels/workbook.xml.rels': new TextEncoder().encode(workbookRelationshipsXml(workbook.sheets.length)),
    'xl/styles.xml': new TextEncoder().encode(registry.stylesXml),
    'xl/theme/theme1.xml': new TextEncoder().encode(themeXml()),
  }
  workbook.sheets.forEach((sheet, index) => {
    zip[`xl/worksheets/sheet${String(index + 1)}.xml`] = new TextEncoder().encode(worksheetXml(sheet, registry))
  })
  return zipSourcePreservingEntries(zip, new Map(), { dosTime: simpleWorkbookZipDosTimeParts })
}
