import * as fc from 'fast-check'
import { strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx } from '../index.js'
import { decodeExcelEscapedText } from '../xlsx-escaped-text.js'

type SheetSnapshot = WorkbookSnapshot['sheets'][number]
type SheetCellSnapshot = SheetSnapshot['cells'][number]
type SheetMetadataSnapshot = NonNullable<SheetSnapshot['metadata']>
type WorkbookMetadataSnapshot = NonNullable<WorkbookSnapshot['workbook']['metadata']>
type AxisEntrySnapshot = NonNullable<SheetMetadataSnapshot['rows']>[number]
type CellStyleSnapshot = NonNullable<WorkbookMetadataSnapshot['styles']>[number]

export interface SemanticXlsxWorkbookSpec {
  workbookName: string
  account: string
  amount: number
  status: 'Open' | 'Closed' | 'Needs Review'
  note: string
  richTextStorage: 'sharedString' | 'inlineString'
  lookupVisibility: 'hidden' | 'veryHidden' | 'visible'
  hiddenAxis: boolean
  protectedAttributes: boolean
  blankFormatAddress: 'E2' | 'F6' | 'H10'
}

export interface CsvEdgeWorkbookSpec {
  fileStem: string
  delimiter: ',' | ';' | '\t'
  newline: '\n' | '\r\n'
  rows: CsvEdgeCellSpec[][]
}

export type CsvEdgeCellSpec =
  | { kind: 'empty' }
  | { kind: 'plainText'; value: string }
  | { kind: 'quotedText'; value: string }
  | { kind: 'leadingZero'; value: string }
  | { kind: 'formula'; formula: string }
  | { kind: 'formulaLikeText'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'number'; raw: string; value: number }
  | { kind: 'accounting'; raw: string; value: number }

export interface MalformedXlsxSpec {
  kind: 'corruptZip' | 'malformedOpenXml'
  path?: string
  replacement?: string
  byteSuffix?: number[]
}

export const semanticXlsxWorkbookSpecArbitrary: fc.Arbitrary<SemanticXlsxWorkbookSpec> = fc.record({
  workbookName: fc.constantFrom('Semantic Fuzz', 'Import Export Fidelity', 'Metadata Digest'),
  account: fc.constantFrom('0010', 'ACC-42', 'North-Ω', '東京-07'),
  amount: fc.integer({ min: -250_000, max: 250_000 }).map((value) => value / 100),
  status: fc.constantFrom('Open', 'Closed', 'Needs Review'),
  note: fc.constantFrom('Reviewed', 'Quote "ok"', 'Café total', 'Line 1\nLine 2', 'Ampersand & less < more >'),
  richTextStorage: fc.constantFrom('sharedString', 'inlineString'),
  lookupVisibility: fc.constantFrom('hidden', 'veryHidden', 'visible'),
  hiddenAxis: fc.boolean(),
  protectedAttributes: fc.boolean(),
  blankFormatAddress: fc.constantFrom('E2', 'F6', 'H10'),
})

const csvEdgeCellSpecArbitrary: fc.Arbitrary<CsvEdgeCellSpec> = fc.oneof(
  fc.constant({ kind: 'empty' }),
  fc.constantFrom('alpha', ' spaced value ', 'café', '東京', 'emoji 😀').map((value) => ({ kind: 'plainText' as const, value })),
  fc
    .constantFrom('contains,comma', 'contains;semicolon', 'contains\ttab', 'quote "inside"', 'line 1\nline 2')
    .map((value) => ({ kind: 'quotedText' as const, value })),
  fc.constantFrom('001', '00042', '021000021').map((value) => ({ kind: 'leadingZero' as const, value })),
  fc.constantFrom('A1+1', 'SUM(B2:B4)', '"text"').map((formula) => ({ kind: 'formula' as const, formula })),
  fc.constantFrom('=not-a-formula', '+SUM(A1:A2)', '@user').map((value) => ({ kind: 'formulaLikeText' as const, value })),
  fc.boolean().map((value) => ({ kind: 'boolean' as const, value })),
  fc
    .constantFrom({ raw: '-12.5', value: -12.5 }, { raw: '0', value: 0 }, { raw: '125.50', value: 125.5 })
    .map((input) => ({ kind: 'number' as const, raw: input.raw, value: input.value })),
  fc
    .constantFrom(
      { raw: '$1,234.56', value: 1234.56 },
      { raw: '($987.65)', value: -987.65 },
      { raw: '12.5%', value: 0.125 },
      { raw: '-3.25%', value: -0.0325 },
    )
    .map((input) => ({ kind: 'accounting' as const, raw: input.raw, value: input.value })),
)

export const csvEdgeWorkbookSpecArbitrary: fc.Arbitrary<CsvEdgeWorkbookSpec> = fc
  .record({
    fileStem: fc.constantFrom('ledger', 'csvedge', 'import_fuzz'),
    delimiter: fc.constantFrom(',', ';', '\t'),
    newline: fc.constantFrom('\n', '\r\n'),
    rows: fc.array(fc.array(csvEdgeCellSpecArbitrary, { minLength: 1, maxLength: 5 }), { minLength: 1, maxLength: 5 }),
  })
  .map((spec) => ({
    fileStem: spec.fileStem,
    delimiter: spec.delimiter,
    newline: spec.newline,
    rows: [
      [
        { kind: 'plainText', value: 'Account ID' },
        { kind: 'plainText', value: 'Amount' },
        { kind: 'plainText', value: 'Formula Text' },
        { kind: 'plainText', value: 'Flag' },
        { kind: 'plainText', value: 'Comment' },
      ],
      [
        { kind: 'formulaLikeText', value: '=00123' },
        { kind: 'accounting', raw: '$1,234.56', value: 1234.56 },
        { kind: 'formula', formula: 'SUM(B2:B2)' },
        { kind: 'boolean', value: true },
        { kind: 'quotedText', value: `comma, semicolon; tab\tquote " newline\nunicode Ω` },
      ],
    ].concat(spec.rows),
  }))

export const malformedXlsxSpecArbitrary: fc.Arbitrary<MalformedXlsxSpec> = fc.oneof(
  fc.array(fc.integer({ min: 0, max: 255 }), { minLength: 0, maxLength: 32 }).map((byteSuffix) => ({
    kind: 'corruptZip' as const,
    byteSuffix,
  })),
  fc
    .record({
      entry: fc.constantFrom(
        {
          path: 'xl/workbook.xml',
          replacement: '<workbook><sheets><sheet name="Broken"',
        },
        {
          path: 'xl/worksheets/sheet1.xml',
          replacement: '<worksheet><sheetData><row r="1"><c r="A1"><v>1',
        },
        {
          path: 'xl/_rels/workbook.xml.rels',
          replacement: '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"',
        },
        {
          path: '[Content_Types].xml',
          replacement: '<Types><Override PartName="/xl/workbook.xml"',
        },
        {
          path: 'xl/styles.xml',
          replacement: '<styleSheet><fonts count="1"><font>',
        },
      ),
    })
    .map(({ entry }) => ({
      kind: 'malformedOpenXml' as const,
      path: entry.path,
      replacement: entry.replacement,
    })),
)

export function buildSemanticXlsxWorkbookSnapshot(spec: SemanticXlsxWorkbookSpec): WorkbookSnapshot {
  const noteText = `Note: ${spec.note}`
  const workbookProtection = spec.protectedAttributes
    ? {
        lockStructure: true,
        lockWindows: false,
        xmlAttributes: [
          { name: 'lockStructure', value: '1' },
          { name: 'lockWindows', value: '0' },
          { name: 'workbookPassword', value: 'AF2B' },
        ],
      }
    : undefined

  return {
    version: 1,
    workbook: {
      name: spec.workbookName,
      metadata: {
        ...(workbookProtection ? { workbookProtection } : {}),
        definedNames: [
          { name: 'InputAmount', value: { kind: 'cell-ref', sheetName: 'Data', address: 'B2' } },
          { name: 'StatusChoices', value: { kind: 'range-ref', sheetName: 'Lookup', startAddress: 'A2', endAddress: 'A4' } },
          { name: 'ReviewThreshold', value: { kind: 'scalar', value: 0.15 } },
        ],
        styles: workbookStyles,
        tables: [
          {
            name: 'DataTable',
            sheetName: 'Data',
            startAddress: 'A1',
            endAddress: 'D4',
            columnNames: ['Account', 'Amount', 'Status', 'Notes'],
            columns: [
              { name: 'Account' },
              { name: 'Amount', totalsRowFunction: 'sum' },
              { name: 'Status' },
              { name: 'Notes', totalsRowLabel: 'Reviewed' },
            ],
            headerRow: true,
            totalsRow: false,
            style: {
              name: 'TableStyleMedium9',
              showFirstColumn: false,
              showLastColumn: false,
              showRowStripes: true,
              showColumnStripes: false,
            },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        metadata: {
          sheetFormatPr: { defaultRowHeight: 14.6, customHeight: true },
          rows: semanticRows(spec.hiddenAxis),
          columns: semanticColumns(spec.hiddenAxis),
          styleRanges: [
            { range: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D1' }, styleId: 'header-style' },
            { range: { sheetName: 'Data', startAddress: 'B2', endAddress: 'B4' }, styleId: 'money-style' },
          ],
          commentThreads: [
            {
              threadId: 'thread:data:D2',
              sheetName: 'Data',
              address: 'D2',
              comments: [{ id: 'comment:data:D2:1', body: `Review ${spec.status}`, authorDisplayName: 'Finance' }],
            },
          ],
          validations: [
            {
              range: { sheetName: 'Data', startAddress: 'C2', endAddress: 'C4' },
              rule: { kind: 'list', source: { kind: 'range-ref', sheetName: 'Lookup', startAddress: 'A2', endAddress: 'A4' } },
              allowBlank: false,
              showDropdown: true,
              promptTitle: 'Status',
              promptMessage: 'Pick a tracked status.',
              errorStyle: 'warning',
              errorTitle: 'Unknown status',
              errorMessage: 'Use the lookup status list.',
            },
            {
              range: { sheetName: 'Data', startAddress: 'B2', endAddress: 'B4' },
              rule: { kind: 'decimal', operator: 'between', values: [-10_000, 10_000] },
              allowBlank: true,
            },
          ],
          hyperlinks: [
            {
              sheetName: 'Data',
              address: 'D2',
              target: 'https://example.com/review',
              tooltip: 'Open review',
              display: noteText,
            },
            {
              sheetName: 'Data',
              address: 'C2',
              target: '#Lookup!A2',
              tooltip: 'Jump to statuses',
              display: spec.status,
            },
          ],
          sheetProtection: spec.protectedAttributes
            ? {
                sheetName: 'Data',
                xmlAttributes: [
                  { name: 'selectLockedCells', value: '1' },
                  { name: 'algorithmName', value: 'fuzz&excel' },
                ],
              }
            : { sheetName: 'Data' },
          protectedRanges: [
            {
              id: 'protect-inputs',
              range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'D4' },
            },
          ],
          richTextArtifacts: {
            cells: [
              {
                address: 'D2',
                text: noteText,
                storage: spec.richTextStorage,
                xml: richTextXml(spec.richTextStorage, 'Note: ', spec.note),
              },
            ],
          },
        },
        cells: [
          { address: 'A1', value: 'Account' },
          { address: 'B1', value: 'Amount' },
          { address: 'C1', value: 'Status' },
          { address: 'D1', value: 'Notes' },
          { address: 'A2', value: spec.account, format: '@' },
          { address: 'B2', value: spec.amount, format: '$#,##0.00' },
          { address: 'C2', value: spec.status },
          { address: 'D2', value: noteText },
          { address: 'A3', value: 'Projected' },
          { address: 'B3', formula: 'B2*1.1', format: '$#,##0.00' },
          { address: 'C3', value: true },
          { address: 'D3', formula: 'CONCATENATE(A2,"-",C2)' },
          { address: 'A4', value: 'Total' },
          { address: 'B4', formula: 'SUM(B2:B3)', format: '$#,##0.00' },
          { address: 'C4', value: false },
          { address: 'D4', value: 'Complete' },
          { address: spec.blankFormatAddress, format: '@' },
        ],
      },
      {
        id: 2,
        name: 'Lookup',
        order: 1,
        metadata: {
          ...(spec.lookupVisibility === 'visible' ? {} : { visibility: spec.lookupVisibility }),
          rows: [{ id: 'row:0', index: 0, size: 22 }],
          columns: [{ id: 'col:0', index: 0, size: 126 }],
        },
        cells: [
          { address: 'A1', value: 'Status' },
          { address: 'A2', value: 'Open' },
          { address: 'A3', value: 'Closed' },
          { address: 'A4', value: 'Needs Review' },
        ],
      },
    ],
  }
}

export function semanticXlsxWorkbookDigest(snapshot: WorkbookSnapshot): unknown {
  const stylesById = new Map((snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
  const portableStyle = (styleId: string) => portableStyleDigest(stylesById.get(styleId))
  return {
    workbook: {
      definedNames: (snapshot.workbook.metadata?.definedNames ?? []).map(projectDefinedName).toSorted(compareByStableJson),
      workbookProtection: snapshot.workbook.metadata?.workbookProtection,
      tables: (snapshot.workbook.metadata?.tables ?? []).map((table) => structuredClone(table)).toSorted(compareByStableJson),
    },
    sheets: snapshot.sheets
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => ({
        name: sheet.name,
        order: sheet.order,
        cells: sheet.cells.map(projectCell).toSorted(compareCellsByAddress),
        metadata: {
          rows: normalizeAxisEntries(sheet.metadata?.rows),
          columns: normalizeAxisEntries(sheet.metadata?.columns),
          sheetFormatPr: sheet.metadata?.sheetFormatPr,
          visibility: sheet.metadata?.visibility,
          sheetProtection: sheet.metadata?.sheetProtection,
          protectedRanges: (sheet.metadata?.protectedRanges ?? []).map((entry) => structuredClone(entry)).toSorted(compareByStableJson),
          validations: (sheet.metadata?.validations ?? []).map((entry) => structuredClone(entry)).toSorted(compareByStableJson),
          hyperlinks: (sheet.metadata?.hyperlinks ?? []).map(projectHyperlink).toSorted(compareByStableJson),
          commentThreads: (sheet.metadata?.commentThreads ?? []).map((thread) => ({
            sheetName: thread.sheetName,
            address: thread.address,
            comments: thread.comments.map((comment) => ({
              body: comment.body,
              ...(comment.authorDisplayName !== undefined ? { authorDisplayName: comment.authorDisplayName } : {}),
            })),
          })),
          richTextArtifacts: (sheet.metadata?.richTextArtifacts?.cells ?? []).map((cell) => ({
            address: cell.address,
            text: cell.text,
            storage: cell.storage,
            xml: cell.xml,
          })),
          styleRanges: (sheet.metadata?.styleRanges ?? [])
            .map((range) => ({
              range: range.range,
              style: portableStyle(range.styleId),
            }))
            .toSorted(compareByStableJson),
        },
      })),
  }
}

export function renderCsvEdgeWorkbook(spec: CsvEdgeWorkbookSpec): string {
  return spec.rows.map((row) => row.map((cell) => escapeCsvCell(csvRawValue(cell), spec.delimiter)).join(spec.delimiter)).join(spec.newline)
}

export function expectedCsvEdgePreviewRows(spec: CsvEdgeWorkbookSpec): string[][] {
  const rawRows = spec.rows.map((row) => row.map(csvRawValue))
  const width = Math.max(...rawRows.map((row) => row.length))
  return rawRows.map((row) => row.concat(Array.from({ length: width - row.length }, () => '')))
}

export function expectedCsvEdgeCells(
  spec: CsvEdgeWorkbookSpec,
): Array<{ address: string; row: number; col: number; value?: unknown; formula?: string }> {
  const decimalSeparator = resolveExpectedCsvDecimalSeparator(spec)
  return spec.rows.flatMap((row, rowIndex) =>
    row.flatMap((cell, columnIndex) => {
      const raw = csvRawValue(cell)
      const parsed = parseExpectedCsvCell(raw, {
        decimalSeparator,
        forceText: columnIndex === 0 && rowIndex > 0 && raw.trim() !== '',
      })
      if (!parsed) {
        return []
      }
      return [{ address: XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex }), row: rowIndex, col: columnIndex, ...parsed }]
    }),
  )
}

export function buildMalformedXlsxBytes(spec: MalformedXlsxSpec): Uint8Array {
  if (spec.kind === 'corruptZip') {
    return Uint8Array.from([...new TextEncoder().encode('not an xlsx zip'), ...(spec.byteSuffix ?? [])])
  }
  const zip = unzipSync(
    exportXlsx({
      version: 1,
      workbook: { name: 'malformed-openxml' },
      sheets: [{ id: 1, name: 'Data', order: 0, cells: [{ address: 'A1', value: 'ok' }] }],
    }),
  )
  zip[spec.path ?? 'xl/workbook.xml'] = strToU8(spec.replacement ?? '<broken')
  return zipSync(zip)
}

function semanticRows(hiddenAxis: boolean): AxisEntrySnapshot[] {
  return [
    { id: 'row:0', index: 0, size: 26 },
    { id: 'row:1', index: 1, size: 20 },
    ...(hiddenAxis ? [{ id: 'row:6', index: 6, hidden: true }] : []),
  ]
}

function semanticColumns(hiddenAxis: boolean): AxisEntrySnapshot[] {
  return [
    { id: 'col:0', index: 0, size: 118 },
    { id: 'col:1', index: 1, size: 104 },
    { id: 'col:2', index: 2, size: 132 },
    { id: 'col:3', index: 3, size: 180 },
    ...(hiddenAxis ? [{ id: 'col:6', index: 6, hidden: true }] : []),
  ]
}

const workbookStyles: CellStyleSnapshot[] = [
  {
    id: 'header-style',
    fill: { backgroundColor: '#1d3989' },
    font: { bold: true, color: '#ffffff' },
    alignment: { horizontal: 'center', vertical: 'middle', wrap: true },
    borders: { bottom: { style: 'solid', weight: 'thin', color: '#000000' } },
  },
  {
    id: 'money-style',
    font: { color: '#1f4e79' },
    protection: { locked: false },
  },
]

function richTextXml(storage: 'sharedString' | 'inlineString', prefix: string, text: string): string {
  const itemName = storage === 'sharedString' ? 'si' : 'is'
  return [
    `<${itemName}>`,
    `<r><rPr><b/></rPr><t xml:space="preserve">${escapeXmlText(prefix)}</t></r>`,
    `<r><rPr><i/></rPr><t xml:space="preserve">${escapeXmlText(text)}</t></r>`,
    `</${itemName}>`,
  ].join('')
}

function escapeXmlText(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function portableStyleDigest(style: CellStyleSnapshot | undefined): Partial<CellStyleSnapshot> | undefined {
  if (!style) {
    return undefined
  }
  return {
    ...(style.fill ? { fill: style.fill } : {}),
    ...(style.font ? { font: style.font } : {}),
    ...(style.alignment ? { alignment: style.alignment } : {}),
    ...(style.borders ? { borders: style.borders } : {}),
    ...(style.protection ? { protection: style.protection } : {}),
  }
}

function projectCell(cell: SheetCellSnapshot): { address: string; value?: unknown; formula?: string; format?: string } {
  return {
    address: cell.address,
    ...(cell.value !== undefined ? { value: cell.value } : {}),
    ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
    ...(cell.format !== undefined ? { format: cell.format } : {}),
  }
}

function projectDefinedName(definedName: NonNullable<WorkbookMetadataSnapshot['definedNames']>[number]): unknown {
  return {
    name: definedName.name,
    ...(definedName.scopeSheetName !== undefined ? { scopeSheetName: definedName.scopeSheetName } : {}),
    value: structuredClone(definedName.value),
  }
}

function projectHyperlink(hyperlink: NonNullable<SheetMetadataSnapshot['hyperlinks']>[number]): unknown {
  return {
    sheetName: hyperlink.sheetName,
    address: hyperlink.address,
    target: hyperlink.target,
    ...(hyperlink.tooltip !== undefined ? { tooltip: decodeExcelEscapedText(decodeXmlEntities(hyperlink.tooltip)) } : {}),
    ...(hyperlink.display !== undefined ? { display: decodeExcelEscapedText(decodeXmlEntities(hyperlink.display)) } : {}),
  }
}

function decodeXmlEntities(value: string): string {
  return value.replace(/&(?:quot|apos|lt|gt|amp);/gu, (entity) => {
    switch (entity) {
      case '&quot;':
        return '"'
      case '&apos;':
        return "'"
      case '&lt;':
        return '<'
      case '&gt;':
        return '>'
      case '&amp;':
        return '&'
      default:
        return entity
    }
  })
}

function compareCellsByAddress(left: { address: string }, right: { address: string }): number {
  const leftCell = XLSX.utils.decode_cell(left.address)
  const rightCell = XLSX.utils.decode_cell(right.address)
  return leftCell.r - rightCell.r || leftCell.c - rightCell.c
}

function normalizeAxisEntries(entries: readonly AxisEntrySnapshot[] | undefined): AxisEntrySnapshot[] {
  return (entries ?? []).map(normalizedAxisEntry).toSorted((left, right) => left.index - right.index)
}

function normalizedAxisEntry(entry: AxisEntrySnapshot): AxisEntrySnapshot {
  const normalized: AxisEntrySnapshot = {
    id: entry.id,
    index: entry.index,
  }
  if (entry.size !== undefined && entry.size !== null) {
    normalized.size = entry.size
  }
  if (entry.hidden !== undefined && entry.hidden !== null) {
    normalized.hidden = entry.hidden
  }
  if (entry.customHeight !== undefined && entry.customHeight !== null) {
    normalized.customHeight = entry.customHeight
  }
  if (entry.customWidth !== undefined && entry.customWidth !== null) {
    normalized.customWidth = entry.customWidth
  }
  return normalized
}

function compareByStableJson(left: unknown, right: unknown): number {
  return stableJson(left).localeCompare(stableJson(right))
}

function stableJson(value: unknown): string {
  return JSON.stringify(value)
}

function csvRawValue(cell: CsvEdgeCellSpec): string {
  switch (cell.kind) {
    case 'empty':
      return ''
    case 'plainText':
    case 'quotedText':
    case 'leadingZero':
    case 'formulaLikeText':
      return cell.value
    case 'formula':
      return `=${cell.formula}`
    case 'boolean':
      return cell.value ? 'TRUE' : 'FALSE'
    case 'number':
    case 'accounting':
      return cell.raw
  }
}

function escapeCsvCell(value: string, delimiter: ',' | ';' | '\t'): string {
  return value.includes(delimiter) || /["\n\r]/u.test(value) ? `"${value.replaceAll('"', '""')}"` : value
}

function resolveExpectedCsvDecimalSeparator(spec: CsvEdgeWorkbookSpec): '.' | ',' {
  if (spec.delimiter === ',') {
    return '.'
  }
  return spec.rows.some((row) => row.some((cell) => /^-?\d+,\d+(%?)$/u.test(csvRawValue(cell).trim()))) ? ',' : '.'
}

function parseExpectedCsvCell(
  raw: string,
  options: { decimalSeparator: '.' | ','; forceText: boolean },
): { value?: unknown; formula?: string } | undefined {
  const normalized = raw.trim()
  if (normalized === '') {
    return undefined
  }
  if (options.forceText) {
    return { value: raw }
  }
  if (normalized.startsWith('=')) {
    return { formula: normalized.slice(1) }
  }
  if (normalized === 'TRUE' || normalized === 'FALSE') {
    return { value: normalized === 'TRUE' }
  }
  if (/^0\d+$/u.test(normalized)) {
    return { value: raw }
  }
  const accountingNumber = parseExpectedAccountingNumber(normalized, options.decimalSeparator)
  return accountingNumber === null ? { value: raw } : { value: accountingNumber }
}

function parseExpectedAccountingNumber(normalized: string, decimalSeparator: '.' | ','): number | null {
  let text = normalized
  let sign = 1

  if (text.startsWith('(') && text.endsWith(')')) {
    sign = -1
    text = text.slice(1, -1).trim()
  }
  if (text.startsWith('-')) {
    sign *= -1
    text = text.slice(1).trim()
  }
  if (text.startsWith('$')) {
    text = text.slice(1).trim()
  }
  if (text.startsWith('-')) {
    sign *= -1
    text = text.slice(1).trim()
  }

  const isPercent = text.endsWith('%')
  if (isPercent) {
    text = text.slice(0, -1).trim()
  }

  const groupSeparator = decimalSeparator === ',' ? '.' : ','
  const decimal = escapeRegExp(decimalSeparator)
  const group = escapeRegExp(groupSeparator)
  if (
    !new RegExp(`^\\d+(?:${decimal}\\d+)?$`, 'u').test(text) &&
    !new RegExp(`^\\d{1,3}(?:${group}\\d{3})+(?:${decimal}\\d+)?$`, 'u').test(text)
  ) {
    return null
  }

  const parsed = Number(text.replaceAll(groupSeparator, '').replace(decimalSeparator, '.'))
  return Number.isFinite(parsed) ? (sign * parsed) / (isPercent ? 100 : 1) : null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
