import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'

import { exportXlsx, importXlsx } from '../index.js'

const protectedRangeSecurityAttributes = [
  { name: 'password', value: 'AF2B' },
  { name: 'algorithmName', value: 'SHA-512' },
  { name: 'hashValue', value: 'QUJDREVGRw==' },
  { name: 'saltValue', value: 'SElKS0w=' },
  { name: 'spinCount', value: '100000' },
  { name: 'securityDescriptor', value: 'user&group' },
]

describe('XLSX protected range attribute roundtrip', () => {
  it('preserves protectedRange security attributes while keeping Bilig range semantics authoritative', () => {
    const imported = importXlsx(buildProtectedRangeAttributesWorkbookBytes(), 'protected-range-security.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.protectedRanges).toEqual([
      {
        id: 'EditableInputs',
        range: { sheetName: 'Protected', startAddress: 'B2', endAddress: 'C3' },
        xmlAttributes: protectedRangeSecurityAttributes,
      },
    ])

    const exportedSheetXml = sheetXml(exportXlsx(imported.snapshot))
    expect(protectedRangeAttributes(exportedSheetXml)).toEqual([
      {
        name: 'EditableInputs',
        sqref: 'B2:C3',
        ...Object.fromEntries(protectedRangeSecurityAttributes.map((attribute) => [attribute.name, attribute.value])),
      },
    ])

    const engine = new SpreadsheetEngine({ workbookName: 'protected-range-security-engine' })
    engine.importSnapshot(imported.snapshot)
    const exportedFromEngineSheetXml = sheetXml(exportXlsx(engine.exportSnapshot()))
    expect(protectedRangeAttributes(exportedFromEngineSheetXml)).toEqual(protectedRangeAttributes(exportedSheetXml))

    const rewrittenSnapshot = structuredClone(imported.snapshot)
    const rewrittenSheet = rewrittenSnapshot.sheets[0]
    if (!rewrittenSheet?.metadata?.protectedRanges) {
      throw new Error('Expected imported protected range metadata')
    }
    rewrittenSheet.metadata.protectedRanges[0] = {
      id: 'EditableInputsAfterEdit',
      range: { sheetName: 'Protected', startAddress: 'D4', endAddress: 'E5' },
      xmlAttributes: [{ name: 'name', value: 'StaleName' }, { name: 'sqref', value: 'A1' }, ...protectedRangeSecurityAttributes],
    }
    expect(protectedRangeAttributes(sheetXml(exportXlsx(rewrittenSnapshot)))).toEqual([
      {
        name: 'EditableInputsAfterEdit',
        sqref: 'D4:E5',
        ...Object.fromEntries(protectedRangeSecurityAttributes.map((attribute) => [attribute.name, attribute.value])),
      },
    ])
  })
})

function buildProtectedRangeAttributesWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Label', 'Input A', 'Input B'],
    ['North', 10, 20],
    ['South', 30, 40],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Protected')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    sourceSheetXml.replace(
      '</sheetData>',
      [
        '</sheetData>',
        '<sheetProtection sheet="1"/>',
        '<protectedRanges>',
        '<protectedRange name="EditableInputs" sqref="B2:C3" password="AF2B" algorithmName="SHA-512" hashValue="QUJDREVGRw==" saltValue="SElKS0w=" spinCount="100000" securityDescriptor="user&amp;group"/>',
        '</protectedRanges>',
      ].join(''),
    ),
  )
  return zipSync(zip)
}

function sheetXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
}

function protectedRangeAttributes(xml: string): Array<Record<string, string>> {
  return [...xml.matchAll(/<protectedRange\b[^>]*\/>/gu)].map((match) => {
    const attributes: Record<string, string> = {}
    for (const attribute of match[0].matchAll(/\s([A-Za-z_:][\w:.-]*)="([^"]*)"/gu)) {
      attributes[attribute[1] ?? ''] = unescapeXml(attribute[2] ?? '')
    }
    return attributes
  })
}

function unescapeXml(value: string): string {
  return value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&')
}
