import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'

import { exportXlsx, importXlsx } from '../index.js'

const printPageSetupElements = ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks'] as const

type PrintPageSetupElementName = (typeof printPageSetupElements)[number]

interface PrintPageSetupSummary {
  readonly definedNames: readonly DefinedNameSummary[]
  readonly sheets: ReadonlyArray<Record<PrintPageSetupElementName, string>>
}

interface DefinedNameSummary {
  readonly name: string | null
  readonly localSheetId: string | null
  readonly text: string
}

describe('print page setup roundtrip', () => {
  it('preserves worksheet print metadata and built-in print defined names across XLSX round trips', () => {
    const source = buildPrintPageSetupWorkbookBytes()
    const imported = importXlsx(source, 'print-page-setup.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.printPageSetup).toMatchObject({
      printOptionsXml: '<printOptions horizontalCentered="1" gridLines="1"/>',
      pageSetupXml: '<pageSetup paperSize="9" scale="60" orientation="landscape" horizontalDpi="4294967292" verticalDpi="4294967292"/>',
      headerFooterXml: '<headerFooter alignWithMargins="0"><oddHeader>&amp;A</oddHeader><oddFooter>Page &amp;P</oddFooter></headerFooter>',
    })
    expect(imported.snapshot.sheets[1]?.metadata?.printPageSetup).toMatchObject({
      pageMarginsXml: '<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>',
      pageSetupXml: '<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="0"/>',
    })

    const exported = exportXlsx(imported.snapshot)
    expect(readPrintPageSetupSummary(exported)).toEqual(readPrintPageSetupSummary(source))
  })
})

function buildPrintPageSetupWorkbookBytes(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Print One',
          cells: [
            cell('A1', 0, 0, 'Header'),
            cell('B1', 0, 1, 'Q1'),
            cell('C1', 0, 2, 'Q2'),
            cell('D1', 0, 3, 'Q3'),
            cell('A2', 1, 0, 'Revenue'),
            cell('B2', 1, 1, 10),
            cell('C2', 1, 2, 20),
            cell('D2', 1, 3, 30),
            cell('A3', 2, 0, 'Expense'),
            cell('B3', 2, 1, 4),
            cell('C3', 2, 2, 5),
            cell('D3', 2, 3, 6),
            cell('A4', 3, 0, 'Profit'),
            cell('B4', 3, 1, 6),
            cell('C4', 3, 2, 15),
            cell('D4', 3, 3, 24),
            cell('A5', 4, 0, 'Notes'),
            cell('A6', 5, 0, 'Footer'),
          ],
        },
        {
          name: 'Print Two',
          cells: [
            cell('A1', 0, 0, 'Metric'),
            cell('B1', 0, 1, 'Value'),
            cell('A2', 1, 0, 'Revenue'),
            cell('B2', 1, 1, 100),
            cell('A3', 2, 0, 'Cost'),
            cell('B3', 2, 1, 70),
            cell('A4', 3, 0, 'Margin'),
            cell('B4', 3, 1, 30),
          ],
        },
      ],
      definedNames: [
        { name: '_xlnm.Print_Area', localSheetIndex: 0, formula: "'Print One'!$A$1:$D$6" },
        { name: '_xlnm.Print_Titles', localSheetIndex: 0, formula: "'Print One'!$1:$2" },
      ],
    }),
  )
  replaceWorksheetPrintElements(zip, 1, [
    '<printOptions horizontalCentered="1" gridLines="1"/>',
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    '<pageSetup paperSize="9" scale="60" orientation="landscape" horizontalDpi="4294967292" verticalDpi="4294967292"/>',
    '<headerFooter alignWithMargins="0"><oddHeader>&amp;A</oddHeader><oddFooter>Page &amp;P</oddFooter></headerFooter>',
    '<rowBreaks count="1" manualBreakCount="1"><brk id="20" max="16383" man="1"/></rowBreaks>',
    '<colBreaks count="1" manualBreakCount="1"><brk id="4" max="1048575" man="1"/></colBreaks>',
  ])
  replaceWorksheetPrintElements(zip, 2, [
    '<pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>',
    '<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="0"/>',
    '<headerFooter differentOddEven="1"><oddFooter>Prepared &amp;D</oddFooter><evenFooter>Confidential</evenFooter></headerFooter>',
  ])
  return zipSync(zip)
}

function cell(
  address: string,
  row: number,
  col: number,
  value: string | number,
): { address: string; row: number; col: number; value: string | number } {
  return { address, row, col, value }
}

function replaceWorksheetPrintElements(zip: Record<string, Uint8Array>, sheetIndex: number, elements: readonly string[]): void {
  const sheetPath = `xl/worksheets/sheet${String(sheetIndex)}.xml`
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const withoutPrintElements = removePrintPageSetupElements(sheetXml)
  zip[sheetPath] = strToU8(withoutPrintElements.replace('</worksheet>', `${elements.join('')}</worksheet>`))
}

function removePrintPageSetupElements(sheetXml: string): string {
  return printPageSetupElements.reduce((xml, elementName) => xml.replace(elementPattern(elementName), ''), sheetXml)
}

function readPrintPageSetupSummary(bytes: Uint8Array): PrintPageSetupSummary {
  const zip = unzipSync(bytes)
  return {
    definedNames: readBuiltInPrintDefinedNames(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())),
    sheets: [1, 2].map((sheetIndex) => {
      const sheetXml = strFromU8(zip[`xl/worksheets/sheet${String(sheetIndex)}.xml`] ?? new Uint8Array())
      return {
        printOptions: readElementXml(sheetXml, 'printOptions'),
        pageMargins: readElementXml(sheetXml, 'pageMargins'),
        pageSetup: readElementXml(sheetXml, 'pageSetup'),
        headerFooter: readElementXml(sheetXml, 'headerFooter'),
        rowBreaks: readElementXml(sheetXml, 'rowBreaks'),
        colBreaks: readElementXml(sheetXml, 'colBreaks'),
      }
    }),
  }
}

function readBuiltInPrintDefinedNames(workbookXml: string): DefinedNameSummary[] {
  return [...workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const name = readAttribute(attributes, 'name')
    if (name !== '_xlnm.Print_Area' && name !== '_xlnm.Print_Titles') {
      return []
    }
    return [
      {
        name,
        localSheetId: readAttribute(attributes, 'localSheetId'),
        text: decodeDefinedNameText(match[2] ?? ''),
      },
    ]
  })
}

function decodeDefinedNameText(text: string): string {
  return text
    .replace(/&apos;/gu, "'")
    .replace(/&quot;/gu, '"')
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function readElementXml(sheetXml: string, elementName: PrintPageSetupElementName): string {
  return elementPattern(elementName).exec(sheetXml)?.[0] ?? ''
}

function elementPattern(elementName: PrintPageSetupElementName): RegExp {
  return new RegExp(`<${elementName}\\b[^>]*(?:/>|>[\\s\\S]*?</${elementName}>)`, 'gu')
}

function readAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\b${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}
