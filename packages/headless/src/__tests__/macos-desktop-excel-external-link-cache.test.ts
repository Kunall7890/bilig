import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx, externalWorkbookReferencesWarning, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, type NormalizedFormulaValue } from '@bilig/excel-fixtures'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const externalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
const externalRangeAddresses = ['C1', 'C2', 'C3'] as const
const initialExternalRangeValues = [
  { address: 'C1', value: { kind: 'number', value: 120 } },
  { address: 'C2', value: { kind: 'number', value: 40 } },
  { address: 'C3', value: { kind: 'number', value: 60 } },
] as const
const changedExternalRangeValues = [
  { address: 'C1', value: { kind: 'number', value: 180 } },
  { address: 'C2', value: { kind: 'number', value: 60 } },
  { address: 'C3', value: { kind: 'number', value: 90 } },
] as const

describe('macOS Desktop Excel external-link cache oracle', () => {
  it('imports cached external ranges as hidden-sheet references', async () => {
    const imported = importXlsx(buildExternalLinkRangeCacheWorkbook(), 'external-link-range-cache.xlsx')
    const cells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')

    expect(cacheSheet).toMatchObject({ metadata: { visibility: 'veryHidden' } })
    expect(cells.get('C1')).toMatchObject({ formula: "SUM('__bilig_ext_1_Rates'!$B$2:$B$4)*B1", value: 120 })
    expect(cells.get('C2')).toMatchObject({
      formula: "XLOOKUP(\"B\",'__bilig_ext_1_Rates'!$A$2:$A$4,'__bilig_ext_1_Rates'!$B$2:$B$4)*B1",
      value: 40,
    })
    expect(cells.get('C3')).toMatchObject({
      formula: "SUMIFS('__bilig_ext_1_Rates'!$B$2:$B$4,'__bilig_ext_1_Rates'!$A$2:$A$4,\"C\")*B1",
      value: 60,
    })
    expect(imported.warnings).toEqual([externalWorkbookReferencesWarning])

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-range-cache-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expectEngineValues(engine, initialExternalRangeValues)

    engine.setCellValue('Model', 'B1', 3)

    expectEngineValues(engine, changedExternalRangeValues)

    const exportedZip = unzipSync(exportXlsx(engine.exportSnapshot()))
    const modelSheetXml = xmlText(exportedZip, 'xl/worksheets/sheet1.xml')
    const workbookXml = xmlText(exportedZip, 'xl/workbook.xml')

    expect(modelSheetXml).toContain('__bilig_ext_1_Rates')
    expect(modelSheetXml).not.toContain('{')
    expect(workbookXml).toContain('name="__bilig_ext_1_Rates"')
    expect(workbookXml).toContain('state="veryHidden"')
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'round-trips cached external ranges through Desktop Excel and Bilig recalc',
    async () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-external-link-range-cache-')
      try {
        const linkedSourceWorkbookPath = join(tempDir, 'rates.xlsx')
        const sourceWorkbookPath = join(tempDir, 'external-link-range-cache.xlsx')
        writeFileSync(linkedSourceWorkbookPath, buildExternalSourceWorkbook())
        writeFileSync(sourceWorkbookPath, buildExternalLinkRangeCacheWorkbook(pathToFileURL(linkedSourceWorkbookPath).href))

        const excelInitial = runMacosExcelInspectionOracle({
          workbookPath: sourceWorkbookPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          companionWorkbookPaths: [linkedSourceWorkbookPath],
          saveWorkbook: true,
          updateLinks: 'external',
        })
        expect(excelInitial.cells.map(({ address, value }) => ({ address, value }))).toEqual(initialExternalRangeValues)

        const imported = importXlsx(new Uint8Array(readFileSync(sourceWorkbookPath)), 'external-link-range-cache-saved.xlsx')
        const engine = new SpreadsheetEngine({ workbookName: 'external-link-range-cache-oracle' })
        await engine.ready()
        engine.importSnapshot(imported.snapshot)
        engine.setCellValue('Model', 'B1', 3)

        expectEngineValues(engine, changedExternalRangeValues)

        const exportedWorkbookPath = join(tempDir, 'external-link-range-cache-materialized.xlsx')
        writeFileSync(exportedWorkbookPath, exportXlsx(engine.exportSnapshot()))
        const excelChanged = runMacosExcelInspectionOracle({
          workbookPath: exportedWorkbookPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          saveWorkbook: false,
        })

        expect(excelChanged.cells.map(({ address, value }) => ({ address, value }))).toEqual(changedExternalRangeValues)
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    60_000,
  )
})

function expectEngineValues(
  engine: SpreadsheetEngine,
  expectedValues: readonly { readonly address: string; readonly value: NormalizedFormulaValue }[],
): void {
  expect(expectedValues.map(({ address }) => ({ address, value: normalizedCellValue(engine.getCellValue('Model', address)) }))).toEqual(
    expectedValues,
  )
}

function normalizedCellValue(value: CellValue): NormalizedFormulaValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return { kind: 'blank' }
    case ValueTag.Boolean:
      return { kind: 'boolean', value: value.value }
    case ValueTag.Error:
      return { kind: 'error', value: String(value.code) }
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.String:
      return { kind: 'string', value: value.value }
  }
}

function buildExternalSourceWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['SKU', 'Rate'],
    ['A', 10],
    ['B', 20],
    ['C', 30],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Rates')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function createExcelAccessibleTempDir(prefix: string): string {
  const root = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp/bilig-headless-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

function buildExternalLinkRangeCacheWorkbook(target = 'file:///tmp/rates.xlsx'): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[null, 2]])
  sheet.C1 = { t: 'n', f: "SUM('[1]Rates'!$B$2:$B$4)*B1", v: 120 }
  sheet.C2 = { t: 'n', f: "_xlfn.XLOOKUP(\"B\",'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1", v: 40 }
  sheet.C3 = { t: 'n', f: "SUMIFS('[1]Rates'!$B$2:$B$4,'[1]Rates'!$A$2:$A$4,\"C\")*B1", v: 60 }
  sheet['!ref'] = 'A1:C3'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/workbook.xml'] = strToU8(
    ensureRelationshipNamespace(xmlText(zip, 'xl/workbook.xml')).replace(
      '</sheets>',
      '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
    ),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    xmlText(zip, 'xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="Rates"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="1"><cell r="A1" t="str"><v>SKU</v></cell><cell r="B1" t="str"><v>Rate</v></cell></row>',
      '<row r="2"><cell r="A2" t="str"><v>A</v></cell><cell r="B2"><v>10</v></cell></row>',
      '<row r="3"><cell r="A3" t="str"><v>B</v></cell><cell r="B3"><v>20</v></cell></row>',
      '<row r="4"><cell r="A4" t="str"><v>C</v></cell><cell r="B4"><v>30</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
      '</Relationships>',
  )
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(xmlText(zip, '[Content_Types].xml'), {
      partName: '/xl/externalLinks/externalLink5.xml',
      contentType: externalLinkContentType,
    }),
  )
  return zipSync(zip)
}

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<workbook\b([^>]*)>/u, `<workbook$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function upsertContentTypeOverride(
  contentTypesXml: string,
  input: {
    readonly partName: string
    readonly contentType: string
  },
): string {
  if (contentTypesXml.includes(`PartName="${input.partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${input.partName}" ContentType="${input.contentType}"/></Types>`)
}

function xmlText(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
}
