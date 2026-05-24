import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const drawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const drawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const chartContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

describe('macOS Desktop Excel chart deleted sheet oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel raw chart formula invalidation after deleting a referenced sheet',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-chart-delete-sheet-oracle-')
      try {
        const sourceBytes = buildUnsupportedChartSourceXlsx()
        expect(chartFormulaRefs(importXlsx(sourceBytes, 'chart-delete-source.xlsx').snapshot)).toEqual([
          'Data!$B$1',
          'Data!$A$2:$A$3',
          'Data!$B$2:$B$3',
        ])

        const excelWorkbookPath = join(tempDir, 'excel-chart-delete-source.xlsx')
        writeFileSync(excelWorkbookPath, sourceBytes)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Dashboard',
          operations: [{ kind: 'deleteSheet', name: 'Data' }],
          inspectCells: ['A1'],
          saveWorkbook: true,
          timeoutMs: 120_000,
        })
        expect(excelResult.cells[0]?.value).toEqual({ kind: 'string', value: 'dashboard' })

        const excelTruth = importXlsx(new Uint8Array(readFileSync(excelWorkbookPath)), 'excel-chart-delete-truth.xlsx').snapshot
        const excelTruthRefs = chartFormulaRefs(excelTruth)
        expect(excelTruth.sheets.map((sheet) => sheet.name)).toEqual(['Dashboard'])
        expect(excelTruthRefs).not.toContain('Data!$B$2:$B$3')

        const importedSource = importXlsx(sourceBytes, 'headless-chart-delete-source.xlsx').snapshot
        const workpaper = WorkPaper.buildFromSnapshot(importedSource)
        try {
          const dataSheet = workpaper.getSheetId('Data')
          if (dataSheet === undefined) {
            throw new Error('Expected Data sheet')
          }
          workpaper.removeSheet(dataSheet)

          const headlessSnapshot = workpaper.exportSnapshot()
          expect(headlessSnapshot.sheets.map((sheet) => sheet.name)).toEqual(['Dashboard'])
          expect(chartFormulaRefs(headlessSnapshot)).toEqual(excelTruthRefs)

          const headlessPath = join(tempDir, 'headless-chart-delete.xlsx')
          writeFileSync(headlessPath, exportXlsx(headlessSnapshot))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Dashboard',
            formulaCells: [],
            inspectCells: ['A1'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = importXlsx(new Uint8Array(readFileSync(headlessPath)), 'headless-chart-delete-truth.xlsx').snapshot
          expect(chartFormulaRefs(headlessTruth)).toEqual(excelTruthRefs)
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    180_000,
  )
})

function chartFormulaRefs(snapshot: WorkbookSnapshot): string[] {
  const parts = snapshot.workbook.metadata?.drawingArtifacts?.parts ?? []
  return parts
    .filter((part) => /^xl\/charts\/chart\d+\.xml$/u.test(part.path))
    .flatMap((part) => {
      const xml = strFromU8(decodeBase64(part.dataBase64))
      return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/gu)].map((match) => match[1] ?? '')
    })
}

function buildUnsupportedChartSourceXlsx(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Quarter', 'Revenue'],
      ['Q1', 10],
      ['Q2', 14],
    ]),
    'Data',
  )
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['dashboard']]), 'Dashboard')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet2.xml'] = strToU8(addWorksheetDrawing(readZipTextFromZip(zip, 'xl/worksheets/sheet2.xml'), 'rIdChartDrawing'))
  zip['xl/worksheets/_rels/sheet2.xml.rels'] = strToU8(
    relationshipsXml([{ id: 'rIdChartDrawing', type: drawingRelationshipType, target: '../drawings/drawing1.xml' }]),
  )
  zip['xl/drawings/drawing1.xml'] = strToU8(worksheetDrawingXml)
  zip['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(
    relationshipsXml([{ id: 'rId1', type: chartRelationshipType, target: '../charts/chart1.xml' }]),
  )
  zip['xl/charts/chart1.xml'] = strToU8(unsupportedChartXml)
  zip['[Content_Types].xml'] = strToU8(
    upsertContentTypeOverride(
      upsertContentTypeOverride(readZipTextFromZip(zip, '[Content_Types].xml'), {
        partName: '/xl/drawings/drawing1.xml',
        contentType: drawingContentType,
      }),
      { partName: '/xl/charts/chart1.xml', contentType: chartContentType },
    ),
  )
  return zipSync(zip)
}

function addWorksheetDrawing(sheetXml: string, relationshipId: string): string {
  const withRelationshipNamespace = /xmlns:r=/u.test(sheetXml)
    ? sheetXml
    : sheetXml.replace(
        /<worksheet\b([^>]*)>/u,
        `<worksheet$1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
      )
  return withRelationshipNamespace.replace('</worksheet>', `<drawing r:id="${relationshipId}"/></worksheet>`)
}

function relationshipsXml(relationships: readonly { id: string; type: string; target: string }[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) => `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`,
    ),
    '</Relationships>',
  ].join('')
}

function readZipTextFromZip(zip: Record<string, Uint8Array>, path: string): string {
  const bytes = zip[path]
  if (!bytes) {
    throw new Error(`Missing XLSX part: ${path}`)
  }
  return strFromU8(bytes)
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

function decodeBase64(dataBase64: string): Uint8Array {
  const binary = globalThis.atob(dataBase64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

const worksheetDrawingXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
  '<xdr:twoCellAnchor>',
  '<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
  '<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>',
  '<xdr:graphicFrame macro="">',
  '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Unsupported Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>',
  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>',
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">',
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${officeRelationshipNamespace}" r:id="rId1"/>`,
  '</a:graphicData></a:graphic>',
  '</xdr:graphicFrame><xdr:clientData/>',
  '</xdr:twoCellAnchor>',
  '</xdr:wsDr>',
].join('')

const unsupportedChartXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
  '<c:lang val="en-US"/>',
  '<c:chart><c:plotArea><c:layout/><c:doughnutChart>',
  '<c:varyColors val="1"/>',
  '<c:ser><c:idx val="0"/><c:order val="0"/>',
  '<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>',
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>',
  '<c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>14</c:v></c:pt></c:numCache></c:numRef></c:val>',
  '</c:ser>',
  '<c:firstSliceAng val="0"/><c:holeSize val="50"/>',
  '</c:doughnutChart></c:plotArea>',
  '<c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart>',
  '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>',
  '</c:chartSpace>',
].join('')
