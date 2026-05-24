import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

import { isMacosExcelInstalled, runMacosExcelInspectionOracle } from '@bilig/excel-fixtures'
import { ValueTag } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper, exportXlsx, recalculateXlsx } from '../index.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const externalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
const externalRangeAddresses = ['C1', 'C2', 'C3'] as const

const staleFormulaCacheValues = {
  C1: '120',
  C2: '40',
  C3: '60',
} as const

const updatedFormulaCacheValues = {
  C1: '180',
  C2: '60',
  C3: '80',
} as const

describe('macOS Desktop Excel external-link cache recalc oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel updated external-link caches after formula recalculation',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-xlsx-recalc-external-link-cache-oracle-')
      try {
        const linkedSourcePath = join(tempDir, 'rates.xlsx')
        const sourcePath = join(tempDir, 'external-link-cache.xlsx')
        writeFileSync(linkedSourcePath, buildExternalSourceWorkbook([20, 30, 40]))

        const sourceBytes = buildExternalLinkRangeCacheWorkbook(pathToFileURL(linkedSourcePath).href)
        expect(worksheetFormulaCacheValues(sourceBytes)).toEqual(staleFormulaCacheValues)
        writeFileSync(sourcePath, sourceBytes)

        const excelTruth = runMacosExcelInspectionOracle({
          workbookPath: sourcePath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          companionWorkbookPaths: [linkedSourcePath],
          saveWorkbook: true,
          timeoutMs: 120_000,
          updateLinks: 'external',
        })
        expect(excelTruth.excelVersion).toMatch(/^\d+\./u)
        expect(excelTruth.cells.map(({ address, value }) => ({ address, value }))).toEqual([
          { address: 'C1', value: { kind: 'number', value: 180 } },
          { address: 'C2', value: { kind: 'number', value: 60 } },
          { address: 'C3', value: { kind: 'number', value: 80 } },
        ])

        const excelUpdatedBytes = new Uint8Array(readFileSync(sourcePath))
        expect(worksheetFormulaCacheValues(excelUpdatedBytes)).toEqual(updatedFormulaCacheValues)

        const recalculated = recalculateXlsx(excelUpdatedBytes, {
          fileName: 'external-link-cache-updated.xlsx',
          reads: externalRangeAddresses.map((address) => `Model!${address}`),
        })
        expect(
          Object.fromEntries(externalRangeAddresses.map((address) => [address, numberCell(recalculated.reads[`Model!${address}`])])),
        ).toEqual({
          C1: 180,
          C2: 60,
          C3: 80,
        })
        expect(worksheetFormulaCacheValues(recalculated.xlsx)).toEqual(updatedFormulaCacheValues)
        expect(externalLinkPackageSummary(unzipSync(recalculated.xlsx))).toEqual(externalLinkPackageSummary(unzipSync(excelUpdatedBytes)))

        const recalculatedPath = join(tempDir, 'external-link-cache-recalculated.xlsx')
        writeFileSync(recalculatedPath, recalculated.xlsx)
        const excelReopened = runMacosExcelInspectionOracle({
          workbookPath: recalculatedPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...externalRangeAddresses],
          calculationPolicy: 'none',
          saveWorkbook: false,
          timeoutMs: 120_000,
          updateLinks: 'never',
        })
        expect(excelReopened.cells.map(({ address, value }) => ({ address, value }))).toEqual(
          excelTruth.cells.map(({ address, value }) => ({ address, value })),
        )
        expect(excelReopened.cells.map(({ address, formula }) => ({ address, formula }))).toEqual([
          { address: 'C1', formula: '=SUM(__bilig_ext_1_Rates!$B$2:$B$4)*B1' },
          { address: 'C2', formula: '=XLOOKUP("B",__bilig_ext_1_Rates!$A$2:$A$4,__bilig_ext_1_Rates!$B$2:$B$4)*B1' },
          { address: 'C3', formula: '=SUMIFS(__bilig_ext_1_Rates!$B$2:$B$4,__bilig_ext_1_Rates!$A$2:$A$4,"C")*B1' },
        ])
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    180_000,
  )
})

function buildExternalSourceWorkbook(rates: readonly [number, number, number]): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Rates: [
      ['SKU', 'Rate'],
      ['A', rates[0]],
      ['B', rates[1]],
      ['C', rates[2]],
    ],
  })
  try {
    return exportXlsx(workbook.exportSnapshot())
  } finally {
    workbook.dispose()
  }
}

function buildExternalLinkRangeCacheWorkbook(target: string): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [
      [null, 2, 120],
      [null, null, 40],
      [null, null, 60],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      xmlText(zip, 'xl/worksheets/sheet1.xml')
        .replace(/<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>120</v></c>')
        .replace(
          /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
          "<c r=\"C2\"><f>_xlfn.XLOOKUP(&quot;B&quot;,'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1</f><v>40</v></c>",
        )
        .replace(
          /<c\b[^>]*\br=(["'])C3\1[^>]*>[\s\S]*?<\/c>/u,
          "<c r=\"C3\"><f>SUMIFS('[1]Rates'!$B$2:$B$4,'[1]Rates'!$A$2:$A$4,&quot;C&quot;)*B1</f><v>60</v></c>",
        ),
    )
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
  } finally {
    workbook.dispose()
  }
}

function numberCell(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'tag' in value && value.tag === ValueTag.Number && 'value' in value) {
    return Number(value.value)
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}

function createExcelAccessibleTempDir(prefix: string): string {
  const root = join(tmpdir(), 'bilig-xlsx-recalc-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

function worksheetFormulaCacheValues(bytes: Uint8Array): Readonly<Record<string, string | null>> {
  const sheetXml = xmlText(unzipSync(bytes), 'xl/worksheets/sheet1.xml')
  return Object.fromEntries(
    externalRangeAddresses.map((address) => {
      const cellXml = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)?.[0]
      if (!cellXml) {
        throw new Error(`Missing formula cell ${address}`)
      }
      return [address, /<v>([\s\S]*?)<\/v>/u.exec(cellXml)?.[1] ?? null]
    }),
  )
}

function externalLinkPackageSummary(zip: Record<string, Uint8Array>): {
  readonly packageParts: readonly (readonly [path: string, xml: string])[]
  readonly workbookExternalReferenceTargets: readonly string[]
  readonly externalLinkPathRelationshipParts: readonly (readonly [path: string, xml: string])[]
  readonly contentTypeOverrides: readonly string[]
} {
  const workbookExternalReferencesXml = extractSingleXml(
    xmlText(zip, 'xl/workbook.xml'),
    /<(?:[A-Za-z_][\w.-]*:)?externalReferences\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?externalReferences>/u,
  )
  const workbookExternalLinkTargets = workbookExternalLinkTargetsByRelationshipId(xmlText(zip, 'xl/_rels/workbook.xml.rels'))
  return {
    packageParts: Object.entries(zip)
      .filter(([path]) => /^xl\/externalLinks\/externalLink[^/]*\.xml$/u.test(path))
      .map(([path, bytes]) => [path, strFromU8(bytes)] as const)
      .toSorted(([left], [right]) => left.localeCompare(right)),
    workbookExternalReferenceTargets: extractCaptureMatches(workbookExternalReferencesXml, /\br:id="([^"]+)"/gu).map((id) => {
      const target = workbookExternalLinkTargets.get(id)
      if (target === undefined) {
        throw new Error(`Missing workbook external-link relationship target for ${id}`)
      }
      return target
    }),
    externalLinkPathRelationshipParts: Object.entries(zip)
      .filter(([path]) => /^xl\/externalLinks\/_rels\/externalLink[^/]*\.xml\.rels$/u.test(path))
      .map(([path, bytes]) => [path, strFromU8(bytes)] as const)
      .toSorted(([left], [right]) => left.localeCompare(right)),
    contentTypeOverrides: extractXmlMatches(
      xmlText(zip, '[Content_Types].xml'),
      /<Override\b[^>]*ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.externalLink\+xml"[^>]*\/>/gu,
    ),
  }
}

function workbookExternalLinkTargetsByRelationshipId(workbookRelationshipsXml: string): ReadonlyMap<string, string> {
  const targetsById = new Map<string, string>()
  for (const relationship of extractXmlMatches(workbookRelationshipsXml, /<Relationship\b[^>]*\/>/gu)) {
    if (!relationship.includes('/relationships/externalLink"')) {
      continue
    }
    targetsById.set(extractAttribute(relationship, 'Id'), extractAttribute(relationship, 'Target'))
  }
  if (targetsById.size === 0) {
    throw new Error('Missing workbook external-link relationships')
  }
  return targetsById
}

function extractSingleXml(xml: string, pattern: RegExp): string {
  const match = xml.match(pattern)
  if (!match) {
    throw new Error(`Missing expected XLSX XML fragment for pattern ${String(pattern)}`)
  }
  return match[0]
}

function extractXmlMatches(xml: string, pattern: RegExp): readonly string[] {
  const matches = [...xml.matchAll(pattern)].map((match) => match[0])
  if (matches.length === 0) {
    throw new Error(`Missing expected XLSX XML fragment for pattern ${String(pattern)}`)
  }
  return matches.toSorted()
}

function extractCaptureMatches(xml: string, pattern: RegExp): readonly string[] {
  const matches = [...xml.matchAll(pattern)].flatMap((match) => match[1] ?? [])
  if (matches.length === 0) {
    throw new Error(`Missing expected XLSX XML capture for pattern ${String(pattern)}`)
  }
  return matches
}

function extractAttribute(xml: string, attributeName: string): string {
  const match = new RegExp(`\\b${attributeName}="([^"]+)"`, 'u').exec(xml)
  if (!match) {
    throw new Error(`Missing ${attributeName} attribute in ${xml}`)
  }
  return match[1]
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
