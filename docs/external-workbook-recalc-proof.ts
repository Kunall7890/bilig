import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, recalculateXlsx, WorkPaper } from '@bilig/xlsx-formula-recalc'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const outputDir = join(process.cwd(), 'bilig-external-workbook-proof-output')
const modelPath = join(outputDir, 'model-with-stale-external-cache.xlsx')
const ratesPath = join(outputDir, 'rates-current.xlsx')
const editedPath = join(outputDir, 'model-recalculated.xlsx')
const externalTarget = 'file:///bilig-proof/rates.xlsx'

mkdirSync(outputDir, { recursive: true })

const modelWithStaleCache = buildExternalLinkRangeCacheWorkbook(externalTarget)
const currentRatesWorkbook = buildExternalSourceWorkbook([20, 30, 40])
writeFileSync(modelPath, modelWithStaleCache)
writeFileSync(ratesPath, currentRatesWorkbook)

const recalculated = recalculateXlsx(modelWithStaleCache, {
  fileName: 'model-with-stale-external-cache.xlsx',
  externalWorkbooks: [
    {
      bytes: currentRatesWorkbook,
      fileName: 'rates-current.xlsx',
      target: externalTarget,
    },
  ],
  reads: ['Model!C1', 'Model!C2'],
})
writeFileSync(editedPath, recalculated.xlsx)

const hydration = readRecord(recalculated.diagnostics?.externalWorkbookHydration, 'external workbook hydration diagnostics')
const refreshedBookIndices = readNumberArray(hydration.refreshedBookIndices, 'refreshedBookIndices')
const sum = readNumberCell(recalculated.reads['Model!C1'], 'Model!C1')
const lookup = readNumberCell(recalculated.reads['Model!C2'], 'Model!C2')
const output = {
  proof:
    'Bilig refreshed an XLSX external-link cache from a companion workbook, recalculated formulas, and wrote a new XLSX without Excel.',
  verified: false,
  sum,
  lookup,
  files: {
    modelWithStaleCache: modelPath,
    companionWorkbook: ratesPath,
    recalculatedWorkbook: editedPath,
  },
  externalTarget,
  reads: recalculated.reads,
  diagnostics: {
    externalWorkbookHydration: hydration,
  },
  checks: {
    externalWorkbookMatched: refreshedBookIndices.includes(1),
    refreshedExternalCells: Number(hydration.refreshedCellCount) === 6,
    recalculatedExternalSum: sum === 180,
    recalculatedExternalLookup: lookup === 60,
    outputXlsxWritten: recalculated.xlsx.byteLength > 0,
    verified: false,
  },
  star: 'https://github.com/proompteng/bilig/stargazers',
  watchReleases: 'https://github.com/proompteng/bilig/subscription',
  adoptionBlocker: 'https://github.com/proompteng/bilig/discussions/new?category=general',
  nextStep:
    'If external workbook cache refresh is the XLSX blocker in your service, star or bookmark Bilig; if the matching rules are not enough, open the exact workbook-link blocker.',
}
output.checks.verified =
  output.checks.externalWorkbookMatched &&
  output.checks.refreshedExternalCells &&
  output.checks.recalculatedExternalSum &&
  output.checks.recalculatedExternalLookup &&
  output.checks.outputXlsxWritten
output.verified = output.checks.verified

if (!output.checks.verified) {
  throw new Error(`External workbook recalculation proof failed: ${JSON.stringify(output, null, 2)}`)
}

console.log(JSON.stringify(output, null, 2))

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
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
        .replace(/<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>120</v></c>')
        .replace(
          /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
          "<c r=\"C2\"><f>_xlfn.XLOOKUP(&quot;B&quot;,'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1</f><v>40</v></c>",
        ),
    )
    zip['xl/workbook.xml'] = strToU8(
      ensureRelationshipNamespace(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())).replace(
        '</sheets>',
        '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
      ),
    )
    zip['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
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
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function ensureRelationshipNamespace(xml: string): string {
  return xml.replace(/<workbook\b([^>]*)>/u, (match) =>
    match.includes('xmlns:r=') ? match : match.replace('>', ` xmlns:r="${officeRelationshipNamespace}">`),
  )
}

function readNumberCell(cell: unknown, target: string): number {
  const record = readRecord(cell, target)
  if (typeof record.value !== 'number') {
    throw new Error(`Expected numeric ${target}, got ${JSON.stringify(cell)}`)
  }
  return record.value
}

function readNumberArray(value: unknown, label: string): readonly number[] {
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'number')) {
    throw new Error(`Expected numeric array ${label}, got ${JSON.stringify(value)}`)
  }
  return value
}

function readRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`Expected object ${label}, got ${JSON.stringify(value)}`)
  }
  return value
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
