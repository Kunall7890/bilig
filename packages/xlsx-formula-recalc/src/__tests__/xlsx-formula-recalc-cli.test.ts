import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { runXlsxFormulaRecalcCli } from '../cli-api.js'
import { WorkPaper, exportXlsx } from '../index.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('xlsx-recalc CLI', () => {
  it('runs a one-command demo that writes a recalculated XLSX and prints proof JSON', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-'))
    try {
      const outputPath = join(tempDir, 'demo.recalculated.xlsx')
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli(['--demo', '--out', outputPath, '--json'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.mode).toBe('demo')
      expect(summary.verified).toBe(true)
      expect(summary.reads['Summary!B2']?.value).toBe(72_000)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('hydrates external-link caches from companion workbook paths', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-external-'))
    try {
      const inputPath = join(tempDir, 'model.xlsx')
      const companionPath = join(tempDir, 'uploaded-rates.xlsx')
      const outputPath = join(tempDir, 'model.recalculated.xlsx')
      writeFileSync(inputPath, buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx'))
      writeFileSync(companionPath, buildExternalSourceWorkbook([20, 30, 40]))
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli(
        [
          inputPath,
          '--external-workbook-target',
          companionPath,
          'file:///tmp/rates.xlsx',
          '--read',
          'Model!C1',
          '--read',
          'Model!C2',
          '--out',
          outputPath,
          '--json',
        ],
        {
          stdout: (text) => {
            stdout += text
          },
        },
      )

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.externalWorkbooks).toBe(1)
      expect(summary.reads['Model!C1']?.value).toBe(180)
      expect(summary.reads['Model!C2']?.value).toBe(60)
      expect(summary.diagnostics?.externalWorkbookHydration).toMatchObject({
        externalWorkbookCount: 1,
        refreshedBookIndices: [1],
        refreshedCellCount: 6,
        references: [
          expect.objectContaining({
            status: 'refreshed',
            matchKind: 'exact-target',
            matchedFileName: 'uploaded-rates.xlsx',
            matchedTarget: 'file:///tmp/rates.xlsx',
          }),
        ],
      })
      expect(readExternalLinkCacheCellValue(readFileBytes(outputPath), 'B2')).toBe('20')
      expect(readExternalLinkCacheCellValue(readFileBytes(outputPath), 'B4')).toBe('40')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('preserves cached external-link values when CLI companion workbook paths are ambiguous', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-ambiguous-'))
    try {
      const inputPath = join(tempDir, 'model.xlsx')
      const firstCompanionPath = join(tempDir, 'one', 'rates.xlsx')
      const secondCompanionPath = join(tempDir, 'two', 'rates.xlsx')
      const outputPath = join(tempDir, 'model.recalculated.xlsx')
      mkdirSync(join(tempDir, 'one'))
      mkdirSync(join(tempDir, 'two'))
      writeFileSync(inputPath, buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx'))
      writeFileSync(firstCompanionPath, buildExternalSourceWorkbook([20, 30, 40]))
      writeFileSync(secondCompanionPath, buildExternalSourceWorkbook([200, 300, 400]))
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli(
        [
          inputPath,
          '--external-workbook',
          firstCompanionPath,
          '--external-workbook',
          secondCompanionPath,
          '--read',
          'Model!C1',
          '--out',
          outputPath,
          '--json',
        ],
        {
          stdout: (text) => {
            stdout += text
          },
        },
      )

      expect(exitCode).toBe(0)
      const summary = readCliSummary(stdout)
      expect(summary.externalWorkbooks).toBe(2)
      expect(summary.reads['Model!C1']?.value).toBe(120)
      expect(summary.warnings).toContain(
        'Some supplied external workbook companions matched ambiguously; existing external-link cache values were preserved.',
      )
      expect(summary.diagnostics?.externalWorkbookHydration).toMatchObject({
        skippedAmbiguousMatchCount: 1,
        references: [
          expect.objectContaining({
            status: 'skipped-ambiguous-match',
            candidateCount: 2,
          }),
        ],
      })
      expect(readExternalLinkCacheCellValue(readFileBytes(outputPath), 'B2')).toBe('10')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

interface CliSummary {
  readonly mode: string
  readonly externalWorkbooks: number
  readonly reads: Readonly<Record<string, { readonly value: unknown }>>
  readonly warnings: readonly string[]
  readonly diagnostics?: {
    readonly externalWorkbookHydration?: Record<string, unknown>
  }
  readonly verified: boolean
}

function readCliSummary(stdout: string): CliSummary {
  const parsed: unknown = JSON.parse(stdout)
  if (typeof parsed !== 'object' || parsed === null) {
    throw new Error(`Expected CLI summary object, received ${stdout}`)
  }
  const mode = Reflect.get(parsed, 'mode')
  const externalWorkbooks = Reflect.get(parsed, 'externalWorkbooks')
  const reads = Reflect.get(parsed, 'reads')
  const warnings = Reflect.get(parsed, 'warnings')
  const diagnostics = Reflect.get(parsed, 'diagnostics')
  const verified = Reflect.get(parsed, 'verified')
  if (
    typeof mode !== 'string' ||
    typeof externalWorkbooks !== 'number' ||
    typeof reads !== 'object' ||
    reads === null ||
    !Array.isArray(warnings) ||
    typeof verified !== 'boolean'
  ) {
    throw new Error(`Unexpected CLI summary shape: ${stdout}`)
  }
  const parsedDiagnostics = readCliSummaryDiagnostics(diagnostics)
  return {
    mode,
    externalWorkbooks,
    reads: readCliSummaryReads(reads),
    warnings: warnings.filter((warning): warning is string => typeof warning === 'string'),
    ...(parsedDiagnostics ? { diagnostics: parsedDiagnostics } : {}),
    verified,
  }
}

function readCliSummaryDiagnostics(value: unknown): CliSummary['diagnostics'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const externalWorkbookHydration = value['externalWorkbookHydration']
  return isRecord(externalWorkbookHydration) ? { externalWorkbookHydration } : undefined
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readCliSummaryReads(value: object): CliSummary['reads'] {
  const reads: Record<string, { readonly value: unknown }> = {}
  for (const [target, cellValue] of Object.entries(value)) {
    if (typeof cellValue !== 'object' || cellValue === null || !Reflect.has(cellValue, 'value')) {
      throw new Error(`Unexpected CLI read value for ${target}`)
    }
    reads[target] = {
      value: Reflect.get(cellValue, 'value'),
    }
  }
  return reads
}

function readFileBytes(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path))
}

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

function readExternalLinkCacheCellValue(bytes: Uint8Array, address: string): string | null {
  const xml = strFromU8(unzipSync(bytes)['xl/externalLinks/externalLink5.xml'] ?? new Uint8Array())
  const match = new RegExp(`<cell\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/cell>`, 'u').exec(xml)
  return match?.[1] ?? null
}
