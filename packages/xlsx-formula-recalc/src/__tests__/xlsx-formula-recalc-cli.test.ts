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
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.recalculationCompleted).toBe(true)
      expect(summary.expectedValueMatched).toBe(true)
      expect(summary.expectedReadback).toEqual({ 'Summary!B2': 72_000 })
      expect(summary.excelParity).toBe('not_proven')
      expect(summary).not.toHaveProperty('star')
      expect(summary).not.toHaveProperty('watchReleases')
      expect(summary).not.toHaveProperty('adoptionBlocker')
      expect(summary.reads['Summary!B2']?.value).toBe(72_000)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('inspects workbook formula cells and stale cached formula values before writing an output file', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-inspect-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--inspect', '--json'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(join(tempDir, 'stale-cache.recalculated.xlsx'))).toBe(false)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.mode).toBe('file')
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.inspectionCompleted).toBe(true)
      expect(summary.recalculationCompleted).toBe(true)
      expect(summary.excelParity).toBe('not_proven')
      expect(summary.formulaCellCount).toBe(1)
      expect(summary.inspectedFormulaCellCount).toBe(1)
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.suggestedReads).toEqual(['Sheet1!B2'])
      expect(summary.formulas[0]).toMatchObject({
        target: 'Sheet1!B2',
        formula: '=A2*10',
        cachedValue: 999,
        literalRecalculatedValue: 20,
        staleCachedValue: true,
      })
      expect(JSON.parse(stdout)).not.toHaveProperty('nextStep')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('runs xlsx-cache-doctor as the default inspection command', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-cli-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(join(tempDir, 'stale-cache.recalculated.xlsx'))).toBe(false)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.inspectionCompleted).toBe(true)
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.suggestedReads).toEqual(['Sheet1!B2'])
      expect(JSON.parse(stdout)).not.toHaveProperty('nextStep')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('keeps xlsx-cache-doctor in recalculation mode when readback output is explicit', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-recalc-cli-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      const outputPath = join(tempDir, 'stale-cache.fixed.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--read', 'Sheet1!B2', '--out', outputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.recalculationCompleted).toBe(true)
      expect(summary.reads['Sheet1!B2']?.value).toBe(20)
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
  readonly commandSucceeded: boolean
  readonly recalculationCompleted: boolean
  readonly expectedReadback?: Readonly<Record<string, number>>
  readonly expectedValueMatched?: boolean
  readonly excelParity: 'not_proven'
}

interface CliInspectionSummary {
  readonly mode: string
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly staleCachedFormulaCount: number
  readonly suggestedReads: readonly string[]
  readonly formulas: ReadonlyArray<{
    readonly target: string
    readonly formula: string
    readonly cachedValue?: unknown
    readonly literalRecalculatedValue?: unknown
    readonly staleCachedValue: boolean | null
  }>
  readonly commandSucceeded: boolean
  readonly inspectionCompleted: boolean
  readonly recalculationCompleted: boolean
  readonly excelParity: 'not_proven'
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
  const commandSucceeded = Reflect.get(parsed, 'commandSucceeded')
  const recalculationCompleted = Reflect.get(parsed, 'recalculationCompleted')
  const expectedReadback = Reflect.get(parsed, 'expectedReadback')
  const expectedValueMatched = Reflect.get(parsed, 'expectedValueMatched')
  const excelParity = Reflect.get(parsed, 'excelParity')
  if (
    typeof mode !== 'string' ||
    typeof externalWorkbooks !== 'number' ||
    typeof reads !== 'object' ||
    reads === null ||
    !Array.isArray(warnings) ||
    typeof commandSucceeded !== 'boolean' ||
    typeof recalculationCompleted !== 'boolean' ||
    typeof excelParity !== 'string'
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
    commandSucceeded,
    recalculationCompleted,
    ...(readNumericRecord(expectedReadback) ? { expectedReadback: readNumericRecord(expectedReadback) } : {}),
    ...(typeof expectedValueMatched === 'boolean' ? { expectedValueMatched } : {}),
    excelParity: excelParity === 'not_proven' ? excelParity : 'not_proven',
  }
}

function readCliInspectionSummary(stdout: string): CliInspectionSummary {
  const parsed: unknown = JSON.parse(stdout)
  if (!isRecord(parsed)) {
    throw new Error(`Expected CLI inspection summary object, received ${stdout}`)
  }
  const mode = parsed['mode']
  const formulaCellCount = parsed['formulaCellCount']
  const inspectedFormulaCellCount = parsed['inspectedFormulaCellCount']
  const staleCachedFormulaCount = parsed['staleCachedFormulaCount']
  const suggestedReads = parsed['suggestedReads']
  const formulas = parsed['formulas']
  const commandSucceeded = parsed['commandSucceeded']
  const inspectionCompleted = parsed['inspectionCompleted']
  const recalculationCompleted = parsed['recalculationCompleted']
  const excelParity = parsed['excelParity']
  if (
    typeof mode !== 'string' ||
    typeof formulaCellCount !== 'number' ||
    typeof inspectedFormulaCellCount !== 'number' ||
    typeof staleCachedFormulaCount !== 'number' ||
    !Array.isArray(suggestedReads) ||
    !Array.isArray(formulas) ||
    typeof commandSucceeded !== 'boolean' ||
    typeof inspectionCompleted !== 'boolean' ||
    typeof recalculationCompleted !== 'boolean' ||
    excelParity !== 'not_proven'
  ) {
    throw new Error(`Unexpected CLI inspection summary shape: ${stdout}`)
  }
  return {
    mode,
    formulaCellCount,
    inspectedFormulaCellCount,
    staleCachedFormulaCount,
    suggestedReads: suggestedReads.filter((read): read is string => typeof read === 'string'),
    formulas: formulas.filter(isCliInspectionFormula),
    commandSucceeded,
    inspectionCompleted,
    recalculationCompleted,
    excelParity,
  }
}

function isCliInspectionFormula(value: unknown): value is CliInspectionSummary['formulas'][number] {
  return (
    isRecord(value) &&
    typeof value['target'] === 'string' &&
    typeof value['formula'] === 'string' &&
    (typeof value['staleCachedValue'] === 'boolean' || value['staleCachedValue'] === null)
  )
}

function readCliSummaryDiagnostics(value: unknown): CliSummary['diagnostics'] | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const externalWorkbookHydration = value['externalWorkbookHydration']
  return isRecord(externalWorkbookHydration) ? { externalWorkbookHydration } : undefined
}

function readNumericRecord(value: unknown): Readonly<Record<string, number>> | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const parsed: Record<string, number> = {}
  for (const [entryKey, entryValue] of Object.entries(value)) {
    if (typeof entryValue !== 'number') {
      return undefined
    }
    parsed[entryKey] = entryValue
  }
  return parsed
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

function buildStaleFormulaCacheWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Sheet1: [
      ['Input', 'Output'],
      [2, '=A2*10'],
    ],
  })
  try {
    return replaceWorksheetCellXml(
      exportXlsx(workbook.exportSnapshot()),
      'xl/worksheets/sheet1.xml',
      'B2',
      '<c r="B2"><f>A2*10</f><v>999</v></c>',
    )
  } finally {
    workbook.dispose()
  }
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

function replaceWorksheetCellXml(bytes: Uint8Array, path: string, address: string, replacement: string): Uint8Array {
  const zip = unzipSync(bytes)
  const xml = strFromU8(zip[path] ?? new Uint8Array())
  zip[path] = strToU8(xml.replace(new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>[\\s\\S]*?<\\/c>`, 'u'), replacement))
  return zipSync(zip)
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
