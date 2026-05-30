import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { parse as parseYaml } from 'yaml'

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

  it('keeps human CLI output focused on recalculation results', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-human-'))
    try {
      const outputPath = join(tempDir, 'demo.recalculated.xlsx')
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli(['--demo', '--out', outputPath, '--read', 'Summary!B2'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      expect(stdout).toContain('Recalculated generated demo workbook ->')
      expect(stdout).toContain('Summary!B2:')
      expect(stdout).not.toContain('star or bookmark')
      expect(stdout).not.toContain('adoption blocker')
      expect(stdout).not.toContain('Watch formula')
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
      expect(summary.schemaVersion).toBe('xlsx-cache-doctor.v1')
      expect(summary.mode).toBe('file')
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.inspectionCompleted).toBe(true)
      expect(summary.recalculationCompleted).toBe(true)
      expect(summary.excelParity).toBe('not_proven')
      expect(summary.formulaCellCount).toBe(1)
      expect(summary.inspectedFormulaCellCount).toBe(1)
      expect(summary.uninspectedFormulaCellCount).toBe(0)
      expect(summary.inspectionLimit).toBe('all')
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.cacheStatusSummary).toEqual({
        inspected: 1,
        stale: 1,
        fresh: 0,
        missingCache: 0,
        unsupportedRecalculation: 0,
      })
      expect(summary.suggestedReads).toEqual(['Sheet1!B2'])
      expect(summary.formulas[0]).toMatchObject({
        target: 'Sheet1!B2',
        formula: '=A2*10',
        cachedValue: 999,
        literalRecalculatedValue: 20,
        cacheStatus: 'stale',
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
      expect(summary.schemaVersion).toBe('xlsx-cache-doctor.v1')
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.inspectionCompleted).toBe(true)
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.cacheStatusSummary.stale).toBe(1)
      expect(summary.uninspectedFormulaCellCount).toBe(0)
      expect(summary.suggestedReads).toEqual(['Sheet1!B2'])
      expect(JSON.parse(stdout)).not.toHaveProperty('nextStep')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('runs a cache-doctor demo that proves a stale cached formula value', () => {
    let stdout = ''

    const exitCode = runXlsxFormulaRecalcCli(['--demo', '--json'], {
      commandName: 'xlsx-cache-doctor',
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const summary = readCliInspectionSummary(stdout)
    expect(summary.schemaVersion).toBe('xlsx-cache-doctor.v1')
    expect(summary.mode).toBe('demo')
    expect(summary.commandSucceeded).toBe(true)
    expect(summary.inspectionCompleted).toBe(true)
    expect(summary.formulaCellCount).toBe(1)
    expect(summary.inspectedFormulaCellCount).toBe(1)
    expect(summary.uninspectedFormulaCellCount).toBe(0)
    expect(summary.staleCachedFormulaCount).toBe(1)
    expect(summary.cacheStatusSummary).toEqual({
      inspected: 1,
      stale: 1,
      fresh: 0,
      missingCache: 0,
      unsupportedRecalculation: 0,
    })
    expect(summary.suggestedReads).toEqual(['Summary!B2'])
    expect(summary.formulas[0]).toMatchObject({
      target: 'Summary!B2',
      formula: '=Inputs!B2*Inputs!B3',
      cachedValue: 60_000,
      literalRecalculatedValue: 72_000,
      cacheStatus: 'stale',
      staleCachedValue: true,
    })
  })

  it('separates missing cached formula values from stale cached values', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-missing-cache-'))
    try {
      const inputPath = join(tempDir, 'missing-cache.xlsx')
      writeFileSync(inputPath, buildMissingFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.staleCachedFormulaCount).toBe(0)
      expect(summary.cacheStatusSummary).toEqual({
        inspected: 1,
        stale: 0,
        fresh: 0,
        missingCache: 1,
        unsupportedRecalculation: 0,
      })
      expect(summary.formulas[0]).toMatchObject({
        target: 'Sheet1!B2',
        formula: '=A2*10',
        literalRecalculatedValue: 20,
        cacheStatus: 'missing-cache',
        staleCachedValue: null,
      })
      expect(summary.formulas[0]).not.toHaveProperty('cachedValue')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('prints a ready-to-commit GitHub Actions workflow for cache doctor adoption', () => {
    let stdout = ''

    const exitCode = runXlsxFormulaRecalcCli(
      [
        '--print-github-action',
        'fixtures/pricing model.xlsx',
        '--fail-on-stale',
        'false',
        '--inspect-limit',
        '50',
        '--json-output',
        '${{ runner.temp }}/custom-cache-doctor.json',
        '--markdown-output',
        '${{ runner.temp }}/custom-cache-doctor.md',
        '--package-version',
        '0.124.1',
        '--workflow-name',
        'workbook cache doctor',
      ],
      {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      },
    )

    expect(exitCode).toBe(0)
    expect(stdout).toContain('name: "workbook cache doctor"')
    expect(stdout).toContain('pull_request:')
    expect(stdout).toContain('- "**/*.xlsx"')
    expect(stdout).toContain('fetch-depth: 0')
    expect(stdout).toContain('uses: actions/setup-node@v6')
    expect(stdout).toContain('node-version: "22"')
    expect(stdout).toContain('package-manager-cache: false')
    expect(stdout).toContain('uses: proompteng/bilig@v1')
    expect(stdout).toContain('workbooks: "fixtures/pricing model.xlsx"')
    expect(stdout).toContain('changed-files-only: "true"')
    expect(stdout).toContain('inspect-limit: "50"')
    expect(stdout).toContain('json-output: "${{ runner.temp }}/custom-cache-doctor.json"')
    expect(stdout).toContain('markdown-output: "${{ runner.temp }}/custom-cache-doctor.md"')
    expect(stdout).toContain('package-version: "0.124.1"')
    expect(stdout).toContain('fail-on-stale: "false"')
    expect(stdout).toContain('name: xlsx-cache-doctor-report')

    const workflow = readGeneratedWorkflow(stdout)
    const jobs = requireRecord(workflow['jobs'])
    const job = requireRecord(jobs['inspect-xlsx-formula-caches'])
    const steps = requireRecordArray(job['steps'])
    const checkout = steps[0] ?? {}
    const setupNode = steps[1] ?? {}
    const cacheDoctor = steps[2] ?? {}
    const cacheDoctorInputs = requireRecord(cacheDoctor['with'])
    const uploadArtifact = steps[3] ?? {}

    expect(workflow['name']).toBe('workbook cache doctor')
    expect(workflow['on']).toEqual({
      pull_request: { paths: ['**/*.xlsx'] },
      workflow_dispatch: null,
    })
    expect(workflow['permissions']).toEqual({ contents: 'read' })
    expect(checkout).toMatchObject({ uses: 'actions/checkout@v5', with: { 'fetch-depth': 0 } })
    expect(setupNode).toMatchObject({
      uses: 'actions/setup-node@v6',
      with: { 'node-version': '22', 'package-manager-cache': false },
    })
    expect(cacheDoctor).toMatchObject({ id: 'cache-doctor', uses: 'proompteng/bilig@v1' })
    expect(cacheDoctorInputs).toEqual({
      workbooks: 'fixtures/pricing model.xlsx',
      'changed-files-only': 'true',
      'package-version': '0.124.1',
      'inspect-limit': '50',
      'json-output': '${{ runner.temp }}/custom-cache-doctor.json',
      'markdown-output': '${{ runner.temp }}/custom-cache-doctor.md',
      'fail-on-stale': 'false',
    })
    expect(uploadArtifact).toMatchObject({
      uses: 'actions/upload-artifact@v4',
      if: 'always()',
      with: {
        name: 'xlsx-cache-doctor-report',
        path: '${{ steps.cache-doctor.outputs.json }}\n${{ steps.cache-doctor.outputs.markdown }}\n',
      },
    })
  })

  it('requires a workbook path when printing a GitHub Actions workflow', () => {
    let stderr = ''

    const exitCode = runXlsxFormulaRecalcCli(['--print-github-action'], {
      commandName: 'xlsx-cache-doctor',
      stderr: (text) => {
        stderr += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stderr).toContain('Expected workbook path or glob after --print-github-action')
  })

  it('allows generated GitHub Actions workflows to scan all matching workbooks', () => {
    let stdout = ''

    const exitCode = runXlsxFormulaRecalcCli(['--print-github-action', '**/*.xlsx', '--changed-files-only', 'false'], {
      commandName: 'xlsx-cache-doctor',
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('workbooks: "**/*.xlsx"')
    expect(stdout).toContain('changed-files-only: "false"')
    expect(stdout).toContain('inspect-limit: "all"')
    expect(stdout).toContain('package-version: "0.130.0"')
    expect(stdout).toContain('json-output: "${{ runner.temp }}/xlsx-cache-doctor.json"')
    expect(stdout).toContain('markdown-output: "${{ runner.temp }}/xlsx-cache-doctor.md"')
    expect(stdout).toContain('fail-on-stale: "false"')
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

  it('checks every formula by default so stale caches after the old sample cutoff fail inspection', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-all-formulas-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.formulaCellCount).toBe(60)
      expect(summary.inspectedFormulaCellCount).toBe(60)
      expect(summary.uninspectedFormulaCellCount).toBe(0)
      expect(summary.inspectionLimit).toBe('all')
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.cacheStatusSummary.stale).toBe(1)
      expect(summary.suggestedReads).toContain('Sheet1!B61')
      expect(summary.formulas.find((formula) => formula.target === 'Sheet1!B61')).toMatchObject({
        formula: '=A61*10',
        cachedValue: 999,
        literalRecalculatedValue: 600,
        cacheStatus: 'stale',
        staleCachedValue: true,
      })
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('reports uninspected formulas when a caller sets an explicit inspection limit', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-limited-formulas-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--inspect-limit', '50', '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.formulaCellCount).toBe(60)
      expect(summary.inspectedFormulaCellCount).toBe(50)
      expect(summary.uninspectedFormulaCellCount).toBe(10)
      expect(summary.inspectionLimit).toBe(50)
      expect(summary.staleCachedFormulaCount).toBe(0)
      expect(summary.cacheStatusSummary).toMatchObject({ inspected: 50, stale: 0 })
      expect(summary.suggestedReads).not.toContain('Sheet1!B61')
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
  readonly schemaVersion: 'xlsx-cache-doctor.v1'
  readonly formulaCellCount: number
  readonly inspectedFormulaCellCount: number
  readonly uninspectedFormulaCellCount: number
  readonly inspectionLimit: number | 'all'
  readonly staleCachedFormulaCount: number
  readonly cacheStatusSummary: {
    readonly inspected: number
    readonly stale: number
    readonly fresh: number
    readonly missingCache: number
    readonly unsupportedRecalculation: number
  }
  readonly suggestedReads: readonly string[]
  readonly formulas: ReadonlyArray<{
    readonly target: string
    readonly formula: string
    readonly cachedValue?: unknown
    readonly literalRecalculatedValue?: unknown
    readonly cacheStatus: 'fresh' | 'stale' | 'missing-cache' | 'unsupported-recalculation'
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
  const schemaVersion = parsed['schemaVersion']
  const formulaCellCount = parsed['formulaCellCount']
  const inspectedFormulaCellCount = parsed['inspectedFormulaCellCount']
  const uninspectedFormulaCellCount = parsed['uninspectedFormulaCellCount']
  const inspectionLimit = parsed['inspectionLimit']
  const staleCachedFormulaCount = parsed['staleCachedFormulaCount']
  const cacheStatusSummary = parsed['cacheStatusSummary']
  const suggestedReads = parsed['suggestedReads']
  const formulas = parsed['formulas']
  const commandSucceeded = parsed['commandSucceeded']
  const inspectionCompleted = parsed['inspectionCompleted']
  const recalculationCompleted = parsed['recalculationCompleted']
  const excelParity = parsed['excelParity']
  if (
    typeof mode !== 'string' ||
    schemaVersion !== 'xlsx-cache-doctor.v1' ||
    typeof formulaCellCount !== 'number' ||
    typeof inspectedFormulaCellCount !== 'number' ||
    typeof uninspectedFormulaCellCount !== 'number' ||
    !isInspectionLimit(inspectionLimit) ||
    typeof staleCachedFormulaCount !== 'number' ||
    !isCliCacheStatusSummary(cacheStatusSummary) ||
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
    schemaVersion,
    formulaCellCount,
    inspectedFormulaCellCount,
    uninspectedFormulaCellCount,
    inspectionLimit,
    staleCachedFormulaCount,
    cacheStatusSummary,
    suggestedReads: suggestedReads.filter((read): read is string => typeof read === 'string'),
    formulas: formulas.filter(isCliInspectionFormula),
    commandSucceeded,
    inspectionCompleted,
    recalculationCompleted,
    excelParity,
  }
}

function readGeneratedWorkflow(stdout: string): Record<string, unknown> {
  return requireRecord(parseYaml(stdout))
}

function requireRecord(value: unknown): Record<string, unknown> {
  if (!isRecord(value) || Array.isArray(value)) {
    throw new Error(`Expected object, received ${typeof value}`)
  }
  return value
}

function requireRecordArray(value: unknown): readonly Record<string, unknown>[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected array, received ${typeof value}`)
  }
  return value.map(requireRecord)
}

function isInspectionLimit(value: unknown): value is CliInspectionSummary['inspectionLimit'] {
  return value === 'all' || typeof value === 'number'
}

function isCliInspectionFormula(value: unknown): value is CliInspectionSummary['formulas'][number] {
  return (
    isRecord(value) &&
    typeof value['target'] === 'string' &&
    typeof value['formula'] === 'string' &&
    isCliCacheStatus(value['cacheStatus']) &&
    (typeof value['staleCachedValue'] === 'boolean' || value['staleCachedValue'] === null)
  )
}

function isCliCacheStatus(value: unknown): value is CliInspectionSummary['formulas'][number]['cacheStatus'] {
  return value === 'fresh' || value === 'stale' || value === 'missing-cache' || value === 'unsupported-recalculation'
}

function isCliCacheStatusSummary(value: unknown): value is CliInspectionSummary['cacheStatusSummary'] {
  return (
    isRecord(value) &&
    typeof value['inspected'] === 'number' &&
    typeof value['stale'] === 'number' &&
    typeof value['fresh'] === 'number' &&
    typeof value['missingCache'] === 'number' &&
    typeof value['unsupportedRecalculation'] === 'number'
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

function buildMissingFormulaCacheWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Sheet1: [
      ['Input', 'Output'],
      [2, '=A2*10'],
    ],
  })
  try {
    return replaceWorksheetCellXml(exportXlsx(workbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f></c>')
  } finally {
    workbook.dispose()
  }
}

function buildManyFormulaCacheWorkbook(): Uint8Array {
  const rows: Array<[number | string, number | string]> = [['Input', 'Output']]
  for (let row = 2; row <= 61; row += 1) {
    rows.push([row - 1, `=A${row}*10`])
  }
  const workbook = WorkPaper.buildFromSheets({
    Sheet1: rows,
  })
  try {
    return replaceWorksheetCellXml(
      exportXlsx(workbook.exportSnapshot()),
      'xl/worksheets/sheet1.xml',
      'B61',
      '<c r="B61"><f>A61*10</f><v>999</v></c>',
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
