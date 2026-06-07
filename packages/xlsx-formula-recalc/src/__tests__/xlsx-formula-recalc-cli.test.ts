import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { parse as parseYaml } from 'yaml'

import { runXlsxFormulaRecalcCli, runXlsxFormulaRecalcCliAsync } from '../cli-api.js'
import { WorkPaper, exportXlsx } from 'bilig-workpaper/xlsx'
import { buildWorkbookCompatibilityReport, runWorkbookCompatibilityReportCli } from '../workbook-compatibility-report.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const packageVersion = readPackageVersion()

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

  it('uses streaming-native for async file-to-file recalculation by default', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-native-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--out', outputPath, '--read', 'Sheet1!B2', '--json'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.reads['Sheet1!B2']?.value).toBe(20)
      expect(summary.diagnostics.engineMode).toBe('streaming-native')
      expect(readCachedFormulaValue(readFileSync(outputPath), 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('does not load SheetJS xlsx for streaming-native file-to-file recalculation', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-native-no-sheetjs-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      const cliApiUrl = pathToFileURL(join(process.cwd(), 'packages/xlsx-formula-recalc/src/cli-api.ts')).href
      const script = `
void (async () => {
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
const loadedXlsxModules = () =>
  Object.keys(requireForCache.cache).filter((path) => /(?:^|[\\\\/])xlsx(?:[\\\\/]|$)|[\\\\/]\\.pnpm[\\\\/]xlsx@/u.test(path))
const before = loadedXlsxModules()
const { runXlsxFormulaRecalcCliAsync } = await import(${JSON.stringify(cliApiUrl)})
let stdout = ''
let stderr = ''
const exitCode = await runXlsxFormulaRecalcCliAsync(
  [
    ${JSON.stringify(inputPath)},
    '--out',
    ${JSON.stringify(outputPath)},
    '--read',
    'Sheet1!B2',
    '--engine',
    'streaming-native',
    '--fallback-policy',
    'error',
    '--json',
  ],
  {
    stdout: (text) => {
      stdout += text
    },
    stderr: (text) => {
      stderr += text
    },
  },
)
process.stdout.write(JSON.stringify({ exitCode, stderr, before, after: loadedXlsxModules(), summary: JSON.parse(stdout) }) + '\\n')
})().catch((error) => {
  console.error(error)
  process.exit(1)
})
`
      const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
        cwd: process.cwd(),
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'pipe'],
      })

      expect(result.status, result.stderr).toBe(0)
      const output = readNoSheetJsChildOutput(result.stdout)
      expect(output.exitCode, output.stderr).toBe(0)
      expect(output.before).toEqual([])
      expect(output.after).toEqual([])
      expect(output.summary.diagnostics?.engineMode).toBe('streaming-native')
      expect(output.summary.reads['Sheet1!B2']?.value).toBe(20)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses synchronous file-to-file recalculation so callers use the file-backed native path', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-sync-file-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stderr = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--out', outputPath, '--read', 'Sheet1!B2', '--json'], {
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(existsSync(outputPath)).toBe(false)
      expect(stderr).toContain('runXlsxFormulaRecalcCliAsync')
      expect(stderr).toContain('file-backed streaming-native engine')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses synchronous file inspection so callers use the file-backed native inspector', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-sync-inspect-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stderr = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--inspect', '--json'], {
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toContain('runXlsxFormulaRecalcCliAsync')
      expect(stderr).toContain('file-backed streaming-native inspector')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses WorkPaper engine selection on the primary async file CLI', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-workpaper-engine-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''
      let stderr = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--out', outputPath, '--engine', 'workpaper', '--json'], {
        stdout: (text) => {
          stdout += text
        },
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toBe('')
      expect(existsSync(outputPath)).toBe(false)
      const summary = readCliErrorSummary(stdout)
      expect(summary.commandSucceeded).toBe(false)
      expect(summary.recalculationCompleted).toBe(false)
      expect(summary.error).toContain('no longer loads or exports WorkPaper')
      expect(summary.error).toContain('@bilig/workpaper')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses WorkPaper fallback policy on the primary async file CLI', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-workpaper-fallback-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''
      let stderr = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--out', outputPath, '--fallback-policy', 'workpaper', '--json'], {
        stdout: (text) => {
          stdout += text
        },
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toBe('')
      expect(existsSync(outputPath)).toBe(false)
      const summary = readCliErrorSummary(stdout)
      expect(summary.commandSucceeded).toBe(false)
      expect(summary.recalculationCompleted).toBe(false)
      expect(summary.error).toContain('no longer loads or exports WorkPaper')
      expect(summary.error).toContain('@bilig/workpaper')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('documents the formula evaluation timeout option', () => {
    let stdout = ''

    const exitCode = runXlsxFormulaRecalcCli(['--help'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('--timeout-ms <n>')
  })

  it('inspects workbook formula cells through streaming-native before writing an output file', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-inspect-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--inspect', '--json'], {
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
      expect(JSON.parse(stdout).diagnostics.engineMode).toBe('streaming-native')
      expect(JSON.parse(stdout)).not.toHaveProperty('nextStep')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('runs xlsx-cache-doctor as the default streaming-native inspection command', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-cli-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--json'], {
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
      expect(JSON.parse(stdout).diagnostics.engineMode).toBe('streaming-native')
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

  it('prints a workbook compatibility report without claiming an Excel compatibility score', () => {
    let stdout = ''

    const exitCode = runWorkbookCompatibilityReportCli(['--demo', '--json'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const report = readWorkbookCompatibilityReport(stdout)
    expect(report.schemaVersion).toBe('bilig-workbook-compatibility-report.v1')
    expect(report.verified).toBe(true)
    expect(report.risk.level).toBe('high')
    expect(report.workbook.formulaCellCount).toBe(3)
    expect(report.findings.unsupportedFunctions).toEqual([{ name: 'CUBEVALUE', count: 1 }])
    expect(report.findings.volatileFunctions).toEqual([{ name: 'NOW', count: 1 }])
    expect(report.findings.staleCachedFormulas.count).toBe(1)
    expect(report.findings.missingCachedFormulaValues.count).toBe(1)
    expect(report.findings.unsupportedRecalculations.count).toBe(1)
    expect(report.recalculationCompleted).toBe(false)
    expect(report.excelParity).toBe('not_proven')
    expect(report.limitations).toContain('It is not an Excel compatibility certification.')
    expect(report.limitations).toContain(
      'It scans workbook package metadata and formula caches; use xlsx-cache-doctor for native recalculation proof.',
    )
    expect(stdout).not.toMatch(/compatibilityScore|excelCompatibilityPercent/u)
  })

  it('keeps workbook compatibility human output explicit about the trust boundary', () => {
    let stdout = ''

    const exitCode = runWorkbookCompatibilityReportCli(['--demo'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('Workbook analyzed. Risk level: HIGH')
    expect(stdout).toContain('Unsupported functions: CUBEVALUE (1)')
    expect(stdout).toContain('It is not an Excel compatibility certification.')
    expect(stdout).not.toContain('compatibility score')
  })

  it('flags provider-backed workbook formulas as high-risk unsupported functions', () => {
    const report = buildWorkbookCompatibilityReport(buildProviderBackedRiskWorkbook(), {
      fileName: 'provider-backed-risk.xlsx',
    })

    expect(report.risk.level).toBe('high')
    expect(report.findings.unsupportedFunctions).toEqual([
      { name: 'GOOGLEFINANCE', count: 1 },
      { name: 'IMPORTDATA', count: 1 },
      { name: 'IMPORTHTML', count: 1 },
      { name: 'IMPORTRANGE', count: 1 },
      { name: 'TRANSLATE', count: 1 },
    ])
    expect(report.risk.reasons.join('\n')).toContain('unsupported functions:')
  })

  it('counts external workbook links at workbook reference grain', () => {
    const report = buildWorkbookCompatibilityReport(buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx'), {
      fileName: 'external-link-risk.xlsx',
    })

    expect(report.findings.externalLinks.count).toBe(1)
    expect(report.findings.externalLinks.unresolvedCount).toBe(0)
    expect(report.risk.level).toBe('high')
    expect(report.risk.reasons).toContain('external workbook links: 1')
  })

  it('raises risk when inspection limits leave formulas unchecked', () => {
    const report = buildWorkbookCompatibilityReport(buildManyFormulaCacheWorkbook(), {
      fileName: 'limited-inspection.xlsx',
      inspectLimit: 50,
    })

    expect(report.cacheInspection).toMatchObject({
      inspectedFormulaCellCount: 50,
      uninspectedFormulaCellCount: 10,
      inspectionLimit: 50,
    })
    expect(report.risk.level).toBe('medium')
    expect(report.risk.reasons).toContain('uninspected formula cells: 10')
  })

  it('separates missing cached formula values from stale cached values', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-missing-cache-'))
    try {
      const inputPath = join(tempDir, 'missing-cache.xlsx')
      writeFileSync(inputPath, buildMissingFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--json'], {
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
    expect(stdout).toContain(`package-version: "${packageVersion}"`)
    expect(stdout).toContain('json-output: "${{ runner.temp }}/xlsx-cache-doctor.json"')
    expect(stdout).toContain('markdown-output: "${{ runner.temp }}/xlsx-cache-doctor.md"')
    expect(stdout).toContain('fail-on-stale: "false"')
  })

  it('keeps xlsx-cache-doctor in recalculation mode when readback output is explicit', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-recalc-cli-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      const outputPath = join(tempDir, 'stale-cache.fixed.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--read', 'Sheet1!B2', '--out', outputPath, '--json'], {
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

  it('checks every formula by default so stale caches after the old sample cutoff fail inspection', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-all-formulas-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--json'], {
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

  it('reports uninspected formulas when a caller sets an explicit inspection limit', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-limited-formulas-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--inspect-limit', '50', '--json'], {
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

  it('hydrates external-link caches from companion workbook paths through streaming-native', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-external-'))
    try {
      const inputPath = join(tempDir, 'model.xlsx')
      const companionPath = join(tempDir, 'uploaded-rates.xlsx')
      const outputPath = join(tempDir, 'model.recalculated.xlsx')
      writeFileSync(inputPath, buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx', { lookupFormulas: false }))
      writeFileSync(companionPath, buildExternalSourceWorkbook([20, 30, 40]))
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync(
        [
          inputPath,
          '--external-workbook-target',
          companionPath,
          'file:///tmp/rates.xlsx',
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
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.externalWorkbooks).toBe(1)
      expect(summary.reads['Model!C1']?.value).toBe(180)
      expect(summary.diagnostics?.externalWorkbookHydration).toMatchObject({
        externalWorkbookCount: 1,
        refreshedBookIndices: [1],
        refreshedCellCount: 3,
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

  it('preserves cached external-link values when CLI companion workbook paths are ambiguous through streaming-native', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-ambiguous-'))
    try {
      const inputPath = join(tempDir, 'model.xlsx')
      const firstCompanionPath = join(tempDir, 'one', 'rates.xlsx')
      const secondCompanionPath = join(tempDir, 'two', 'rates.xlsx')
      const outputPath = join(tempDir, 'model.recalculated.xlsx')
      mkdirSync(join(tempDir, 'one'))
      mkdirSync(join(tempDir, 'two'))
      writeFileSync(inputPath, buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx', { lookupFormulas: false }))
      writeFileSync(firstCompanionPath, buildExternalSourceWorkbook([20, 30, 40]))
      writeFileSync(secondCompanionPath, buildExternalSourceWorkbook([200, 300, 400]))
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync(
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
    readonly engineMode?: string
    readonly externalWorkbookHydration?: Record<string, unknown>
  }
  readonly commandSucceeded: boolean
  readonly recalculationCompleted: boolean
  readonly expectedReadback?: Readonly<Record<string, number>>
  readonly expectedValueMatched?: boolean
  readonly excelParity: 'not_proven'
}

interface NoSheetJsChildOutput {
  readonly exitCode: number
  readonly stderr: string
  readonly before: readonly string[]
  readonly after: readonly string[]
  readonly summary: CliSummary
}

interface CliErrorSummary {
  readonly commandSucceeded: false
  readonly recalculationCompleted: false
  readonly error: string
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

interface WorkbookCompatibilityReportForTest {
  readonly schemaVersion: string
  readonly verified: boolean
  readonly workbook: {
    readonly formulaCellCount: number
  }
  readonly findings: {
    readonly unsupportedFunctions: readonly { readonly name: string; readonly count: number }[]
    readonly volatileFunctions: readonly { readonly name: string; readonly count: number }[]
    readonly staleCachedFormulas: { readonly count: number }
    readonly missingCachedFormulaValues: { readonly count: number }
    readonly unsupportedRecalculations: { readonly count: number }
  }
  readonly risk: {
    readonly level: string
  }
  readonly recalculationCompleted: boolean
  readonly excelParity: 'not_proven'
  readonly limitations: readonly string[]
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

function readNoSheetJsChildOutput(stdout: string): NoSheetJsChildOutput {
  const parsed: unknown = JSON.parse(stdout)
  const record = requireRecord(parsed)
  return {
    exitCode: requireNumber(record['exitCode']),
    stderr: requireString(record['stderr']),
    before: requireStringArray(record['before']),
    after: requireStringArray(record['after']),
    summary: readCliSummary(JSON.stringify(record['summary'])),
  }
}

function readCliErrorSummary(stdout: string): CliErrorSummary {
  const record = requireRecord(JSON.parse(stdout))
  return {
    commandSucceeded: requireFalse(record['commandSucceeded']),
    recalculationCompleted: requireFalse(record['recalculationCompleted']),
    error: requireString(record['error']),
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

function readWorkbookCompatibilityReport(stdout: string): WorkbookCompatibilityReportForTest {
  const parsed: unknown = JSON.parse(stdout)
  if (!isRecord(parsed)) {
    throw new Error(`Expected workbook compatibility report object, received ${stdout}`)
  }
  const workbook = parsed['workbook']
  const findings = parsed['findings']
  const risk = parsed['risk']
  const limitations = parsed['limitations']
  if (!isRecord(workbook) || !isRecord(findings) || !isRecord(risk) || !Array.isArray(limitations)) {
    throw new Error(`Unexpected workbook compatibility report shape: ${stdout}`)
  }
  return {
    schemaVersion: requireString(parsed['schemaVersion']),
    verified: parsed['verified'] === true,
    workbook: {
      formulaCellCount: requireNumber(workbook['formulaCellCount']),
    },
    findings: {
      unsupportedFunctions: requireNamedCounts(findings['unsupportedFunctions']),
      volatileFunctions: requireNamedCounts(findings['volatileFunctions']),
      staleCachedFormulas: requireCountObject(findings['staleCachedFormulas']),
      missingCachedFormulaValues: requireCountObject(findings['missingCachedFormulaValues']),
      unsupportedRecalculations: requireCountObject(findings['unsupportedRecalculations']),
    },
    risk: {
      level: requireString(risk['level']),
    },
    recalculationCompleted: parsed['recalculationCompleted'] === true,
    excelParity: parsed['excelParity'] === 'not_proven' ? 'not_proven' : 'not_proven',
    limitations: limitations.filter((limitation): limitation is string => typeof limitation === 'string'),
  }
}

function readGeneratedWorkflow(stdout: string): Record<string, unknown> {
  return requireRecord(parseYaml(stdout))
}

function readPackageVersion(): string {
  const parsed: unknown = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'))
  if (!isRecord(parsed) || Array.isArray(parsed)) {
    throw new Error('Expected package.json object')
  }
  const version = parsed['version']
  if (typeof version !== 'string' || version.length === 0) {
    throw new Error('Expected package.json version')
  }
  return version
}

function requireRecord(value: unknown): Record<string, unknown> {
  if (!isRecord(value) || Array.isArray(value)) {
    throw new Error(`Expected object, received ${typeof value}`)
  }
  return value
}

function requireString(value: unknown): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected string, received ${typeof value}`)
  }
  return value
}

function requireStringArray(value: unknown): readonly string[] {
  if (!Array.isArray(value) || !value.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected string array, received ${typeof value}`)
  }
  return value
}

function requireFalse(value: unknown): false {
  if (value !== false) {
    throw new Error(`Expected false, received ${typeof value}`)
  }
  return value
}

function requireNumber(value: unknown): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`Expected finite number, received ${typeof value}`)
  }
  return value
}

function requireCountObject(value: unknown): { readonly count: number } {
  const record = requireRecord(value)
  return { count: requireNumber(record['count']) }
}

function requireNamedCounts(value: unknown): readonly { readonly name: string; readonly count: number }[] {
  return requireRecordArray(value).map((entry) => ({
    name: requireString(entry['name']),
    count: requireNumber(entry['count']),
  }))
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
  const engineMode = value['engineMode']
  const parsed = {
    ...(typeof engineMode === 'string' ? { engineMode } : {}),
    ...(isRecord(externalWorkbookHydration) ? { externalWorkbookHydration } : {}),
  }
  return Object.keys(parsed).length > 0 ? parsed : undefined
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

function readCachedFormulaValue(bytes: Uint8Array, sheetPath: string, address: string): string | null {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const match = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)
  return match?.[1] ?? null
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

function buildProviderBackedRiskWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Risks: [
      ['Signal', 'Formula'],
      ['Market data', '=GOOGLEFINANCE("GOOG","price")'],
      ['CSV import', '=IMPORTDATA("https://example.com/data.csv")'],
      ['HTML import', '=IMPORTHTML("https://example.com","table",1)'],
      ['Range import', '=IMPORTRANGE("source","Revenue!B2")'],
      ['Translate', '=TRANSLATE("hello","en","es")'],
    ],
  })
  try {
    return exportXlsx(workbook.exportSnapshot())
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

function buildExternalLinkRangeCacheWorkbook(target: string, options: { readonly lookupFormulas?: boolean } = {}): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [
      [null, 2, 120],
      [null, null, 40],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array()).replace(
      /<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u,
      '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>120</v></c>',
    )
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      options.lookupFormulas === false
        ? sheetXml
        : sheetXml.replace(
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
