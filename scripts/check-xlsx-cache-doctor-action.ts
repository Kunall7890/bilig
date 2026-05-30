import { chmodSync, existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const rootActionPath = join(repoRoot, 'action.yml')
const nestedActionPath = join(repoRoot, 'actions', 'xlsx-cache-doctor', 'action.yml')
const resolveScriptPath = join(repoRoot, 'actions', 'xlsx-cache-doctor', 'resolve-workbooks.mjs')
const inspectScriptPath = join(repoRoot, 'actions', 'xlsx-cache-doctor', 'inspect-workbooks.mjs')

const [rootAction, nestedAction] = await Promise.all([readFile(rootActionPath, 'utf8'), readFile(nestedActionPath, 'utf8')])

if (rootAction !== nestedAction) {
  throw new Error(
    [
      'Root action.yml must stay byte-for-byte aligned with actions/xlsx-cache-doctor/action.yml.',
      'Edit both files together so Marketplace and subdirectory action users get the same inputs, outputs, and behavior.',
    ].join('\n'),
  )
}

if (!existsSync(resolveScriptPath) || !existsSync(inspectScriptPath)) {
  throw new Error('XLSX Cache Doctor action helper scripts must exist next to the nested action.yml.')
}

if (!rootAction.includes('workbooks:') || !rootAction.includes('changed-files-only:')) {
  throw new Error('XLSX Cache Doctor action.yml must expose repo-scale workbooks and changed-files-only inputs.')
}

const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-cache-doctor-action-'))
try {
  writeFileSync(join(tempDir, 'a.xlsx'), '')
  writeFileSync(join(tempDir, 'notes.txt'), '')
  const outputPath = join(tempDir, 'github-output.txt')
  const result = spawnSync(process.execPath, [resolveScriptPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      GITHUB_WORKSPACE: tempDir,
      GITHUB_OUTPUT: outputPath,
      BILIG_WORKBOOKS: '**/*.xlsx',
      BILIG_CHANGED_FILES_ONLY: 'false',
    },
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`XLSX Cache Doctor resolver smoke failed:\n${result.stdout}\n${result.stderr}`)
  }
  const output = await readFile(outputPath, 'utf8')
  if (!output.includes('workbook-count=1') || !output.includes('workbooks-json=["a.xlsx"]')) {
    throw new Error(`Unexpected resolver output:\n${output}`)
  }
} finally {
  rmSync(tempDir, { recursive: true, force: true })
}

const changedFilesTempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-cache-doctor-changed-files-'))
try {
  runGit(changedFilesTempDir, ['init'])
  runGit(changedFilesTempDir, ['config', 'user.email', 'cache-doctor@example.com'])
  runGit(changedFilesTempDir, ['config', 'user.name', 'XLSX Cache Doctor'])
  runGit(changedFilesTempDir, ['config', 'commit.gpgsign', 'false'])
  writeFileSync(join(changedFilesTempDir, 'unchanged.xlsx'), '')
  runGit(changedFilesTempDir, ['add', 'unchanged.xlsx'])
  runGit(changedFilesTempDir, ['commit', '-m', 'base'])
  writeFileSync(join(changedFilesTempDir, 'changed.xlsx'), '')
  runGit(changedFilesTempDir, ['add', 'changed.xlsx'])
  runGit(changedFilesTempDir, ['commit', '-m', 'change workbook'])

  const outputPath = join(changedFilesTempDir, 'github-output.txt')
  const result = spawnSync(process.execPath, [resolveScriptPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      GITHUB_WORKSPACE: changedFilesTempDir,
      GITHUB_OUTPUT: outputPath,
      BILIG_WORKBOOKS: '**/*.xlsx',
      BILIG_CHANGED_FILES_ONLY: 'true',
    },
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`XLSX Cache Doctor changed-file resolver smoke failed:\n${result.stdout}\n${result.stderr}`)
  }
  const output = await readFile(outputPath, 'utf8')
  if (!output.includes('workbook-count=1') || !output.includes('workbooks-json=["changed.xlsx"]')) {
    throw new Error(`Unexpected changed-file resolver output:\n${output}`)
  }
} finally {
  rmSync(changedFilesTempDir, { recursive: true, force: true })
}

const inspectTempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-cache-doctor-inspect-'))
try {
  const fakeBinDir = join(inspectTempDir, 'bin')
  mkdirSync(fakeBinDir)
  const fakeNpmPath = join(fakeBinDir, 'npm')
  writeFileSync(
    fakeNpmPath,
    [
      '#!/usr/bin/env node',
      "const commandIndex = process.argv.indexOf('xlsx-cache-doctor')",
      "const workbook = commandIndex >= 0 ? process.argv[commandIndex + 1] : 'unknown.xlsx'",
      'const report = {',
      "  mode: 'file',",
      '  input: workbook,',
      '  formulaCellCount: 3,',
      '  inspectedFormulaCellCount: 3,',
      '  uninspectedFormulaCellCount: 0,',
      '  staleCachedFormulaCount: 2,',
      '  cacheStatusSummary: { inspected: 3, stale: 2, fresh: 0, missingCache: 0, unsupportedRecalculation: 1 },',
      "  suggestedReads: ['Sheet1!B2', 'Sheet1!B3'],",
      '  formulas: [',
      "    { target: 'Sheet1!B2', formula: '=A2*10', cachedValue: 10, cacheStatus: 'unsupported-recalculation', staleCachedValue: null },",
      "    { target: 'Sheet1!B3', formula: '=A3*10', cachedValue: 999, literalRecalculatedValue: 30, cacheStatus: 'stale', staleCachedValue: true },",
      "    { target: 'Sheet1!B4', formula: '=A4&`|`', cachedValue: 'old\\n%value', literalRecalculatedValue: 'new|value', cacheStatus: 'stale', staleCachedValue: true },",
      '  ],',
      '  warnings: [],',
      '};',
      'console.log(JSON.stringify(report));',
      '',
    ].join('\n'),
  )
  chmodSync(fakeNpmPath, 0o755)

  const outputPath = join(inspectTempDir, 'github-output.txt')
  const summaryPath = join(inspectTempDir, 'summary.md')
  const reportPath = join(inspectTempDir, 'report.json')
  const markdownPath = join(inspectTempDir, 'report.md')
  const result = spawnSync(process.execPath, [inspectScriptPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      PATH: `${fakeBinDir}:${process.env.PATH || ''}`,
      GITHUB_OUTPUT: outputPath,
      GITHUB_STEP_SUMMARY: summaryPath,
      BILIG_WORKBOOKS_JSON: '["fixtures/stale-pricing.xlsx"]',
      BILIG_PACKAGE_VERSION: 'test',
      BILIG_INSPECT_LIMIT: 'all',
      BILIG_JSON_OUTPUT: reportPath,
      BILIG_MARKDOWN_OUTPUT: markdownPath,
      BILIG_FAIL_ON_STALE: 'false',
    },
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`XLSX Cache Doctor inspector smoke failed:\n${result.stdout}\n${result.stderr}`)
  }
  if (
    !result.stdout.includes(
      '::warning title=Stale cached XLSX formula::fixtures/stale-pricing.xlsx#Sheet1!B3 cached 999 but recalculated 30',
    )
  ) {
    throw new Error(`Expected stale-formula warning annotation in stdout:\n${result.stdout}`)
  }
  if (
    !result.stdout.includes(
      '::warning title=Stale cached XLSX formula::fixtures/stale-pricing.xlsx#Sheet1!B4 cached old%0A%25value but recalculated new|value',
    )
  ) {
    throw new Error(`Expected escaped stale-formula warning annotation in stdout:\n${result.stdout}`)
  }
  const output = await readFile(outputPath, 'utf8')
  if (
    !output.includes('stale-count=2') ||
    !output.includes('fresh-count=0') ||
    !output.includes('missing-cache-count=0') ||
    !output.includes('unsupported-recalculation-count=1') ||
    !output.includes(`markdown=${markdownPath}`) ||
    !output.includes('suggested-reads=fixtures/stale-pricing.xlsx#Sheet1!B2,fixtures/stale-pricing.xlsx#Sheet1!B3')
  ) {
    throw new Error(`Unexpected inspector outputs:\n${output}`)
  }
  const summary = await readFile(summaryPath, 'utf8')
  const markdown = await readFile(markdownPath, 'utf8')
  for (const expected of [
    '#### Stale cached formula values',
    '- Unsupported recalculation results: 1',
    '| fixtures/stale-pricing.xlsx | 3 | 2 | 0 | 0 | 1 | Sheet1!B2, Sheet1!B3 |',
    '| fixtures/stale-pricing.xlsx | Sheet1!B3 | `=A3*10` | 999 | 30 |',
    '| fixtures/stale-pricing.xlsx | Sheet1!B4 | `=A4&\\`\\|\\`` | old %value | new\\|value |',
    '#### Follow-up check command',
    "xlsx-recalc 'fixtures/stale-pricing.xlsx' --read 'Sheet1!B3'",
  ]) {
    if (!summary.includes(expected) || !markdown.includes(expected)) {
      throw new Error(`Missing "${expected}" in inspector Markdown output:\nsummary:\n${summary}\nreport:\n${markdown}`)
    }
  }
  const aggregate = parseAggregateReport(JSON.parse(await readFile(reportPath, 'utf8')))
  if (
    aggregate.staleCachedFormulaCount !== 2 ||
    aggregate.cacheStatusSummary?.unsupportedRecalculation !== 1 ||
    aggregate.workbooks?.[0]?.staleFormulas?.length !== 2
  ) {
    throw new Error(`Unexpected inspector JSON report:\n${JSON.stringify(aggregate, null, 2)}`)
  }
} finally {
  rmSync(inspectTempDir, { recursive: true, force: true })
}

function runGit(cwd: string, args: readonly string[]): void {
  const result = spawnSync('git', args, {
    cwd,
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`git ${args.join(' ')} failed:\n${result.stdout}\n${result.stderr}`)
  }
}

function parseAggregateReport(value: unknown): {
  staleCachedFormulaCount?: number
  cacheStatusSummary?: { unsupportedRecalculation?: number }
  workbooks?: Array<{ staleFormulas?: unknown[] }>
} {
  if (!isRecord(value)) {
    throw new Error(`Expected aggregate JSON report object, got ${JSON.stringify(value)}`)
  }
  const workbooksValue = Reflect.get(value, 'workbooks')
  const workbooks = Array.isArray(workbooksValue)
    ? workbooksValue.map((workbook) => {
        if (!isRecord(workbook)) {
          return {}
        }
        const staleFormulas = Reflect.get(workbook, 'staleFormulas')
        return {
          staleFormulas: Array.isArray(staleFormulas) ? staleFormulas : undefined,
        }
      })
    : undefined
  const staleCachedFormulaCount = Reflect.get(value, 'staleCachedFormulaCount')
  const cacheStatusSummaryValue = Reflect.get(value, 'cacheStatusSummary')
  const unsupportedRecalculation = isRecord(cacheStatusSummaryValue)
    ? Reflect.get(cacheStatusSummaryValue, 'unsupportedRecalculation')
    : undefined
  const cacheStatusSummary = isRecord(cacheStatusSummaryValue)
    ? {
        unsupportedRecalculation: typeof unsupportedRecalculation === 'number' ? unsupportedRecalculation : undefined,
      }
    : undefined
  return {
    staleCachedFormulaCount: typeof staleCachedFormulaCount === 'number' ? staleCachedFormulaCount : undefined,
    cacheStatusSummary,
    workbooks,
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
