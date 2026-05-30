import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
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
const xlsxFormulaRecalcPackage = parsePackageJson(
  await readFile(join(repoRoot, 'packages', 'bilig-xlsx-formula-recalc', 'package.json'), 'utf8'),
)

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

if (!rootAction.includes(`default: '${xlsxFormulaRecalcPackage.version}'`)) {
  throw new Error(
    `XLSX Cache Doctor action.yml must pin package-version default to @bilig/xlsx-formula-recalc ${xlsxFormulaRecalcPackage.version}.`,
  )
}

if (!rootAction.includes('Validate Node.js runtime') || !rootAction.includes('requires Node.js 22+')) {
  throw new Error('XLSX Cache Doctor action.yml must fail clearly when consumers omit Node.js 22+ setup.')
}

if (
  !rootAction.includes('Install XLSX Cache Doctor package') ||
  !rootAction.includes('npm install --prefix "$package_dir"') ||
  !rootAction.includes('BILIG_PACKAGE_ROOT: ${{ steps.package.outputs.package-root }}')
) {
  throw new Error('XLSX Cache Doctor action.yml must install @bilig/xlsx-formula-recalc once and pass its package root to the inspector.')
}

const inspectScript = await readFile(inspectScriptPath, 'utf8')
if (inspectScript.includes('spawnSync')) {
  throw new Error(
    'XLSX Cache Doctor inspector must not spawn npm per workbook; install the package once and call inspectXlsxCache in-process.',
  )
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
  const outputs = parseGithubOutput(output)
  if (outputs['workbook-count'] !== '1' || outputs['workbooks-json'] !== '["a.xlsx"]') {
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
  const outputs = parseGithubOutput(output)
  if (outputs['workbook-count'] !== '1' || outputs['workbooks-json'] !== '["changed.xlsx"]') {
    throw new Error(`Unexpected changed-file resolver output:\n${output}`)
  }
} finally {
  rmSync(changedFilesTempDir, { recursive: true, force: true })
}

const inspectTempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-cache-doctor-inspect-'))
try {
  const fakePackageRoot = join(inspectTempDir, 'package-root')
  const fakeDistDir = join(fakePackageRoot, 'dist')
  const fixtureDir = join(inspectTempDir, 'fixtures')
  mkdirSync(fakeDistDir, { recursive: true })
  mkdirSync(fixtureDir, { recursive: true })
  writeFileSync(join(fakePackageRoot, 'package.json'), '{"type":"module"}\n')
  writeFileSync(
    join(fakeDistDir, 'index.js'),
    [
      "import { appendFileSync } from 'node:fs';",
      "if (process.env.BILIG_FAKE_IMPORT_MARKER) appendFileSync(process.env.BILIG_FAKE_IMPORT_MARKER, 'loaded\\n');",
      'export function inspectXlsxCache(_input, options = {}) {',
      "  const workbook = options.fileName || 'unknown.xlsx';",
      '  return {',
      "  schemaVersion: 'xlsx-cache-doctor.v1',",
      "  sheetNames: ['Sheet1'],",
      '  formulaCellCount: workbook.includes("fresh") ? 1 : 3,',
      '  inspectedFormulaCellCount: workbook.includes("fresh") ? 1 : 3,',
      '  uninspectedFormulaCellCount: 0,',
      '  staleCachedFormulaCount: workbook.includes("fresh") ? 0 : 2,',
      '  cacheStatusSummary: workbook.includes("fresh")',
      '    ? { inspected: 1, stale: 0, fresh: 1, missingCache: 0, unsupportedRecalculation: 0 }',
      '    : { inspected: 3, stale: 2, fresh: 0, missingCache: 0, unsupportedRecalculation: 1 },',
      "  suggestedReads: workbook.includes('fresh') ? ['Sheet1!B9'] : ['Sheet1!B2', 'Sheet1!B3'],",
      '  formulas: workbook.includes("fresh") ? [',
      "    { target: 'Sheet1!B9', formula: '=A9*10', cachedValue: 90, literalRecalculatedValue: 90, cacheStatus: 'fresh', staleCachedValue: false },",
      '  ] : [',
      "    { target: 'Sheet1!B2', formula: '=A2*10', cachedValue: 10, cacheStatus: 'unsupported-recalculation', staleCachedValue: null },",
      "    { target: 'Sheet1!B3', formula: '=A3*10', cachedValue: 999, literalRecalculatedValue: 30, cacheStatus: 'stale', staleCachedValue: true },",
      "    { target: 'Sheet1!B4', formula: '=A4&`|`', cachedValue: 'old\\n%value', literalRecalculatedValue: 'new|value', cacheStatus: 'stale', staleCachedValue: true },",
      '  ],',
      '  warnings: [],',
      '  inspectionCompleted: true,',
      '  recalculationCompleted: true,',
      "  excelParity: 'not_proven',",
      '  };',
      '}',
      '',
    ].join('\n'),
  )
  writeFileSync(join(fixtureDir, 'stale-pricing.xlsx'), '')
  writeFileSync(join(fixtureDir, 'fresh-pricing.xlsx'), '')

  const outputPath = join(inspectTempDir, 'github-output.txt')
  const summaryPath = join(inspectTempDir, 'summary.md')
  const reportPath = join(inspectTempDir, 'report.json')
  const markdownPath = join(inspectTempDir, 'report.md')
  const importMarkerPath = join(inspectTempDir, 'import-marker.txt')
  const result = spawnSync(process.execPath, [inspectScriptPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      GITHUB_WORKSPACE: inspectTempDir,
      GITHUB_OUTPUT: outputPath,
      GITHUB_STEP_SUMMARY: summaryPath,
      BILIG_WORKBOOKS_JSON: '["fixtures/stale-pricing.xlsx","fixtures/fresh-pricing.xlsx"]',
      BILIG_PACKAGE_ROOT: fakePackageRoot,
      BILIG_PACKAGE_VERSION: 'test',
      BILIG_INSPECT_LIMIT: 'all',
      BILIG_JSON_OUTPUT: reportPath,
      BILIG_MARKDOWN_OUTPUT: markdownPath,
      BILIG_FAIL_ON_STALE: 'false',
      BILIG_FAKE_IMPORT_MARKER: importMarkerPath,
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
  const outputs = parseGithubOutput(output)
  if (
    outputs['stale-count'] !== '2' ||
    outputs['fresh-count'] !== '1' ||
    outputs['missing-cache-count'] !== '0' ||
    outputs['unsupported-recalculation-count'] !== '1' ||
    outputs['markdown'] !== markdownPath ||
    outputs['suggested-reads'] !==
      'fixtures/stale-pricing.xlsx#Sheet1!B2,fixtures/stale-pricing.xlsx#Sheet1!B3,fixtures/fresh-pricing.xlsx#Sheet1!B9'
  ) {
    throw new Error(`Unexpected inspector outputs:\n${output}`)
  }
  const summary = await readFile(summaryPath, 'utf8')
  const markdown = await readFile(markdownPath, 'utf8')
  for (const expected of [
    '#### Stale cached formula values',
    '- Unsupported recalculation results: 1',
    '| fixtures/stale-pricing.xlsx | 3 | 2 | 0 | 0 | 1 | Sheet1!B2, Sheet1!B3 |',
    '| fixtures/fresh-pricing.xlsx | 1 | 0 | 1 | 0 | 0 | Sheet1!B9 |',
    '| fixtures/stale-pricing.xlsx | Sheet1!B3 | `=A3*10` | 999 | 30 |',
    '| fixtures/stale-pricing.xlsx | Sheet1!B4 | `=A4&\\`\\|\\`` | old %value | new\\|value |',
    '#### Follow-up check command',
    'npm exec --package @bilig/xlsx-formula-recalc@test -- xlsx-recalc',
    "xlsx-recalc 'fixtures/stale-pricing.xlsx' --read 'Sheet1!B3'",
  ]) {
    if (!summary.includes(expected) || !markdown.includes(expected)) {
      throw new Error(`Missing "${expected}" in inspector Markdown output:\nsummary:\n${summary}\nreport:\n${markdown}`)
    }
  }
  const importMarker = await readFile(importMarkerPath, 'utf8')
  if (importMarker.trim().split('\n').length !== 1) {
    throw new Error(`Expected @bilig/xlsx-formula-recalc to be imported once, got marker:\n${importMarker}`)
  }
  const aggregate = parseAggregateReport(JSON.parse(await readFile(reportPath, 'utf8')))
  if (
    aggregate.schemaVersion !== 'xlsx-cache-doctor-action.v1' ||
    aggregate.staleCachedFormulaCount !== 2 ||
    aggregate.cacheStatusSummary?.unsupportedRecalculation !== 1 ||
    aggregate.workbooks?.[0]?.staleFormulas?.length !== 2
  ) {
    throw new Error(`Unexpected inspector JSON report:\n${JSON.stringify(aggregate, null, 2)}`)
  }
} finally {
  rmSync(inspectTempDir, { recursive: true, force: true })
}

const emptyInspectTempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-cache-doctor-empty-'))
try {
  const outputPath = join(emptyInspectTempDir, 'github-output.txt')
  const summaryPath = join(emptyInspectTempDir, 'summary.md')
  const reportPath = join(emptyInspectTempDir, 'report.json')
  const markdownPath = join(emptyInspectTempDir, 'report.md')
  const result = spawnSync(process.execPath, [inspectScriptPath], {
    cwd: repoRoot,
    env: {
      ...process.env,
      GITHUB_WORKSPACE: emptyInspectTempDir,
      GITHUB_OUTPUT: outputPath,
      GITHUB_STEP_SUMMARY: summaryPath,
      BILIG_WORKBOOKS_JSON: '[]',
      BILIG_CHANGED_FILES_ONLY: 'true',
      BILIG_PACKAGE_VERSION: 'test',
      BILIG_INSPECT_LIMIT: 'all',
      BILIG_JSON_OUTPUT: reportPath,
      BILIG_MARKDOWN_OUTPUT: markdownPath,
      BILIG_FAIL_ON_STALE: 'false',
    },
    encoding: 'utf8',
  })
  if (result.status !== 0) {
    throw new Error(`XLSX Cache Doctor empty inspector smoke failed:\n${result.stdout}\n${result.stderr}`)
  }
  const output = await readFile(outputPath, 'utf8')
  const outputs = parseGithubOutput(output)
  if (outputs['workbook-count'] !== '0' || outputs['stale-count'] !== '0') {
    throw new Error(`Unexpected empty inspector outputs:\n${output}`)
  }
  const summary = await readFile(summaryPath, 'utf8')
  const markdown = await readFile(markdownPath, 'utf8')
  for (const expected of [
    '- Workbooks inspected: 0',
    '- Result: no XLSX workbooks were inspected.',
    '- Reason: changed-files-only matched no XLSX workbook changes.',
  ]) {
    if (!summary.includes(expected) || !markdown.includes(expected)) {
      throw new Error(`Missing "${expected}" in empty inspector Markdown output:\nsummary:\n${summary}\nreport:\n${markdown}`)
    }
  }
  const aggregate = parseEmptyAggregateReport(JSON.parse(await readFile(reportPath, 'utf8')))
  if (aggregate.workbookCount !== 0 || aggregate.skipReason !== 'changed-files-only matched no XLSX workbook changes') {
    throw new Error(`Unexpected empty inspector JSON report:\n${JSON.stringify(aggregate, null, 2)}`)
  }
} finally {
  rmSync(emptyInspectTempDir, { recursive: true, force: true })
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
  schemaVersion?: string
  staleCachedFormulaCount?: number
  cacheStatusSummary?: { unsupportedRecalculation?: number }
  workbooks?: Array<{ staleFormulas?: unknown[] }>
} {
  if (!isRecord(value)) {
    throw new Error(`Expected aggregate JSON report object, got ${JSON.stringify(value)}`)
  }
  const workbooksValue = Reflect.get(value, 'workbooks')
  const schemaVersion = Reflect.get(value, 'schemaVersion')
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
    schemaVersion: typeof schemaVersion === 'string' ? schemaVersion : undefined,
    staleCachedFormulaCount: typeof staleCachedFormulaCount === 'number' ? staleCachedFormulaCount : undefined,
    cacheStatusSummary,
    workbooks,
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function parseEmptyAggregateReport(value: unknown): { skipReason?: string; workbookCount?: number } {
  if (!isRecord(value)) {
    throw new Error(`Expected empty aggregate JSON report object, got ${JSON.stringify(value)}`)
  }
  const skipReason = Reflect.get(value, 'skipReason')
  const workbookCount = Reflect.get(value, 'workbookCount')
  return {
    skipReason: typeof skipReason === 'string' ? skipReason : undefined,
    workbookCount: typeof workbookCount === 'number' ? workbookCount : undefined,
  }
}

function parsePackageJson(raw: string): { version: string } {
  const value: unknown = JSON.parse(raw)
  if (!isRecord(value) || typeof value.version !== 'string') {
    throw new Error('Expected package.json with string version.')
  }
  return { version: value.version }
}

function parseGithubOutput(raw: string): Record<string, string> {
  const outputs: Record<string, string> = {}
  const lines = raw.split('\n')
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index]
    if (!line) {
      continue
    }
    const delimiterMatch = /^(?<name>[A-Za-z_][A-Za-z0-9_-]*)<<(?<delimiter>.+)$/u.exec(line)
    if (delimiterMatch?.groups) {
      const { name, delimiter } = delimiterMatch.groups
      const valueLines: string[] = []
      index += 1
      while (index < lines.length && lines[index] !== delimiter) {
        valueLines.push(lines[index])
        index += 1
      }
      outputs[name] = valueLines.join('\n')
      continue
    }
    const equalsIndex = line.indexOf('=')
    if (equalsIndex > 0) {
      outputs[line.slice(0, equalsIndex)] = line.slice(equalsIndex + 1)
    }
  }
  return outputs
}
