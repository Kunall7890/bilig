#!/usr/bin/env node

import { appendFileSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, isAbsolute, join } from 'node:path'
import { pathToFileURL } from 'node:url'

const workbooks = JSON.parse(process.env.BILIG_WORKBOOKS_JSON || '[]')
const workspace = process.env.GITHUB_WORKSPACE || process.cwd()
const packageRoot = process.env.BILIG_PACKAGE_ROOT || ''
const packageVersion = process.env.BILIG_PACKAGE_VERSION || 'latest'
const inspectLimit = process.env.BILIG_INSPECT_LIMIT || 'all'
const outputPath = process.env.BILIG_JSON_OUTPUT || join(process.env.RUNNER_TEMP || process.cwd(), 'bilig-xlsx-cache-doctor.json')
const markdownOutputPath = process.env.BILIG_MARKDOWN_OUTPUT || join(process.env.RUNNER_TEMP || process.cwd(), 'bilig-xlsx-cache-doctor.md')
const failOnStale = process.env.BILIG_FAIL_ON_STALE === 'true'
const changedFilesOnly = process.env.BILIG_CHANGED_FILES_ONLY === 'true'
const actionSchemaVersion = 'xlsx-cache-doctor-action.v1'
const reports = await inspectWorkbooks(workbooks)
const aggregate = buildAggregateReport(reports)

mkdirSync(dirname(outputPath), { recursive: true })
writeFileSync(outputPath, `${JSON.stringify(aggregate, null, 2)}\n`)
const markdownSummary = buildMarkdownSummary(aggregate)
mkdirSync(dirname(markdownOutputPath), { recursive: true })
writeFileSync(markdownOutputPath, `${markdownSummary}\n`)

writeGithubOutput('json', outputPath)
writeGithubOutput('markdown', markdownOutputPath)
writeGithubOutput('workbook-count', String(aggregate.workbookCount))
writeGithubOutput('formula-count', String(aggregate.formulaCellCount))
writeGithubOutput('stale-count', String(aggregate.staleCachedFormulaCount))
writeGithubOutput('fresh-count', String(aggregate.cacheStatusSummary.fresh))
writeGithubOutput('missing-cache-count', String(aggregate.cacheStatusSummary.missingCache))
writeGithubOutput('unsupported-recalculation-count', String(aggregate.cacheStatusSummary.unsupportedRecalculation))
writeGithubOutput('uninspected-count', String(aggregate.uninspectedFormulaCellCount))
writeGithubOutput('suggested-reads', aggregate.suggestedReads.slice(0, 25).join(','))
writeStaleAnnotations(aggregate)
writeStepSummary(markdownSummary)

if (failOnStale && aggregate.staleCachedFormulaCount > 0) {
  console.error(`xlsx-cache-doctor found ${aggregate.staleCachedFormulaCount.toString()} stale cached formula value(s).`)
  process.exit(2)
}

async function inspectWorkbooks(workbookPaths) {
  if (workbookPaths.length === 0) {
    return []
  }
  const inspectXlsxCache = await loadInspectXlsxCache()
  return workbookPaths.map((workbook) => inspectWorkbook(workbook, inspectXlsxCache))
}

function inspectWorkbook(workbook, inspectXlsxCache) {
  const workbookPath = isAbsolute(workbook) ? workbook : join(workspace, workbook)
  return {
    workbook,
    report: inspectXlsxCache(readFileSync(workbookPath), {
      fileName: workbook,
      inspectLimit: parseInspectLimit(inspectLimit),
    }),
  }
}

async function loadInspectXlsxCache() {
  if (!packageRoot) {
    throw new Error(
      'BILIG_PACKAGE_ROOT is required when XLSX workbooks are inspected. The action must install @bilig/xlsx-formula-recalc first.',
    )
  }
  const moduleUrl = pathToFileURL(join(packageRoot, 'dist', 'index.js')).href
  const module = await import(moduleUrl)
  if (typeof module.inspectXlsxCache !== 'function') {
    throw new Error(`@bilig/xlsx-formula-recalc at ${packageRoot} does not export inspectXlsxCache.`)
  }
  return module.inspectXlsxCache
}

function parseInspectLimit(raw) {
  if (raw === 'all') {
    return raw
  }
  if (/^[1-9]\d*$/u.test(raw)) {
    return Number(raw)
  }
  throw new Error(`inspect-limit must be "all" or a positive integer, got ${JSON.stringify(raw)}.`)
}

function buildAggregateReport(items) {
  const suggestedReads = items.flatMap((item) => readStringArray(item.report.suggestedReads).map((target) => `${item.workbook}#${target}`))
  const workbookReports = items.map((item) => {
    const cacheStatusSummary = readCacheStatusSummary(item.report)
    return {
      workbook: item.workbook,
      formulaCellCount: numberField(item.report.formulaCellCount),
      inspectedFormulaCellCount: numberField(item.report.inspectedFormulaCellCount),
      uninspectedFormulaCellCount: numberField(item.report.uninspectedFormulaCellCount),
      staleCachedFormulaCount: numberField(item.report.staleCachedFormulaCount),
      cacheStatusSummary,
      suggestedReads: readStringArray(item.report.suggestedReads),
      staleFormulas: readFormulaArray(item.report.formulas)
        .filter((formula) => formula.staleCachedValue === true)
        .map((formula) => ({
          target: formula.target,
          formula: formula.formula,
          cachedValue: formula.cachedValue,
          literalRecalculatedValue: formula.literalRecalculatedValue,
        })),
      warnings: readStringArray(item.report.warnings),
    }
  })
  return {
    schemaVersion: actionSchemaVersion,
    mode: 'github-action',
    packageVersion,
    inspectLimit,
    workbookCount: items.length,
    workbooks: workbookReports,
    formulaCellCount: sum(items, (item) => numberField(item.report.formulaCellCount)),
    inspectedFormulaCellCount: sum(items, (item) => numberField(item.report.inspectedFormulaCellCount)),
    uninspectedFormulaCellCount: sum(items, (item) => numberField(item.report.uninspectedFormulaCellCount)),
    staleCachedFormulaCount: sum(items, (item) => numberField(item.report.staleCachedFormulaCount)),
    cacheStatusSummary: sumCacheStatusSummaries(workbookReports.map((workbook) => workbook.cacheStatusSummary)),
    suggestedReads,
    ...(items.length === 0
      ? { skipReason: changedFilesOnly ? 'changed-files-only matched no XLSX workbook changes' : 'no XLSX workbooks resolved' }
      : {}),
    commandSucceeded: true,
    inspectionCompleted: true,
    recalculationCompleted: true,
    excelParity: 'not_proven',
  }
}

function buildMarkdownSummary(report) {
  const staleFormulas = report.workbooks.flatMap((workbook) =>
    workbook.staleFormulas.map((formula) => ({
      workbook: workbook.workbook,
      ...formula,
    })),
  )
  const lines = [
    '### XLSX cache doctor',
    '',
    `- Workbooks inspected: ${report.workbookCount.toString()}`,
    `- Formula cells: ${report.formulaCellCount.toString()}`,
    `- Inspected formula cells: ${report.inspectedFormulaCellCount.toString()}`,
    `- Uninspected formula cells: ${report.uninspectedFormulaCellCount.toString()}`,
    `- Stale cached values: ${report.staleCachedFormulaCount.toString()}`,
    `- Fresh cached values: ${report.cacheStatusSummary.fresh.toString()}`,
    `- Missing cached values: ${report.cacheStatusSummary.missingCache.toString()}`,
    `- Unsupported recalculation results: ${report.cacheStatusSummary.unsupportedRecalculation.toString()}`,
    `- JSON report: ${outputPath}`,
    '',
  ]
  if (report.workbookCount === 0) {
    lines.push('- Result: no XLSX workbooks were inspected.')
    if (report.skipReason) {
      lines.push(`- Reason: ${report.skipReason}.`)
    }
    lines.push('')
  }
  if (report.workbooks.length > 0) {
    lines.push('| Workbook | Formula cells | Stale | Fresh | Missing cache | Unsupported | Suggested reads |')
    lines.push('| --- | ---: | ---: | ---: | ---: | ---: | --- |')
    for (const workbook of report.workbooks) {
      lines.push(
        `| ${escapeMarkdown(workbook.workbook)} | ${workbook.formulaCellCount.toString()} | ${workbook.staleCachedFormulaCount.toString()} | ${workbook.cacheStatusSummary.fresh.toString()} | ${workbook.cacheStatusSummary.missingCache.toString()} | ${workbook.cacheStatusSummary.unsupportedRecalculation.toString()} | ${
          workbook.suggestedReads.slice(0, 5).map(escapeMarkdown).join(', ') || 'none'
        } |`,
      )
    }
    lines.push('')
  }
  if (staleFormulas.length > 0) {
    lines.push('#### Stale cached formula values')
    lines.push('')
    lines.push('| Workbook | Cell | Formula | Cached value | Recalculated value |')
    lines.push('| --- | --- | --- | --- | --- |')
    for (const formula of staleFormulas.slice(0, 15)) {
      lines.push(
        `| ${escapeMarkdown(formula.workbook)} | ${escapeMarkdown(formula.target)} | \`${escapeInlineCode(formatValue(formula.formula))}\` | ${escapeMarkdown(
          formatValue(formula.cachedValue),
        )} | ${escapeMarkdown(formatValue(formula.literalRecalculatedValue))} |`,
      )
    }
    if (staleFormulas.length > 15) {
      lines.push(`| ... | ... | ... | ... | ${String(staleFormulas.length - 15)} more stale value(s) in the JSON report. |`)
    }
    lines.push('')
  }
  const followUpCommand = buildFollowUpCommand(report)
  if (followUpCommand) {
    lines.push('#### Follow-up check command')
    lines.push('')
    lines.push('```sh')
    lines.push(followUpCommand)
    lines.push('```')
    lines.push('')
  }
  return lines.join('\n')
}

function writeStepSummary(summaryMarkdown) {
  if (!process.env.GITHUB_STEP_SUMMARY) {
    return
  }
  appendFileSync(process.env.GITHUB_STEP_SUMMARY, `${summaryMarkdown}\n`)
}

function writeStaleAnnotations(report) {
  const staleFormulas = report.workbooks.flatMap((workbook) =>
    workbook.staleFormulas.map((formula) => ({
      workbook: workbook.workbook,
      ...formula,
    })),
  )
  for (const formula of staleFormulas.slice(0, 10)) {
    console.log(
      [
        '::warning title=Stale cached XLSX formula::',
        escapeWorkflowCommand(
          `${formula.workbook}#${formula.target} cached ${formatValue(formula.cachedValue)} but recalculated ${formatValue(
            formula.literalRecalculatedValue,
          )}`,
        ),
      ].join(''),
    )
  }
  if (staleFormulas.length > 10) {
    console.log(
      `::warning title=Stale cached XLSX formula::${escapeWorkflowCommand(
        `${String(staleFormulas.length - 10)} more stale value(s) in ${outputPath}`,
      )}`,
    )
  }
}

function buildFollowUpCommand(report) {
  const workbook =
    report.workbooks.find((item) => item.staleFormulas.length > 0) || report.workbooks.find((item) => item.suggestedReads.length > 0)
  if (!workbook) {
    return undefined
  }
  const readTarget = workbook.staleFormulas[0]?.target || workbook.suggestedReads[0]
  if (!readTarget) {
    return undefined
  }
  return [
    `npm exec --package @bilig/xlsx-formula-recalc@${packageVersion} -- xlsx-recalc`,
    shellQuote(workbook.workbook),
    '--read',
    shellQuote(readTarget),
    '--out',
    shellQuote(recalculatedWorkbookPath(workbook.workbook)),
    '--json',
  ].join(' ')
}

function recalculatedWorkbookPath(workbook) {
  return workbook.replace(/\.xlsx$/iu, '.recalculated.xlsx')
}

function formatValue(value) {
  if (value === undefined) {
    return '(missing)'
  }
  if (value === null) {
    return 'null'
  }
  if (typeof value === 'string') {
    return value
  }
  return JSON.stringify(value)
}

function escapeInlineCode(value) {
  return value.replaceAll('`', '\\`').replaceAll('|', '\\|')
}

function escapeWorkflowCommand(value) {
  return value.replaceAll('%', '%25').replaceAll('\r', '%0D').replaceAll('\n', '%0A')
}

function shellQuote(value) {
  return `'${value.replaceAll("'", "'\\''")}'`
}

function numberField(value) {
  return typeof value === 'number' && Number.isFinite(value) ? value : 0
}

function readStringArray(value) {
  return Array.isArray(value) ? value.filter((item) => typeof item === 'string') : []
}

function readFormulaArray(value) {
  return Array.isArray(value) ? value.filter((item) => typeof item === 'object' && item !== null) : []
}

function readCacheStatusSummary(report) {
  if (typeof report.cacheStatusSummary === 'object' && report.cacheStatusSummary !== null) {
    return {
      inspected: numberField(report.cacheStatusSummary.inspected),
      stale: numberField(report.cacheStatusSummary.stale),
      fresh: numberField(report.cacheStatusSummary.fresh),
      missingCache: numberField(report.cacheStatusSummary.missingCache),
      unsupportedRecalculation: numberField(report.cacheStatusSummary.unsupportedRecalculation),
    }
  }
  const formulas = readFormulaArray(report.formulas)
  return {
    inspected: formulas.length,
    stale: formulas.filter((formula) => cacheStatusForFormula(formula) === 'stale').length,
    fresh: formulas.filter((formula) => cacheStatusForFormula(formula) === 'fresh').length,
    missingCache: formulas.filter((formula) => cacheStatusForFormula(formula) === 'missing-cache').length,
    unsupportedRecalculation: formulas.filter((formula) => cacheStatusForFormula(formula) === 'unsupported-recalculation').length,
  }
}

function cacheStatusForFormula(formula) {
  if (
    formula.cacheStatus === 'fresh' ||
    formula.cacheStatus === 'stale' ||
    formula.cacheStatus === 'missing-cache' ||
    formula.cacheStatus === 'unsupported-recalculation'
  ) {
    return formula.cacheStatus
  }
  if (formula.staleCachedValue === true) {
    return 'stale'
  }
  if (formula.staleCachedValue === false) {
    return 'fresh'
  }
  if (!Object.hasOwn(formula, 'cachedValue')) {
    return 'missing-cache'
  }
  return 'unsupported-recalculation'
}

function sumCacheStatusSummaries(summaries) {
  return {
    inspected: sum(summaries, (summary) => summary.inspected),
    stale: sum(summaries, (summary) => summary.stale),
    fresh: sum(summaries, (summary) => summary.fresh),
    missingCache: sum(summaries, (summary) => summary.missingCache),
    unsupportedRecalculation: sum(summaries, (summary) => summary.unsupportedRecalculation),
  }
}

function sum(items, read) {
  return items.reduce((total, item) => total + read(item), 0)
}

function escapeMarkdown(value) {
  return value.replaceAll('\r', ' ').replaceAll('\n', ' ').replaceAll('|', '\\|')
}

function writeGithubOutput(name, value) {
  const githubOutputPath = process.env.GITHUB_OUTPUT
  if (!githubOutputPath) {
    return
  }
  const delimiter = githubOutputDelimiter(name, value)
  appendFileSync(githubOutputPath, `${name}<<${delimiter}\n${value}\n${delimiter}\n`)
}

function githubOutputDelimiter(name, value) {
  const safeName = name.replaceAll(/[^A-Za-z0-9_]/gu, '_')
  let delimiter = `BILIG_${safeName}_OUTPUT`
  while (value.includes(delimiter)) {
    delimiter = `${delimiter}_END`
  }
  return delimiter
}
