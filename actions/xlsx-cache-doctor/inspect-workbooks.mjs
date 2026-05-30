#!/usr/bin/env node

import { appendFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { spawnSync } from 'node:child_process'

const workbooks = JSON.parse(process.env.BILIG_WORKBOOKS_JSON || '[]')
const packageVersion = process.env.BILIG_PACKAGE_VERSION || 'latest'
const inspectLimit = process.env.BILIG_INSPECT_LIMIT || 'all'
const outputPath = process.env.BILIG_JSON_OUTPUT || join(process.env.RUNNER_TEMP || process.cwd(), 'bilig-xlsx-cache-doctor.json')
const markdownOutputPath = process.env.BILIG_MARKDOWN_OUTPUT || join(process.env.RUNNER_TEMP || process.cwd(), 'bilig-xlsx-cache-doctor.md')
const failOnStale = process.env.BILIG_FAIL_ON_STALE === 'true'
const reports = workbooks.map((workbook) => inspectWorkbook(workbook))
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
writeGithubOutput('uninspected-count', String(aggregate.uninspectedFormulaCellCount))
writeGithubOutput('suggested-reads', aggregate.suggestedReads.slice(0, 25).join(','))
writeStaleAnnotations(aggregate)
writeStepSummary(markdownSummary)

if (failOnStale && aggregate.staleCachedFormulaCount > 0) {
  console.error(`xlsx-cache-doctor found ${aggregate.staleCachedFormulaCount.toString()} stale cached formula value(s).`)
  process.exit(2)
}

function inspectWorkbook(workbook) {
  const result = spawnSync(
    'npm',
    [
      'exec',
      '--yes',
      '--package',
      `@bilig/xlsx-formula-recalc@${packageVersion}`,
      '--',
      'xlsx-cache-doctor',
      workbook,
      '--inspect-limit',
      inspectLimit,
      '--json',
    ],
    {
      encoding: 'utf8',
    },
  )
  if (result.status !== 0) {
    process.stdout.write(result.stdout)
    process.stderr.write(result.stderr)
    process.exit(result.status || 1)
  }
  return {
    workbook,
    report: JSON.parse(result.stdout),
  }
}

function buildAggregateReport(items) {
  const suggestedReads = items.flatMap((item) => readStringArray(item.report.suggestedReads).map((target) => `${item.workbook}#${target}`))
  return {
    mode: 'github-action',
    packageVersion,
    inspectLimit,
    workbookCount: items.length,
    workbooks: items.map((item) => ({
      workbook: item.workbook,
      formulaCellCount: numberField(item.report.formulaCellCount),
      inspectedFormulaCellCount: numberField(item.report.inspectedFormulaCellCount),
      uninspectedFormulaCellCount: numberField(item.report.uninspectedFormulaCellCount),
      staleCachedFormulaCount: numberField(item.report.staleCachedFormulaCount),
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
    })),
    formulaCellCount: sum(items, (item) => numberField(item.report.formulaCellCount)),
    inspectedFormulaCellCount: sum(items, (item) => numberField(item.report.inspectedFormulaCellCount)),
    uninspectedFormulaCellCount: sum(items, (item) => numberField(item.report.uninspectedFormulaCellCount)),
    staleCachedFormulaCount: sum(items, (item) => numberField(item.report.staleCachedFormulaCount)),
    suggestedReads,
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
    `- JSON report: ${outputPath}`,
    '',
  ]
  if (report.workbooks.length > 0) {
    lines.push('| Workbook | Formula cells | Stale cached values | Suggested reads |')
    lines.push('| --- | ---: | ---: | --- |')
    for (const workbook of report.workbooks) {
      lines.push(
        `| ${escapeMarkdown(workbook.workbook)} | ${workbook.formulaCellCount.toString()} | ${workbook.staleCachedFormulaCount.toString()} | ${
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
    'npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc',
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

function sum(items, read) {
  return items.reduce((total, item) => total + read(item), 0)
}

function escapeMarkdown(value) {
  return value.replaceAll('\r', ' ').replaceAll('\n', ' ').replaceAll('|', '\\|')
}

function writeGithubOutput(name, value) {
  appendFileSync(process.env.GITHUB_OUTPUT, `${name}=${value}\n`)
}
