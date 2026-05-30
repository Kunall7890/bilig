#!/usr/bin/env node

import { appendFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { spawnSync } from 'node:child_process'

const workbooks = JSON.parse(process.env.BILIG_WORKBOOKS_JSON || '[]')
const packageVersion = process.env.BILIG_PACKAGE_VERSION || 'latest'
const inspectLimit = process.env.BILIG_INSPECT_LIMIT || 'all'
const outputPath = process.env.BILIG_JSON_OUTPUT || join(process.env.RUNNER_TEMP || process.cwd(), 'bilig-xlsx-cache-doctor.json')
const failOnStale = process.env.BILIG_FAIL_ON_STALE === 'true'
const reports = workbooks.map((workbook) => inspectWorkbook(workbook))
const aggregate = buildAggregateReport(reports)

mkdirSync(dirname(outputPath), { recursive: true })
writeFileSync(outputPath, `${JSON.stringify(aggregate, null, 2)}\n`)

writeGithubOutput('json', outputPath)
writeGithubOutput('workbook-count', String(aggregate.workbookCount))
writeGithubOutput('formula-count', String(aggregate.formulaCellCount))
writeGithubOutput('stale-count', String(aggregate.staleCachedFormulaCount))
writeGithubOutput('uninspected-count', String(aggregate.uninspectedFormulaCellCount))
writeGithubOutput('suggested-reads', aggregate.suggestedReads.join(','))
writeStepSummary(aggregate)

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

function writeStepSummary(report) {
  if (!process.env.GITHUB_STEP_SUMMARY) {
    return
  }
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
  appendFileSync(process.env.GITHUB_STEP_SUMMARY, `${lines.join('\n')}\n`)
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
  return value.replaceAll('|', '\\|')
}

function writeGithubOutput(name, value) {
  appendFileSync(process.env.GITHUB_OUTPUT, `${name}=${value}\n`)
}
