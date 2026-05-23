import type { ComparativeBenchmarkSuiteOptions } from './benchmark-workpaper-vs-hyperformula.js'
import { DEFAULT_COMPETITIVE_WARMUP_COUNT, DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT } from './benchmark-workpaper-vs-hyperformula.js'
import { runWorkPaperVsHyperFormulaExpandedBenchmarkSuite } from './benchmark-workpaper-vs-hyperformula-expanded-scenarios.js'
import type { ExpandedComparativeBenchmarkResult } from './benchmark-workpaper-vs-hyperformula-expanded-runner.js'
import { EXPANDED_COMPARATIVE_WORKLOADS, type ExpandedComparativeBenchmarkWorkload } from './expanded-competitive-workloads.js'
import { buildExpandedCompetitiveFamilyReport, type ExpandedCompetitiveFamilySummary } from './report-competitive-families.js'

export { EXPANDED_COMPARATIVE_WORKLOAD_SCORECARD_LANE, EXPANDED_COMPARATIVE_WORKLOADS } from './expanded-competitive-workloads.js'
export { DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT } from './benchmark-workpaper-vs-hyperformula.js'
export type { ExpandedComparativeBenchmarkWorkload } from './expanded-competitive-workloads.js'
export { runWorkPaperVsHyperFormulaExpandedBenchmarkSuite } from './benchmark-workpaper-vs-hyperformula-expanded-scenarios.js'
export type {
  EngineCounterNumericSummary,
  ExpandedComparativeBenchmarkResult,
  ExpandedComparativeComparableResult,
  ExpandedComparativeLeadershipResult,
} from './benchmark-workpaper-vs-hyperformula-expanded-runner.js'

export interface ExpandedComparativeBenchmarkReport {
  suite: 'workpaper-vs-hyperformula'
  results: readonly ExpandedComparativeBenchmarkResult[]
  families: readonly ExpandedCompetitiveFamilySummary[]
  scorecard: ReturnType<typeof buildExpandedCompetitiveFamilyReport>['scorecard']
}

export interface ExpandedBenchmarkCliOptions extends ComparativeBenchmarkSuiteOptions {
  jobs?: number
  workloads?: readonly ExpandedComparativeBenchmarkWorkload[]
}

export function buildExpandedComparativeBenchmarkReport(
  results: readonly ExpandedComparativeBenchmarkResult[],
): ExpandedComparativeBenchmarkReport {
  const familyReport = buildExpandedCompetitiveFamilyReport(results)
  return {
    suite: familyReport.suite,
    results: [...results],
    families: familyReport.families,
    scorecard: familyReport.scorecard,
  }
}

const expandedComparativeWorkloadSet: ReadonlySet<string> = new Set(EXPANDED_COMPARATIVE_WORKLOADS)

if (import.meta.url === `file://${process.argv[1]}`) {
  const cliOptions = parseExpandedBenchmarkCliOptions(process.argv.slice(2))
  const benchmarkResults = runWorkPaperVsHyperFormulaExpandedBenchmarkSuite({
    sampleCount: cliOptions.sampleCount ?? DEFAULT_EXPANDED_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: cliOptions.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
    ...(cliOptions.workloads ? { workloads: cliOptions.workloads } : {}),
  })
  console.log(JSON.stringify(buildExpandedComparativeBenchmarkReport(benchmarkResults), null, 2))
}

export function parseExpandedBenchmarkCliOptions(args: readonly string[]): ExpandedBenchmarkCliOptions {
  const options: {
    jobs?: number
    sampleCount?: number
    warmupCount?: number
    workloads?: ExpandedComparativeBenchmarkWorkload[]
  } = {}
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!
    if (arg === '--jobs') {
      const raw = args[index + 1]
      if (raw === undefined) {
        throw new Error('Missing value for --jobs')
      }
      options.jobs = parsePositiveDecimalInteger(raw, '--jobs')
      index += 1
      continue
    }
    if (arg === '--sample-count') {
      const raw = args[index + 1]
      if (raw === undefined) {
        throw new Error('Missing value for --sample-count')
      }
      options.sampleCount = parsePositiveDecimalInteger(raw, '--sample-count')
      index += 1
      continue
    }
    if (arg === '--warmup-count') {
      const raw = args[index + 1]
      if (raw === undefined) {
        throw new Error('Missing value for --warmup-count')
      }
      options.warmupCount = parseNonNegativeDecimalInteger(raw, '--warmup-count')
      index += 1
      continue
    }
    if (arg === '--workload') {
      const raw = args[index + 1]
      if (raw === undefined) {
        throw new Error('Missing value for --workload')
      }
      ;(options.workloads ??= []).push(parseExpandedComparativeBenchmarkWorkload(raw))
      index += 1
      continue
    }
    throw new Error(`Unknown expanded benchmark argument: ${arg}`)
  }
  return options
}

function parsePositiveDecimalInteger(value: string, option: string): number {
  const parsed = parseNonNegativeDecimalInteger(value, option)
  if (parsed < 1) {
    throw new Error(`${option} expects a positive integer, got ${value}`)
  }
  return parsed
}

function parseNonNegativeDecimalInteger(value: string, option: string): number {
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    throw new Error(`${option} expects a non-negative integer, got ${value}`)
  }
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed)) {
    throw new Error(`${option} expects a safe integer, got ${value}`)
  }
  return parsed
}

function parseExpandedComparativeBenchmarkWorkload(value: string): ExpandedComparativeBenchmarkWorkload {
  if (isExpandedComparativeBenchmarkWorkload(value)) {
    return value
  }
  throw new Error(`Unknown expanded benchmark workload: ${value}`)
}

function isExpandedComparativeBenchmarkWorkload(value: string): value is ExpandedComparativeBenchmarkWorkload {
  return expandedComparativeWorkloadSet.has(value)
}
