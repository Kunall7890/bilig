#!/usr/bin/env bun

import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'

export interface PerfSmokeBenchmarkResult {
  readonly elapsedMs: number
  readonly downstreamCount: number
  readonly metrics: {
    readonly changedInputCount: number
    readonly dirtyFormulaCount: number
    readonly wasmFormulaCount: number
    readonly jsFormulaCount: number
  }
  readonly performanceCounters: {
    readonly directScalarDeltaApplications: number
  }
  readonly verification: {
    readonly terminalAddress: string
    readonly expectedTerminalValue: number
    readonly terminalValue: number | null
  }
}

export interface PerfSmokeDependencies {
  readonly runBenchmark: (downstreamCount?: number) => Promise<PerfSmokeBenchmarkResult>
  readonly buildWasm: () => Promise<void>
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isNumberOrNull(value: unknown): value is number | null {
  return typeof value === 'number' || value === null
}

function isPerfSmokeBenchmarkResult(value: unknown): value is PerfSmokeBenchmarkResult {
  if (!isRecord(value)) {
    return false
  }
  const metrics = value['metrics']
  const performanceCounters = value['performanceCounters']
  const verification = value['verification']
  return (
    typeof value['elapsedMs'] === 'number' &&
    typeof value['downstreamCount'] === 'number' &&
    isRecord(metrics) &&
    typeof metrics['changedInputCount'] === 'number' &&
    typeof metrics['dirtyFormulaCount'] === 'number' &&
    typeof metrics['wasmFormulaCount'] === 'number' &&
    typeof metrics['jsFormulaCount'] === 'number' &&
    isRecord(performanceCounters) &&
    typeof performanceCounters['directScalarDeltaApplications'] === 'number' &&
    isRecord(verification) &&
    typeof verification['terminalAddress'] === 'string' &&
    typeof verification['expectedTerminalValue'] === 'number' &&
    isNumberOrNull(verification['terminalValue'])
  )
}

function benchmarkEditScriptPath(): string {
  return fileURLToPath(new URL('../packages/benchmarks/src/benchmark-edit.ts', import.meta.url))
}

async function spawnCommand(command: string, args: readonly string[]): Promise<string> {
  return await new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: process.cwd(),
      env: process.env,
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    let stdout = ''
    let stderr = ''

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString()
    })
    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString()
    })
    child.on('error', reject)
    child.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(stderr || stdout || `${command} exited with code ${String(code)}`))
        return
      }
      resolve(stdout)
    })
  })
}

export async function buildWasmKernelForPerfSmoke(): Promise<void> {
  await spawnCommand('pnpm', ['wasm:build'])
}

export async function runPerfSmokeBenchmark(downstreamCount = 1_000): Promise<PerfSmokeBenchmarkResult> {
  const stdout = await spawnCommand('node', ['--import', 'tsx', benchmarkEditScriptPath(), String(downstreamCount)])

  try {
    const parsed = JSON.parse(stdout) as unknown
    if (!isPerfSmokeBenchmarkResult(parsed)) {
      throw new Error('Perf smoke benchmark output did not match the expected shape')
    }
    return parsed
  } catch (error) {
    throw new Error(`Failed to parse perf smoke benchmark output: ${error instanceof Error ? error.message : String(error)}`, {
      cause: error,
    })
  }
}

export async function runPerfSmokeGate(
  downstreamCount = 1_000,
  dependencies: PerfSmokeDependencies = {
    runBenchmark: runPerfSmokeBenchmark,
    buildWasm: buildWasmKernelForPerfSmoke,
  },
): Promise<PerfSmokeBenchmarkResult> {
  const firstPass = await dependencies.runBenchmark(downstreamCount)
  if (usesSupportedFormulaPath(firstPass)) {
    return firstPass
  }
  await dependencies.buildWasm()
  return await dependencies.runBenchmark(downstreamCount)
}

function getDownstreamPropagationCount(result: PerfSmokeBenchmarkResult): number {
  return Math.max(result.metrics.dirtyFormulaCount, result.performanceCounters.directScalarDeltaApplications)
}

function usesDirectScalarPath(result: PerfSmokeBenchmarkResult): boolean {
  return result.performanceCounters.directScalarDeltaApplications >= result.downstreamCount
}

function usesSupportedFormulaPath(result: PerfSmokeBenchmarkResult): boolean {
  return result.metrics.wasmFormulaCount > 0 || usesDirectScalarPath(result)
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const result = await runPerfSmokeGate()
  const { elapsedMs: elapsed, metrics, downstreamCount, verification } = result

  if (elapsed > 250) {
    console.warn(`perf smoke exceeded threshold: ${elapsed.toFixed(2)}ms`)
    process.exit(1)
  }

  if (metrics.changedInputCount !== 1) {
    console.warn(`perf smoke expected exactly one edited input, got ${metrics.changedInputCount}`)
    process.exit(1)
  }

  if (verification.terminalValue !== verification.expectedTerminalValue) {
    console.warn(
      `perf smoke terminal readback mismatch at ${verification.terminalAddress}: expected ${verification.expectedTerminalValue}, got ${String(verification.terminalValue)}`,
    )
    process.exit(1)
  }

  const downstreamPropagationCount = getDownstreamPropagationCount(result)
  if (downstreamPropagationCount < downstreamCount) {
    console.warn(
      `perf smoke failed to update the expected downstream formulas: expected at least ${downstreamCount}, got ${downstreamPropagationCount}`,
    )
    process.exit(1)
  }

  if (metrics.jsFormulaCount !== 0) {
    console.warn(`perf smoke fell back to js formula evaluation for ${metrics.jsFormulaCount} formulas`)
    process.exit(1)
  }

  if (!usesSupportedFormulaPath(result)) {
    console.warn('perf smoke did not exercise a supported formula fast path')
    process.exit(1)
  }

  console.log(JSON.stringify(result, null, 2))
}
