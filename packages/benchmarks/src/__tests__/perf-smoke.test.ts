import { describe, expect, it, vi } from 'vitest'
import { runPerfSmokeBenchmark, runPerfSmokeGate, type PerfSmokeBenchmarkResult } from '../../../../scripts/perf-smoke.ts'

describe('perf smoke', () => {
  it('runs the benchmark edit scenario through node and tsx', async () => {
    const result = await runPerfSmokeBenchmark(100)

    expect(result.downstreamCount).toBe(100)
    expect(result.elapsedMs).toBeGreaterThanOrEqual(0)
    expect(result.metrics.changedInputCount).toBe(1)
    expect(Math.max(result.metrics.dirtyFormulaCount, result.performanceCounters.directScalarDeltaApplications)).toBeGreaterThanOrEqual(100)
    expect(result.metrics.wasmFormulaCount).toBeGreaterThanOrEqual(0)
    expect(result.metrics.jsFormulaCount).toBe(0)
    expect(result.verification).toEqual({
      terminalAddress: 'B100',
      terminalValue: 298,
      expectedTerminalValue: 298,
    })
  }, 30_000)

  it('retries once after building wasm when the first pass falls back to js', async () => {
    const jsOnly: PerfSmokeBenchmarkResult = {
      elapsedMs: 5,
      downstreamCount: 100,
      metrics: {
        changedInputCount: 1,
        dirtyFormulaCount: 100,
        wasmFormulaCount: 0,
        jsFormulaCount: 100,
      },
      performanceCounters: {
        directScalarDeltaApplications: 0,
      },
      verification: {
        terminalAddress: 'B100',
        expectedTerminalValue: 298,
        terminalValue: 298,
      },
    }
    const wasmReady: PerfSmokeBenchmarkResult = {
      ...jsOnly,
      metrics: {
        ...jsOnly.metrics,
        wasmFormulaCount: 100,
        jsFormulaCount: 0,
      },
    }
    const runBenchmark = vi.fn(async () => (runBenchmark.mock.calls.length === 1 ? jsOnly : wasmReady))
    const buildWasm = vi.fn(async () => {})

    const result = await runPerfSmokeGate(100, {
      runBenchmark,
      buildWasm,
    })

    expect(result).toEqual(wasmReady)
    expect(runBenchmark).toHaveBeenCalledTimes(2)
    expect(buildWasm).toHaveBeenCalledTimes(1)
  })

  it('accepts direct scalar propagation without forcing an unnecessary wasm build', async () => {
    const directScalarReady: PerfSmokeBenchmarkResult = {
      elapsedMs: 5,
      downstreamCount: 100,
      metrics: {
        changedInputCount: 1,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
      },
      performanceCounters: {
        directScalarDeltaApplications: 100,
      },
      verification: {
        terminalAddress: 'B100',
        expectedTerminalValue: 298,
        terminalValue: 298,
      },
    }
    const runBenchmark = vi.fn(async () => directScalarReady)
    const buildWasm = vi.fn(async () => {})

    const result = await runPerfSmokeGate(100, {
      runBenchmark,
      buildWasm,
    })

    expect(result).toEqual(directScalarReady)
    expect(runBenchmark).toHaveBeenCalledTimes(1)
    expect(buildWasm).not.toHaveBeenCalled()
  })

  it('does not build wasm when the first pass already uses the fast path', async () => {
    const wasmReady: PerfSmokeBenchmarkResult = {
      elapsedMs: 5,
      downstreamCount: 100,
      metrics: {
        changedInputCount: 1,
        dirtyFormulaCount: 100,
        wasmFormulaCount: 100,
        jsFormulaCount: 0,
      },
      performanceCounters: {
        directScalarDeltaApplications: 0,
      },
      verification: {
        terminalAddress: 'B100',
        expectedTerminalValue: 298,
        terminalValue: 298,
      },
    }
    const runBenchmark = vi.fn(async () => wasmReady)
    const buildWasm = vi.fn(async () => {})

    const result = await runPerfSmokeGate(100, {
      runBenchmark,
      buildWasm,
    })

    expect(result).toEqual(wasmReady)
    expect(runBenchmark).toHaveBeenCalledTimes(1)
    expect(buildWasm).not.toHaveBeenCalled()
  })

  it('does not build wasm when the direct formula fast path is semantically verified', async () => {
    const directReady: PerfSmokeBenchmarkResult = {
      elapsedMs: 5,
      downstreamCount: 100,
      metrics: {
        changedInputCount: 1,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
      },
      performanceCounters: {
        directScalarDeltaApplications: 0,
      },
      verification: {
        terminalAddress: 'B100',
        terminalValue: 298,
        expectedTerminalValue: 298,
      },
    }
    const runBenchmark = vi.fn(async () => directReady)
    const buildWasm = vi.fn(async () => {})

    const result = await runPerfSmokeGate(100, {
      runBenchmark,
      buildWasm,
    })

    expect(result).toEqual(directReady)
    expect(runBenchmark).toHaveBeenCalledTimes(1)
    expect(buildWasm).not.toHaveBeenCalled()
  })
})
