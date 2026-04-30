import { performance } from 'node:perf_hooks'
import { SpreadsheetEngine, type EngineCounters } from '@bilig/core'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import { seedDownstreamWorkbook } from './generate-workbook.js'
import { measureMemory, sampleMemory, type MemoryMeasurement } from './metrics.js'

export interface EditBenchmarkResult {
  scenario: 'single-edit'
  downstreamCount: number
  elapsedMs: number
  metrics: RecalcMetrics
  performanceCounters: EngineCounters
  verification: {
    terminalAddress: string
    expectedTerminalValue: number
    terminalValue: number | null
  }
  memory: MemoryMeasurement
}

export async function runEditBenchmark(downstreamCount = 10_000): Promise<EditBenchmarkResult> {
  const engine = new SpreadsheetEngine({ workbookName: 'benchmark-edit' })
  await engine.ready()
  seedDownstreamWorkbook(engine, downstreamCount)

  engine.resetPerformanceCounters()
  const memoryBefore = sampleMemory()
  const started = performance.now()
  engine.setCellValue('Sheet1', 'A1', 99)
  const elapsed = performance.now() - started
  const memoryAfter = sampleMemory()
  const terminalAddress = `B${downstreamCount}`
  const terminalCellValue = engine.getCellValue('Sheet1', terminalAddress)
  const terminalValue = terminalCellValue.tag === ValueTag.Number ? terminalCellValue.value : null

  return {
    scenario: 'single-edit',
    downstreamCount,
    elapsedMs: elapsed,
    metrics: engine.getLastMetrics(),
    performanceCounters: engine.getPerformanceCounters(),
    verification: {
      terminalAddress,
      expectedTerminalValue: 99 * 2 + downstreamCount,
      terminalValue,
    },
    memory: measureMemory(memoryBefore, memoryAfter),
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const downstreamCount = Number.parseInt(process.argv[2] ?? '10000', 10)
  console.log(JSON.stringify(await runEditBenchmark(downstreamCount), null, 2))
}
