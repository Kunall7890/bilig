import type { WorkbookScrollPerfFixture } from '../apps/web/src/perf/workbook-scroll-perf.js'

export interface BiligScrollPerfBenchmarkState {
  readonly error: string | null
  readonly fixture: WorkbookScrollPerfFixture | null
  readonly state: string
}

export interface BiligScrollPerfWindowCollector {
  readonly getBenchmarkState?: () => BiligScrollPerfBenchmarkState
}

declare global {
  interface Window {
    __biligScrollPerf?: BiligScrollPerfWindowCollector
  }
}
