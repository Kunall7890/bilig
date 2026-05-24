import { performance } from 'node:perf_hooks'

import type { UniverRuntime, WorkPaperUniverScenario } from './benchmark-workpaper-vs-univer.js'

export async function waitForUniverVerification(
  runtime: UniverRuntime,
  scenario: WorkPaperUniverScenario,
  expectedVerification: Record<string, unknown>,
  timeoutMs: number,
  operationResult?: unknown,
): Promise<Record<string, unknown>> {
  const expected = JSON.stringify(expectedVerification)
  const deadline = performance.now() + timeoutMs
  let actualVerification = scenario.verifyUniver(runtime, operationResult)
  while (JSON.stringify(actualVerification) !== expected) {
    if (performance.now() >= deadline) {
      throw new Error(
        `Timed out waiting for Univer verification for ${scenario.fixture.formula}: expected ${expected}, received ${JSON.stringify(actualVerification)}`,
      )
    }
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Polling waits sequentially until Univer exposes the expected calculated values.
    await new Promise((resolve) => setTimeout(resolve, 1))
    actualVerification = scenario.verifyUniver(runtime, operationResult)
  }
  return actualVerification
}
