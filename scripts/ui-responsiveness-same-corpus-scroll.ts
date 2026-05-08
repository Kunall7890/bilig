export interface ScrollSample {
  readonly operationResponseMs: number
  readonly postOperationFrameMs: number
  readonly scrollEventResponseMs: number
  readonly scrollMovementPx: number
}

export interface ScrollTriggerResult {
  readonly scrollEventResponseMs: number
}

export interface ScrollPositionSnapshot {
  readonly scrollLeft: number
  readonly scrollTop: number
}

export interface VisibleScrollResponseMeasurementHooks {
  readonly collectFrameIntervals: (frameCount: number) => Promise<readonly number[]>
  readonly movePointer: () => Promise<void>
  readonly now: () => number
  readonly readScrollPosition: () => Promise<ScrollPositionSnapshot | null>
  readonly scroll: () => Promise<ScrollTriggerResult>
  readonly waitForNextFrame: () => Promise<void>
}

export class ScrollMovementVerificationError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'ScrollMovementVerificationError'
  }
}

const visibleScrollMovementFrameBudget = 60

export async function measureVisibleScrollResponseWithHooks(hooks: VisibleScrollResponseMeasurementHooks): Promise<ScrollSample> {
  await hooks.movePointer()
  const before = await hooks.readScrollPosition()
  const startedAt = hooks.now()
  const scrollResult = await hooks.scroll()
  const { after, scrollMovementPx } = await waitForScrollMovement(hooks, before)
  const operationResponseMs = hooks.now() - startedAt
  if (scrollMovementPx < 1) {
    throw new ScrollMovementVerificationError(
      `Visible scroll response sample did not move the workbook viewport: before=${JSON.stringify(before)}, after=${JSON.stringify(after)}`,
    )
  }
  const frameIntervals = await hooks.collectFrameIntervals(12)
  return {
    operationResponseMs,
    postOperationFrameMs: percentile(frameIntervals, 0.95),
    scrollEventResponseMs: scrollResult.scrollEventResponseMs,
    scrollMovementPx,
  }
}

async function waitForScrollMovement(
  hooks: VisibleScrollResponseMeasurementHooks,
  before: ScrollPositionSnapshot | null,
  frame = 0,
  after: ScrollPositionSnapshot | null = null,
  scrollMovementPx = 0,
): Promise<{ readonly after: ScrollPositionSnapshot | null; readonly scrollMovementPx: number }> {
  if (frame >= visibleScrollMovementFrameBudget) {
    return { after, scrollMovementPx }
  }
  await hooks.waitForNextFrame()
  const nextAfter = await hooks.readScrollPosition()
  const nextScrollMovementPx = resolveScrollMovementPx(before, nextAfter)
  if (nextScrollMovementPx >= 1) {
    return { after: nextAfter, scrollMovementPx: nextScrollMovementPx }
  }
  return waitForScrollMovement(hooks, before, frame + 1, nextAfter, nextScrollMovementPx)
}

function resolveScrollMovementPx(before: ScrollPositionSnapshot | null, after: ScrollPositionSnapshot | null): number {
  if (!before || !after) {
    return 0
  }
  return Math.max(Math.abs(after.scrollLeft - before.scrollLeft), Math.abs(after.scrollTop - before.scrollTop))
}

function percentile(values: readonly number[], percentileValue: number): number {
  if (values.length === 0) {
    throw new Error('Cannot compute percentile for an empty sample set')
  }
  const sorted = [...values].toSorted((left, right) => left - right)
  const index = Math.ceil(percentileValue * sorted.length) - 1
  return sorted[Math.max(0, Math.min(sorted.length - 1, index))]!
}
