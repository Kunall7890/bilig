import type { SpreadsheetEngine } from '@bilig/core'
import type { EngineEvent } from '@bilig/protocol'
import { encodeRenderTileDeltaBatch, type RenderTileDeltaSubscription } from '@bilig/worker-transport'
import { TextOverflowIndexV3 } from '../../../packages/grid/src/renderer-v3/text-overflow-index.js'
import { buildWorkerRenderTileDeltaBatch } from './worker-runtime-render-tile-delta.js'
import type { WorkerEngine } from './worker-runtime-support.js'

export interface WorkerRuntimeRenderTileDiagnostics {
  readonly errorCount: number
  readonly lastError: WorkerRuntimeRenderTileErrorDiagnostic | null
}

export interface WorkerRuntimeRenderTileErrorDiagnostic {
  readonly message: string
  readonly phase: 'initialize' | 'publish'
  readonly sheetId: number
  readonly sheetName: string
}

export class WorkerRuntimeRenderTileDeltaPublisher {
  private nextGeneration = 0
  private errorCount = 0
  private lastError: WorkerRuntimeRenderTileErrorDiagnostic | null = null
  private readonly textOverflowIndex = new TextOverflowIndexV3()

  reset(): void {
    this.nextGeneration = 0
    this.errorCount = 0
    this.lastError = null
    this.textOverflowIndex.clear()
  }

  getDiagnostics(): WorkerRuntimeRenderTileDiagnostics {
    return {
      errorCount: this.errorCount,
      lastError: this.lastError,
    }
  }

  subscribe(input: {
    readonly subscription: RenderTileDeltaSubscription
    readonly listener: (delta: Uint8Array) => void
    readonly getProjectionEngine: () => Promise<SpreadsheetEngine & WorkerEngine>
  }): () => void {
    let disposed = false
    let unsubscribeEngine: (() => void) | null = null
    let publishInFlight = false
    let publishQueued = false
    let queuedEvent: EngineEvent | undefined

    const publish = (engine: SpreadsheetEngine & WorkerEngine, event?: EngineEvent): void => {
      if (publishInFlight) {
        publishQueued = true
        queuedEvent = event
        return
      }
      publishInFlight = true
      try {
        if (disposed) {
          return
        }
        if (!engine.workbook.getSheet(input.subscription.sheetName)) {
          return
        }
        const batch = buildWorkerRenderTileDeltaBatch({
          engine,
          event,
          generation: ++this.nextGeneration,
          subscription: input.subscription,
          textOverflowIndex: this.textOverflowIndex,
        })
        if (event && batch.mutations.length === 0) {
          return
        }
        input.listener(encodeRenderTileDeltaBatch(batch))
      } catch (error) {
        this.recordError('publish', input.subscription, error)
        return
      } finally {
        publishInFlight = false
        if (publishQueued && !disposed) {
          const nextEvent = queuedEvent
          publishQueued = false
          queuedEvent = undefined
          publish(engine, nextEvent)
        }
      }
    }

    void (async () => {
      try {
        const engine = await input.getProjectionEngine()
        if (disposed) {
          return
        }
        if (input.subscription.initialDelta !== 'none') {
          publish(engine)
        }
        unsubscribeEngine = engine.subscribe((event) => {
          publish(engine, event)
        })
      } catch (error) {
        this.recordError('initialize', input.subscription, error)
        return
      }
    })()

    return () => {
      disposed = true
      unsubscribeEngine?.()
      unsubscribeEngine = null
    }
  }

  private recordError(
    phase: WorkerRuntimeRenderTileErrorDiagnostic['phase'],
    subscription: RenderTileDeltaSubscription,
    error: unknown,
  ): void {
    this.errorCount += 1
    this.lastError = {
      message: error instanceof Error ? error.message : String(error),
      phase,
      sheetId: subscription.sheetId,
      sheetName: subscription.sheetName,
    }
  }
}
