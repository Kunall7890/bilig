import type { CellSnapshot } from '@bilig/protocol'
import type { WorkbookDeltaBatchV3 } from '@bilig/worker-transport'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import {
  buildLocalAxisWorkbookDelta,
  buildLocalCellSnapshotsWorkbookDelta,
  buildLocalRangeWorkbookDelta,
  type ProjectedWorkbookLocalDeltaRange,
  type ProjectedWorkbookLocalDeltaSheetIdentity,
} from './projected-workbook-local-delta.js'

type ProjectedWorkbookLocalDeltaAxis = 'column' | 'row'
type ProjectedWorkbookLocalDeltaListener = (batch: WorkbookDeltaBatchV3) => void

interface ProjectedWorkbookLocalDeltaPublisherOptions {
  readonly getLastAuthoritativeRevision: () => number | null
  readonly getLastBatchId: () => number
  readonly now?: () => number
  readonly resolveSheetIdentity: (sheetName: string) => ProjectedWorkbookLocalDeltaSheetIdentity | null
}

export class ProjectedWorkbookLocalDeltaPublisher {
  private readonly listeners = new Set<ProjectedWorkbookLocalDeltaListener>()
  private localWorkbookDeltaSeq = 0

  constructor(private readonly options: ProjectedWorkbookLocalDeltaPublisherOptions) {}

  subscribe(listener: ProjectedWorkbookLocalDeltaListener): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  emitCellSnapshot(snapshot: CellSnapshot): void {
    this.emitCellSnapshots(snapshot.sheetName, [snapshot])
  }

  emitCellSnapshots(sheetName: string, snapshots: readonly CellSnapshot[]): void {
    if (this.listeners.size === 0 || snapshots.length === 0) {
      return
    }
    const identity = this.options.resolveSheetIdentity(sheetName)
    if (!identity) {
      return
    }
    const startedAt = this.now()
    const batch = buildLocalCellSnapshotsWorkbookDelta({
      identity,
      seq: this.nextLocalWorkbookDeltaSeq(),
      snapshots,
    })
    this.publish(batch)
    this.noteRendererDeltaApply(startedAt, snapshots.length)
  }

  emitRange(sheetName: string, range: ProjectedWorkbookLocalDeltaRange): void {
    if (this.listeners.size === 0) {
      return
    }
    const identity = this.options.resolveSheetIdentity(sheetName)
    if (!identity) {
      return
    }
    const startedAt = this.now()
    const batch = buildLocalRangeWorkbookDelta({
      identity,
      range,
      seq: this.nextLocalWorkbookDeltaSeq(),
    })
    this.publish(batch)
    this.noteRendererDeltaApply(startedAt, 1)
  }

  emitAxis(sheetName: string, axis: ProjectedWorkbookLocalDeltaAxis, index: number): void {
    if (this.listeners.size === 0) {
      return
    }
    const identity = this.options.resolveSheetIdentity(sheetName)
    if (!identity) {
      return
    }
    const startedAt = this.now()
    const batch = buildLocalAxisWorkbookDelta({
      axis,
      identity,
      index,
      seq: this.nextLocalWorkbookDeltaSeq(),
    })
    this.publish(batch)
    this.noteRendererDeltaApply(startedAt, 1)
  }

  private publish(batch: WorkbookDeltaBatchV3): void {
    this.listeners.forEach((listener) => {
      listener(batch)
    })
  }

  private nextLocalWorkbookDeltaSeq(): number {
    this.localWorkbookDeltaSeq = Math.max(
      this.localWorkbookDeltaSeq,
      this.options.getLastBatchId(),
      this.options.getLastAuthoritativeRevision() ?? 0,
    )
    return ++this.localWorkbookDeltaSeq
  }

  private noteRendererDeltaApply(startedAt: number, mutationCount: number): void {
    getWorkbookScrollPerfCollector()?.noteRendererDeltaApply({
      dirtyTileCount: Math.max(1, mutationCount),
      durationMs: Math.max(0, this.now() - startedAt),
      mutationCount,
    })
  }

  private now(): number {
    if (this.options.now) {
      return this.options.now()
    }
    return typeof performance === 'undefined' ? Date.now() : performance.now()
  }
}
