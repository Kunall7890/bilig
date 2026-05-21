import { formatAddress } from '@bilig/formula'
import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { VisibleRegionState } from '../gridPointer.js'
import type { Item, Rectangle } from '../gridTypes.js'
import { collectViewportItems } from '../gridViewportItems.js'
import { sameViewportBounds } from '../gridViewportController.js'
import { viewportFromVisibleRegion } from '../useGridCameraState.js'
import { resolveResidentViewport } from '../workbookGridViewport.js'

export interface GridResidentHeaderRegion {
  readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  readonly tx: number
  readonly ty: number
  readonly freezeRows: number
  readonly freezeCols: number
}

export interface GridViewportResidencyState {
  readonly viewport: Viewport
  readonly residentAddresses: readonly string[]
  readonly residentViewport: Viewport
  readonly renderTileViewport: Viewport
  readonly residentHeaderItems: readonly Item[]
  readonly residentItems: readonly Item[]
  readonly residentHeaderRegion: GridResidentHeaderRegion
  readonly sceneRevision: number
  readonly visibleAddresses: readonly string[]
  readonly visibleItems: readonly Item[]
}

export interface GridViewportResidencyRuntimeInput {
  readonly freezeCols: number
  readonly freezeRows: number
  readonly visibleRegion: VisibleRegionState
}

export interface GridViewportResidencyInvalidationInput {
  readonly engine: GridEngineLike
  readonly residentAddresses: readonly string[]
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
}

interface RuntimeConnection<Identity> {
  readonly identity: Identity
  readonly unsubscribe: (() => void) | undefined
}

interface LocalSceneInvalidationConnectionIdentity {
  readonly engine: GridEngineLike
  readonly residentAddresses: readonly string[]
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
}

interface GridViewportResidentCache {
  readonly freezeCols: number
  readonly freezeRows: number
  readonly residentAddresses: readonly string[]
  readonly renderTileViewport: Viewport
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: GridResidentHeaderRegion
  readonly residentItems: readonly Item[]
  readonly residentViewport: Viewport
}

export class GridViewportResidencyRuntime {
  private residentCache: GridViewportResidentCache | null = null
  private residentViewport: Viewport | null = null
  private localSceneInvalidationConnection: RuntimeConnection<LocalSceneInvalidationConnectionIdentity> | null = null
  private readonly sceneRevisionListeners = new Set<() => void>()
  private sceneRevision = 0

  resolve(input: GridViewportResidencyRuntimeInput): GridViewportResidencyState {
    const viewport = viewportFromVisibleRegion(input.visibleRegion)
    const nextResidentViewport = resolveResidentViewport(viewport)
    if (!this.residentViewport || !sameViewportBounds(this.residentViewport, nextResidentViewport)) {
      this.residentViewport = nextResidentViewport
    }
    const residentViewport = this.residentViewport
    const residentCache = this.resolveResidentCache(input, residentViewport)
    const visibleItems = collectViewportItems(viewport, {
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
    })

    return {
      residentAddresses: residentCache.residentAddresses,
      residentItems: residentCache.residentItems,
      renderTileViewport: residentCache.renderTileViewport,
      residentHeaderItems: residentCache.residentHeaderItems,
      residentHeaderRegion: residentCache.residentHeaderRegion,
      residentViewport,
      sceneRevision: this.sceneRevision,
      viewport,
      visibleAddresses: visibleItems.map(([col, row]) => formatAddress(row, col)),
      visibleItems,
    }
  }

  invalidateScene(): number {
    this.sceneRevision += 1
    this.emitSceneRevision()
    return this.sceneRevision
  }

  snapshotSceneRevision(): number {
    return this.sceneRevision
  }

  subscribeSceneRevision(listener: () => void): () => void {
    this.sceneRevisionListeners.add(listener)
    return () => {
      this.sceneRevisionListeners.delete(listener)
    }
  }

  connectLocalSceneInvalidation(input: GridViewportResidencyInvalidationInput, listener?: () => void): (() => void) | undefined {
    if (input.shouldUseRemoteRenderTileSource || input.residentAddresses.length === 0) {
      return undefined
    }
    const invalidate = () => {
      this.invalidateScene()
      listener?.()
    }
    const unsubscribeCells = input.engine.subscribeCells(input.sheetName, input.residentAddresses, invalidate)
    const unsubscribeMerges = input.engine.subscribeSheetChannel?.(input.sheetName, 'merges', invalidate)
    if (!unsubscribeMerges) {
      return unsubscribeCells
    }
    return () => {
      unsubscribeCells()
      unsubscribeMerges()
    }
  }

  syncLocalSceneInvalidation(input: GridViewportResidencyInvalidationInput): void {
    const identity: LocalSceneInvalidationConnectionIdentity = {
      engine: input.engine,
      residentAddresses: input.residentAddresses,
      sheetName: input.sheetName,
      shouldUseRemoteRenderTileSource: input.shouldUseRemoteRenderTileSource,
    }
    if (
      this.localSceneInvalidationConnection &&
      sameLocalSceneInvalidationConnectionIdentity(this.localSceneInvalidationConnection.identity, identity)
    ) {
      return
    }
    this.localSceneInvalidationConnection?.unsubscribe?.()
    this.localSceneInvalidationConnection = {
      identity,
      unsubscribe: this.connectLocalSceneInvalidation(input),
    }
  }

  disconnectLocalSceneInvalidation(): void {
    this.localSceneInvalidationConnection?.unsubscribe?.()
    this.localSceneInvalidationConnection = null
  }

  private resolveResidentCache(input: GridViewportResidencyRuntimeInput, residentViewport: Viewport): GridViewportResidentCache {
    const current = this.residentCache
    if (
      current?.residentViewport === residentViewport &&
      current.freezeCols === input.freezeCols &&
      current.freezeRows === input.freezeRows
    ) {
      return current
    }

    const residentItems = collectViewportItems(residentViewport, {
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
    })
    const next: GridViewportResidentCache = {
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      residentAddresses: residentItems.map(([col, row]) => formatAddress(row, col)),
      renderTileViewport: {
        rowStart: input.freezeRows > 0 ? 0 : residentViewport.rowStart,
        rowEnd: residentViewport.rowEnd,
        colStart: input.freezeCols > 0 ? 0 : residentViewport.colStart,
        colEnd: residentViewport.colEnd,
      },
      residentHeaderItems: residentItems,
      residentItems,
      residentHeaderRegion: {
        range: {
          x: residentViewport.colStart,
          y: residentViewport.rowStart,
          width: residentViewport.colEnd - residentViewport.colStart + 1,
          height: residentViewport.rowEnd - residentViewport.rowStart + 1,
        },
        tx: 0,
        ty: 0,
        freezeRows: input.freezeRows,
        freezeCols: input.freezeCols,
      },
      residentViewport,
    }
    this.residentCache = next
    return next
  }

  private emitSceneRevision(): void {
    this.sceneRevisionListeners.forEach((listener) => {
      listener()
    })
  }
}

function sameLocalSceneInvalidationConnectionIdentity(
  left: LocalSceneInvalidationConnectionIdentity,
  right: LocalSceneInvalidationConnectionIdentity,
): boolean {
  return (
    left.engine === right.engine &&
    left.sheetName === right.sheetName &&
    left.shouldUseRemoteRenderTileSource === right.shouldUseRemoteRenderTileSource &&
    sameStringListIdentity(left.residentAddresses, right.residentAddresses)
  )
}

function sameStringListIdentity(left: readonly string[], right: readonly string[]): boolean {
  if (left === right) {
    return true
  }
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}
