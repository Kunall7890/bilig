import type { Viewport } from '@bilig/protocol'
import { noteRendererTileReadiness } from '../grid-render-counters.js'
import type { GridEngineLike } from '../grid-engine.js'
import type { GridMetrics } from '../gridMetrics.js'
import { buildLocalFixedRenderTiles } from '../renderer-v3/local-render-tile-materializer.js'
import { buildFixedRenderTilePaneStates } from '../renderer-v3/render-tile-pane-builder.js'
import type { GridRenderTile, GridRenderTileSceneChange, GridRenderTileSource } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import type { WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
import { MAX_TILE_COLUMN_INDEX, MAX_TILE_ROW_INDEX, packTileKey53, unpackTileKey53, type TileKey53 } from '../renderer-v3/tile-key.js'
import type { GridTileInterestBatchV3, GridTileReadinessSnapshotV3 } from './gridTileCoordinator.js'
import type { GridRuntimeHost } from './gridRuntimeHost.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export interface GridRenderTilePaneRuntimeState {
  readonly needsLocalCellInvalidation: boolean
  readonly preloadDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly renderTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly tileReadiness: GridTileReadinessSnapshotV3
}

export interface GridRenderTilePaneBridgeState {
  readonly forceLocalTiles: boolean
  readonly localFallbackRevision: number
  readonly renderTileRevision: number
}

export interface GridRenderTilePaneRuntimeInput {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly dprBucket: number
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly forceLocalTiles?: boolean | undefined
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostReady: boolean
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sceneRevision: number
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly visibleViewport: Viewport
}

export interface GridRenderTileDeltaRuntimeInput {
  readonly dprBucket: number
  readonly freezeCols?: number | undefined
  readonly freezeRows?: number | undefined
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly residentViewport?: Viewport | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
}

export interface GridRenderTileDamageRuntimeInput {
  readonly dprBucket: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}

export interface GridRenderTileLocalInvalidationRuntimeInput {
  readonly engine: GridEngineLike
  readonly needsLocalCellInvalidation: boolean
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

export interface GridRenderTileConnectionRuntimeInput {
  readonly dprBucket: number
  readonly engine: GridEngineLike
  readonly freezeCols?: number | undefined
  readonly freezeRows?: number | undefined
  readonly gridRuntimeHost: GridRuntimeHost
  readonly needsLocalCellInvalidation: boolean
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly residentViewport?: Viewport | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

interface GridRenderTileInterestRuntimeInput {
  readonly dprBucket: number
  readonly freezeCols?: number | undefined
  readonly freezeRows?: number | undefined
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileViewport: Viewport
  readonly residentViewport?: Viewport | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}

const EMPTY_TILE_PANE_RUNTIME_STATE: GridRenderTilePaneRuntimeState = Object.freeze({
  needsLocalCellInvalidation: false,
  preloadDataPanes: [],
  renderTilePanes: [],
  residentBodyPane: null,
  residentDataPanes: [],
  tileReadiness: {
    exactHits: [],
    misses: [],
    staleHits: [],
    visibleDirtyTileKeys: [],
    warmDirtyTileKeys: [],
  },
})

const INITIAL_RENDER_TILE_PANE_BRIDGE_STATE: GridRenderTilePaneBridgeState = Object.freeze({
  forceLocalTiles: false,
  localFallbackRevision: 0,
  renderTileRevision: 0,
})

type TileResolutionSource = 'local' | 'remote'

interface FixedRenderTileDataPanesResolution {
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly source: TileResolutionSource
}

interface GridRenderTileResolution {
  readonly tiles: readonly GridRenderTile[]
  readonly source: TileResolutionSource
}

interface GridRenderTilePreloadResolution {
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly tiles: readonly GridRenderTile[]
}

interface RuntimeConnection<Identity> {
  readonly identity: Identity
  readonly unsubscribe: (() => void) | undefined
}

interface RenderTileSheetIdentity {
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}

interface RenderTileDeltaConnectionIdentity {
  readonly dprBucket: number
  readonly freezeCols: number
  readonly freezeRows: number
  readonly renderTileSource: GridRenderTileSource | undefined
  readonly residentViewport: Viewport | undefined
  readonly sheetId: number | undefined
  readonly sheetOrdinal: number | undefined
  readonly sheetName: string
  readonly viewport: Viewport
}

interface WorkbookDeltaConnectionIdentity {
  readonly dprBucket: number
  readonly renderTileSource: GridRenderTileSource | undefined
  readonly sheetId: number | undefined
  readonly sheetOrdinal: number | undefined
}

interface LocalInvalidationConnectionIdentity {
  readonly engine: GridEngineLike
  readonly needsLocalCellInvalidation: boolean
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

export class GridRenderTilePaneRuntime {
  private retainedFixedRenderTileDataPanes: {
    readonly compatibility: RetainedFixedRenderTileDataPanesCompatibility
    readonly panes: readonly WorkbookRenderTilePaneState[]
  } | null = null
  private bridgeState = INITIAL_RENDER_TILE_PANE_BRIDGE_STATE
  private readonly bridgeListeners = new Set<() => void>()
  private readonly lastWorkbookDeltaSeqBySheetAndSource = new Map<string, number>()
  private renderTileDeltaConnection: RuntimeConnection<RenderTileDeltaConnectionIdentity> | null = null
  private workbookDeltaConnection: RuntimeConnection<WorkbookDeltaConnectionIdentity> | null = null
  private localInvalidationConnection: RuntimeConnection<LocalInvalidationConnectionIdentity> | null = null

  resolve(input: GridRenderTilePaneRuntimeInput): GridRenderTilePaneRuntimeState {
    if (!input.hostReady) {
      return EMPTY_TILE_PANE_RUNTIME_STATE
    }
    const resolution = this.resolveTiles(input)
    const preloadResolution = this.resolvePreloadPanes(input, resolution?.tiles ?? [])
    const tileReadiness = this.resolveTileReadiness(input, resolution?.tiles ?? [], preloadResolution.tiles)
    const fixedRenderTileDataPanes = resolution ? this.buildFixedRenderTileDataPanes(input, resolution) : null
    if (input.sheetId !== undefined && fixedRenderTileDataPanes?.source === 'remote') {
      this.retainedFixedRenderTileDataPanes = {
        compatibility: buildRetainedFixedRenderTileDataPanesCompatibility(input),
        panes: fixedRenderTileDataPanes.panes,
      }
    }

    const shouldUseRemoteRenderTileSource = input.renderTileSource !== undefined && input.sheetId !== undefined
    const retainedFixedRenderTileDataPanes =
      fixedRenderTileDataPanes?.panes ??
      (shouldUseRemoteRenderTileSource &&
      this.retainedFixedRenderTileDataPanes &&
      sameRetainedFixedRenderTileDataPanesCompatibility(
        this.retainedFixedRenderTileDataPanes.compatibility,
        buildRetainedFixedRenderTileDataPanesCompatibility(input),
      )
        ? this.retainedFixedRenderTileDataPanes.panes
        : null)
    const residentDataPanes = retainedFixedRenderTileDataPanes ?? []
    return {
      needsLocalCellInvalidation: !shouldUseRemoteRenderTileSource,
      preloadDataPanes: preloadResolution.panes,
      renderTilePanes: residentDataPanes,
      residentBodyPane: residentDataPanes.find((pane) => pane.paneId === 'body') ?? null,
      residentDataPanes,
      tileReadiness,
    }
  }

  clearRetainedPanes(): void {
    this.retainedFixedRenderTileDataPanes = null
  }

  snapshotBridgeState(): GridRenderTilePaneBridgeState {
    return this.bridgeState
  }

  subscribeBridgeState(listener: () => void): () => void {
    this.bridgeListeners.add(listener)
    return () => {
      this.bridgeListeners.delete(listener)
    }
  }

  noteRenderTileDelta(): GridRenderTilePaneBridgeState {
    const previous = this.bridgeState
    this.bridgeState = {
      forceLocalTiles: false,
      localFallbackRevision: previous.localFallbackRevision,
      renderTileRevision: previous.renderTileRevision + 1,
    }
    this.emitBridgeState()
    return this.bridgeState
  }

  noteWorkbookDeltaDamage(input: { readonly forceLocalTiles?: boolean | undefined } = {}): GridRenderTilePaneBridgeState {
    const previous = this.bridgeState
    this.bridgeState = {
      forceLocalTiles: input.forceLocalTiles ?? true,
      localFallbackRevision: previous.localFallbackRevision + 1,
      renderTileRevision: previous.renderTileRevision + 1,
    }
    this.emitBridgeState()
    return this.bridgeState
  }

  noteLocalFallbackInvalidation(): GridRenderTilePaneBridgeState {
    const previous = this.bridgeState
    this.bridgeState = {
      forceLocalTiles: true,
      localFallbackRevision: previous.localFallbackRevision + 1,
      renderTileRevision: previous.renderTileRevision,
    }
    this.emitBridgeState()
    return this.bridgeState
  }

  noteTileReadiness(readiness: GridTileReadinessSnapshotV3): void {
    const exactHits = readiness.exactHits.length
    const staleHits = readiness.staleHits.length
    const misses = readiness.misses.length
    const visibleDirtyTiles = readiness.visibleDirtyTileKeys.length
    const warmDirtyTiles = readiness.warmDirtyTileKeys.length
    if (exactHits + staleHits + misses + visibleDirtyTiles + warmDirtyTiles === 0) {
      return
    }
    noteRendererTileReadiness({
      exactHits,
      misses,
      staleHits,
      visibleDirtyTiles,
      warmDirtyTiles,
    })
  }

  connectLocalCellInvalidation(input: GridRenderTileLocalInvalidationRuntimeInput, listener?: () => void): (() => void) | undefined {
    if (!input.needsLocalCellInvalidation || input.visibleAddresses.length === 0) {
      return undefined
    }
    const invalidate = () => {
      this.clearRetainedPanes()
      this.noteLocalFallbackInvalidation()
      listener?.()
    }
    const unsubscribeCells = input.engine.subscribeCells(input.sheetName, input.visibleAddresses, invalidate)
    const unsubscribeMerges = input.engine.subscribeSheetChannel?.(input.sheetName, 'merges', invalidate)
    if (!unsubscribeMerges) {
      return unsubscribeCells
    }
    return () => {
      unsubscribeCells()
      unsubscribeMerges()
    }
  }

  syncConnections(input: GridRenderTileConnectionRuntimeInput): void {
    this.syncRenderTileDeltaConnection(input)
    this.syncWorkbookDeltaConnection(input)
    this.syncLocalInvalidationConnection(input)
  }

  disconnectConnections(): void {
    this.renderTileDeltaConnection?.unsubscribe?.()
    this.renderTileDeltaConnection = null
    this.workbookDeltaConnection?.unsubscribe?.()
    this.workbookDeltaConnection = null
    this.localInvalidationConnection?.unsubscribe?.()
    this.localInvalidationConnection = null
  }

  connectWorkbookDeltaDamage(
    input: GridRenderTileDamageRuntimeInput,
    listener?: (batch: WorkbookDeltaBatchLikeV3) => void,
  ): (() => void) | undefined {
    const renderTileSource = input.renderTileSource
    if (!renderTileSource?.subscribeWorkbookDeltas || input.sheetId === undefined) {
      return undefined
    }
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    return renderTileSource.subscribeWorkbookDeltas((batch) => {
      if (!matchesRenderTileSheetIdentity(batch, { sheetId: input.sheetId, sheetOrdinal })) {
        return
      }
      if (!this.applyWorkbookDeltaDamage(input, batch)) {
        return
      }
      this.noteWorkbookDeltaDamage({
        forceLocalTiles: shouldForceLocalTilesForWorkbookDelta(batch),
      })
      listener?.(batch)
    })
  }

  connectRenderTileDeltas(
    input: GridRenderTileDeltaRuntimeInput,
    listener?: (change: GridRenderTileSceneChange) => void,
  ): (() => void) | undefined {
    if (!input.renderTileSource || input.sheetId === undefined) {
      return undefined
    }
    const tileInterest = this.buildViewportTileInterest({
      ...input,
      sheetId: input.sheetId,
    })
    return input.renderTileSource.subscribeRenderTileDeltas(
      {
        ...input.renderTileViewport,
        cameraSeq: tileInterest.cameraSeq,
        dprBucket: input.dprBucket,
        initialDelta: 'full',
        sheetId: input.sheetId,
        sheetOrdinal: tileInterest.sheetOrdinal,
        sheetName: input.sheetName,
        tileInterest: {
          axisSeqX: tileInterest.axisSeqX,
          axisSeqY: tileInterest.axisSeqY,
          freezeSeq: tileInterest.freezeSeq,
          pinnedTileKeys: tileInterest.pinnedTileKeys,
          reason: tileInterest.reason,
          seq: tileInterest.seq,
          sheetOrdinal: tileInterest.sheetOrdinal,
          visibleTileKeys: tileInterest.visibleTileKeys,
          warmTileKeys: tileInterest.warmTileKeys,
        },
        warmTileKeys: tileInterest.warmTileKeys,
      },
      (change) => {
        if (change) {
          this.applyRenderTileSceneChange(input, change)
        }
        this.noteRenderTileDelta()
        listener?.(change)
      },
    )
  }

  private applyWorkbookDeltaDamage(input: GridRenderTileDamageRuntimeInput, batch: WorkbookDeltaBatchLikeV3): boolean {
    if (batch.seq !== undefined) {
      const source = 'source' in batch && typeof batch.source === 'string' ? batch.source : 'unknown'
      const sequenceKey = `${batch.sheetId ?? 'x'}:${batch.sheetOrdinal ?? 'x'}:${source}`
      const lastSeq = this.lastWorkbookDeltaSeqBySheetAndSource.get(sequenceKey) ?? -1
      if (batch.seq <= lastSeq) {
        return false
      }
      this.lastWorkbookDeltaSeqBySheetAndSource.set(sequenceKey, batch.seq)
    }
    input.gridRuntimeHost.tiles.applyWorkbookDelta(batch, { dprBucket: input.dprBucket })
    return true
  }

  private syncRenderTileDeltaConnection(input: GridRenderTileConnectionRuntimeInput): void {
    const identity: RenderTileDeltaConnectionIdentity = {
      dprBucket: input.dprBucket,
      freezeCols: input.freezeCols ?? 0,
      freezeRows: input.freezeRows ?? 0,
      renderTileSource: input.renderTileSource,
      residentViewport: input.residentViewport,
      sheetId: input.sheetId,
      sheetName: input.sheetName,
      sheetOrdinal: input.sheetOrdinal,
      viewport: input.renderTileViewport,
    }
    if (this.renderTileDeltaConnection && sameRenderTileDeltaConnectionIdentity(this.renderTileDeltaConnection.identity, identity)) {
      return
    }
    this.renderTileDeltaConnection?.unsubscribe?.()
    this.renderTileDeltaConnection = {
      identity,
      unsubscribe: this.connectRenderTileDeltas({
        dprBucket: input.dprBucket,
        gridRuntimeHost: input.gridRuntimeHost,
        freezeCols: input.freezeCols,
        freezeRows: input.freezeRows,
        renderTileSource: input.renderTileSource,
        renderTileViewport: input.renderTileViewport,
        residentViewport: input.residentViewport,
        sheetId: input.sheetId,
        sheetName: input.sheetName,
        sheetOrdinal: input.sheetOrdinal,
      }),
    }
  }

  private syncWorkbookDeltaConnection(input: GridRenderTileConnectionRuntimeInput): void {
    const identity: WorkbookDeltaConnectionIdentity = {
      dprBucket: input.dprBucket,
      renderTileSource: input.renderTileSource,
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetOrdinal,
    }
    if (this.workbookDeltaConnection && sameWorkbookDeltaConnectionIdentity(this.workbookDeltaConnection.identity, identity)) {
      return
    }
    this.workbookDeltaConnection?.unsubscribe?.()
    this.workbookDeltaConnection = {
      identity,
      unsubscribe: this.connectWorkbookDeltaDamage({
        dprBucket: input.dprBucket,
        gridRuntimeHost: input.gridRuntimeHost,
        renderTileSource: input.renderTileSource,
        sheetId: input.sheetId,
        sheetOrdinal: input.sheetOrdinal,
      }),
    }
  }

  private syncLocalInvalidationConnection(input: GridRenderTileConnectionRuntimeInput): void {
    const identity: LocalInvalidationConnectionIdentity = {
      engine: input.engine,
      needsLocalCellInvalidation: input.needsLocalCellInvalidation,
      sheetName: input.sheetName,
      visibleAddresses: input.visibleAddresses,
    }
    if (this.localInvalidationConnection && sameLocalInvalidationConnectionIdentity(this.localInvalidationConnection.identity, identity)) {
      return
    }
    this.localInvalidationConnection?.unsubscribe?.()
    this.localInvalidationConnection = {
      identity,
      unsubscribe: this.connectLocalCellInvalidation({
        engine: input.engine,
        needsLocalCellInvalidation: input.needsLocalCellInvalidation,
        sheetName: input.sheetName,
        visibleAddresses: input.visibleAddresses,
      }),
    }
  }

  private buildFixedRenderTileDataPanes(
    input: GridRenderTilePaneRuntimeInput,
    resolution: GridRenderTileResolution,
  ): FixedRenderTileDataPanesResolution | null {
    const panes = buildFixedRenderTilePaneStates({
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      frozenColumnWidth: input.frozenColumnWidth,
      frozenRowHeight: input.frozenRowHeight,
      gridMetrics: input.gridMetrics,
      hostHeight: input.hostClientHeight,
      hostWidth: input.hostClientWidth,
      residentViewport: input.residentViewport,
      sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
      sortedRowHeightOverrides: input.sortedRowHeightOverrides,
      tiles: resolution.tiles,
      visibleViewport: input.visibleViewport,
    })
    return panes.length > 0 ? { panes, source: resolution.source } : null
  }

  private resolveTileReadiness(
    input: GridRenderTilePaneRuntimeInput,
    tiles: readonly GridRenderTile[],
    preloadTiles: readonly GridRenderTile[] = [],
  ): GridTileReadinessSnapshotV3 {
    if (input.sheetId === undefined) {
      return EMPTY_TILE_PANE_RUNTIME_STATE.tileReadiness
    }
    const sheetId = input.sheetId
    for (const tile of tiles) {
      this.upsertHostTile(input, tile)
    }
    for (const tile of preloadTiles) {
      this.upsertHostTile(input, tile)
    }
    const interest = this.buildViewportTileInterest({
      ...input,
      sheetId,
    })
    return input.gridRuntimeHost.tiles.reconcileInterest(interest)
  }

  private buildViewportTileInterest(input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number }): GridTileInterestBatchV3 {
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    const snapshot = input.gridRuntimeHost.snapshot()
    const visibleTileKeys = resolveRenderTileInterestTileKeys(input)
    return input.gridRuntimeHost.tiles.buildInterest({
      axisSeqX: snapshot.axisSeqX,
      axisSeqY: snapshot.axisSeqY,
      cameraSeq: snapshot.camera.seq,
      freezeSeq: snapshot.freezeSeq,
      pinnedTileKeys: [],
      reason: 'scroll',
      sheetId: input.sheetId,
      sheetOrdinal,
      visibleTileKeys,
      warmTileKeys: this.resolveWarmTileKeys(input, visibleTileKeys),
    })
  }

  private resolveWarmTileKeys(
    input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number },
    visibleTileKeys = resolveRenderTileInterestTileKeys(input),
  ): readonly TileKey53[] {
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    if (visibleTileKeys.length === 0) {
      return []
    }
    const visibleSet = new Set(visibleTileKeys)
    const warmSet = new Set<number>()
    const warmTileKeys: number[] = []
    for (const viewport of resolveRenderTileInterestViewports(input)) {
      let minRowTile = Number.POSITIVE_INFINITY
      let maxRowTile = Number.NEGATIVE_INFINITY
      let minColTile = Number.POSITIVE_INFINITY
      let maxColTile = Number.NEGATIVE_INFINITY
      for (const key of input.gridRuntimeHost.viewportTileKeys({ dprBucket: input.dprBucket, sheetOrdinal, viewport })) {
        const fields = unpackTileKey53(key)
        minRowTile = Math.min(minRowTile, fields.rowTile)
        maxRowTile = Math.max(maxRowTile, fields.rowTile)
        minColTile = Math.min(minColTile, fields.colTile)
        maxColTile = Math.max(maxColTile, fields.colTile)
      }
      if (!Number.isFinite(minRowTile)) {
        continue
      }
      for (let rowTile = Math.max(0, minRowTile - 1); rowTile <= Math.min(MAX_TILE_ROW_INDEX, maxRowTile + 1); rowTile += 1) {
        for (let colTile = Math.max(0, minColTile - 1); colTile <= Math.min(MAX_TILE_COLUMN_INDEX, maxColTile + 1); colTile += 1) {
          const key = packTileKey53({
            colTile,
            dprBucket: input.dprBucket,
            rowTile,
            sheetOrdinal,
          })
          if (!visibleSet.has(key) && !warmSet.has(key)) {
            warmSet.add(key)
            warmTileKeys.push(key)
          }
        }
      }
    }
    return warmTileKeys
  }

  private resolvePreloadPanes(
    input: GridRenderTilePaneRuntimeInput,
    visibleTiles: readonly GridRenderTile[],
  ): GridRenderTilePreloadResolution {
    if (!input.renderTileSource || input.sheetId === undefined) {
      return { panes: EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes, tiles: [] }
    }
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    const visibleTileIds = new Set(visibleTiles.map((tile) => tile.tileId))
    const tiles: GridRenderTile[] = []
    for (const tileKey of this.resolveWarmTileKeys({
      dprBucket: input.dprBucket,
      gridRuntimeHost: input.gridRuntimeHost,
      renderTileViewport: input.renderTileViewport,
      sheetId: input.sheetId,
      sheetOrdinal,
    })) {
      if (visibleTileIds.has(tileKey)) {
        continue
      }
      const tile = input.renderTileSource.peekRenderTile(tileKey)
      if (!tile || !matchesRenderTileSheetIdentity(tile.coord, { sheetId: input.sheetId, sheetOrdinal })) {
        continue
      }
      tiles.push(tile)
    }
    if (tiles.length === 0) {
      return { panes: EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes, tiles }
    }
    return {
      panes: this.buildPreloadPanes(input, tiles),
      tiles,
    }
  }

  private buildPreloadPanes(
    input: GridRenderTilePaneRuntimeInput,
    tiles: readonly GridRenderTile[],
  ): readonly WorkbookRenderTilePaneState[] {
    const preloadViewport = resolveRenderTileViewportUnion(tiles)
    if (!preloadViewport) {
      return EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes
    }
    return buildFixedRenderTilePaneStates({
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      frozenColumnWidth: input.frozenColumnWidth,
      frozenRowHeight: input.frozenRowHeight,
      gridMetrics: input.gridMetrics,
      hostHeight: input.hostClientHeight,
      hostWidth: input.hostClientWidth,
      residentViewport: preloadViewport,
      sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
      sortedRowHeightOverrides: input.sortedRowHeightOverrides,
      tiles,
      visibleViewport: input.visibleViewport,
    })
  }

  private upsertHostTile(input: GridRenderTilePaneRuntimeInput, tile: GridRenderTile): void {
    upsertRenderTileIntoHost(input.gridRuntimeHost, tile)
  }

  private applyRenderTileSceneChange(input: GridRenderTileDeltaRuntimeInput, change: GridRenderTileSceneChange): void {
    change.invalidatedTileIds.forEach((tileId) => {
      input.gridRuntimeHost.tiles.deleteTile(tileId)
    })
    const renderTileSource = input.renderTileSource
    if (!renderTileSource) {
      return
    }
    change.changedTileIds.forEach((tileId) => {
      const tile = renderTileSource.peekRenderTile(tileId)
      if (
        !tile ||
        !matchesRenderTileSheetIdentity(tile.coord, {
          sheetId: input.sheetId,
          sheetOrdinal: resolveGridRenderTileInputSheetOrdinal(input),
        })
      ) {
        return
      }
      upsertRenderTileIntoHost(input.gridRuntimeHost, tile)
    })
  }

  private resolveTiles(input: GridRenderTilePaneRuntimeInput): GridRenderTileResolution | null {
    if (input.forceLocalTiles) {
      return this.buildLocalTiles(input, { mergeCleanRemoteTiles: true })
    }

    if (input.renderTileSource && input.sheetId !== undefined) {
      const tiles: GridRenderTile[] = []
      const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
      const tileKeys = input.gridRuntimeHost.viewportTileKeys({
        dprBucket: input.dprBucket,
        sheetOrdinal,
        viewport: input.renderTileViewport,
      })
      for (const tileKey of tileKeys) {
        const tile = input.renderTileSource.peekRenderTile(tileKey)
        if (!tile || !matchesRenderTileSheetIdentity(tile.coord, { sheetId: input.sheetId, sheetOrdinal })) {
          continue
        }
        tiles.push(tile)
      }
      return { source: 'remote', tiles }
    }

    return this.buildLocalTiles(input)
  }

  private buildLocalTiles(
    input: GridRenderTilePaneRuntimeInput,
    options: { readonly mergeCleanRemoteTiles?: boolean } = {},
  ): GridRenderTileResolution {
    if (options.mergeCleanRemoteTiles && input.renderTileSource && input.sheetId !== undefined) {
      const hybrid = this.buildHybridLocalDirtyTiles(input, input.renderTileSource, input.sheetId)
      if (hybrid) {
        return hybrid
      }
    }
    return {
      source: 'local',
      tiles: buildLocalFixedRenderTiles({
        cameraSeq: input.gridRuntimeHost.snapshot().camera.seq,
        columnWidths: input.columnWidths,
        dirtySpansForTile: (tileId) => input.gridRuntimeHost.tiles.dirtyTiles.getSpans(tileId),
        dprBucket: input.dprBucket,
        engine: input.engine,
        freezeSeq: input.gridRuntimeHost.snapshot().freezeSeq,
        generation: input.sceneRevision,
        gridMetrics: input.gridMetrics,
        rowHeights: input.rowHeights,
        sheetId: input.sheetId ?? 0,
        sheetOrdinal: resolveGridRenderTileInputSheetOrdinal(input),
        sheetName: input.sheetName,
        sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
        sortedRowHeightOverrides: input.sortedRowHeightOverrides,
        viewport: input.renderTileViewport,
      }),
    }
  }

  private buildHybridLocalDirtyTiles(
    input: GridRenderTilePaneRuntimeInput,
    renderTileSource: GridRenderTileSource,
    sheetId: number,
  ): GridRenderTileResolution | null {
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    const tileKeys = input.gridRuntimeHost.viewportTileKeys({
      dprBucket: input.dprBucket,
      sheetOrdinal,
      viewport: input.renderTileViewport,
    })
    if (tileKeys.length === 0) {
      return { source: 'local', tiles: [] }
    }

    const remoteTiles = new Map<number, GridRenderTile>()
    const dirtyBaseTiles = new Map<number, GridRenderTile>()
    const dirtyTileKeys: number[] = []
    for (const tileKey of tileKeys) {
      const sourceTile = renderTileSource.peekRenderTile(tileKey)
      const tile =
        sourceTile && sourceTile.coord.sheetId === sheetId && sourceTile.coord.sheetOrdinal === sheetOrdinal
          ? sourceTile
          : this.resolveResidentRenderTile(input, tileKey, sheetId, sheetOrdinal)
      const isDirty = input.gridRuntimeHost.tiles.dirtyTiles.getUnconsumedMask(tileKey) !== 0
      if (isDirty || !tile) {
        if (isDirty && tile) {
          dirtyBaseTiles.set(tileKey, tile)
        }
        dirtyTileKeys.push(tileKey)
        continue
      }
      remoteTiles.set(tileKey, tile)
    }

    if (dirtyTileKeys.length === 0 && remoteTiles.size === tileKeys.length) {
      return { source: 'remote', tiles: tileKeys.map((tileKey) => remoteTiles.get(tileKey)!) }
    }

    const localTiles = new Map(
      buildLocalFixedRenderTiles({
        cameraSeq: input.gridRuntimeHost.snapshot().camera.seq,
        columnWidths: input.columnWidths,
        dirtySpansForTile: (tileId) => input.gridRuntimeHost.tiles.dirtyTiles.getSpans(tileId),
        dprBucket: input.dprBucket,
        engine: input.engine,
        generation: input.sceneRevision,
        gridMetrics: input.gridMetrics,
        rowHeights: input.rowHeights,
        sheetId,
        sheetOrdinal,
        sheetName: input.sheetName,
        sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
        sortedRowHeightOverrides: input.sortedRowHeightOverrides,
        tileKeys: dirtyTileKeys,
        reuseStaticGridRectsByTileId: dirtyBaseTiles,
        viewport: input.renderTileViewport,
      }).map((tile) => [tile.tileId, tile] as const),
    )

    return {
      source: 'local',
      tiles: tileKeys.flatMap((tileKey) => {
        const localTile = localTiles.get(tileKey)
        if (localTile) {
          return [localTile]
        }
        const remoteTile = remoteTiles.get(tileKey)
        return remoteTile ? [remoteTile] : []
      }),
    }
  }

  private resolveResidentRenderTile(
    input: GridRenderTilePaneRuntimeInput,
    tileKey: number,
    sheetId: number,
    sheetOrdinal: number,
  ): GridRenderTile | null {
    const packet = input.gridRuntimeHost.tiles.residency.getExact(tileKey)?.packet
    if (!isResidentGridRenderTile(packet)) {
      return null
    }
    return packet.coord.sheetId === sheetId && packet.coord.sheetOrdinal === sheetOrdinal ? packet : null
  }

  private emitBridgeState(): void {
    this.bridgeListeners.forEach((listener) => {
      listener()
    })
  }
}

interface RetainedFixedRenderTileDataPanesCompatibility {
  readonly dprBucket: number
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly renderTileViewport: Viewport
  readonly residentViewport: Viewport
  readonly sceneRevision: number
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly visibleViewport: Viewport
}

export function getGridRenderTilePaneRuntime(current: unknown): GridRenderTilePaneRuntime {
  return current instanceof GridRenderTilePaneRuntime ? current : new GridRenderTilePaneRuntime()
}

function isResidentGridRenderTile(value: unknown): value is GridRenderTile {
  if (!value || typeof value !== 'object') {
    return false
  }
  const candidate = value as Partial<GridRenderTile>
  return (
    typeof candidate.tileId === 'number' &&
    candidate.coord !== undefined &&
    typeof candidate.coord.sheetId === 'number' &&
    typeof candidate.coord.sheetOrdinal === 'number' &&
    candidate.bounds !== undefined &&
    candidate.rectInstances instanceof Float32Array &&
    candidate.textMetrics instanceof Float32Array &&
    Array.isArray(candidate.textRuns)
  )
}

function resolveGridRenderTileInputSheetOrdinal(input: {
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}): number {
  return input.sheetOrdinal ?? input.sheetId ?? 0
}

function resolveRenderTileInterestTileKeys(input: GridRenderTileInterestRuntimeInput): readonly TileKey53[] {
  const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
  const keys = new Set<number>()
  const result: number[] = []
  for (const viewport of resolveRenderTileInterestViewports(input)) {
    for (const key of input.gridRuntimeHost.viewportTileKeys({ dprBucket: input.dprBucket, sheetOrdinal, viewport })) {
      if (keys.has(key)) {
        continue
      }
      keys.add(key)
      result.push(key)
    }
  }
  return result
}

function resolveRenderTileInterestViewports(input: GridRenderTileInterestRuntimeInput): readonly Viewport[] {
  const bodyViewport = input.residentViewport ?? input.renderTileViewport
  const freezeRows = Math.max(0, input.freezeRows ?? 0)
  const freezeCols = Math.max(0, input.freezeCols ?? 0)
  const viewports: Viewport[] = []
  addRenderTileInterestViewport(viewports, bodyViewport)
  if (freezeRows > 0) {
    addRenderTileInterestViewport(viewports, {
      colEnd: bodyViewport.colEnd,
      colStart: bodyViewport.colStart,
      rowEnd: freezeRows - 1,
      rowStart: 0,
    })
  }
  if (freezeCols > 0) {
    addRenderTileInterestViewport(viewports, {
      colEnd: freezeCols - 1,
      colStart: 0,
      rowEnd: bodyViewport.rowEnd,
      rowStart: bodyViewport.rowStart,
    })
  }
  if (freezeRows > 0 && freezeCols > 0) {
    addRenderTileInterestViewport(viewports, {
      colEnd: freezeCols - 1,
      colStart: 0,
      rowEnd: freezeRows - 1,
      rowStart: 0,
    })
  }
  return viewports
}

function addRenderTileInterestViewport(viewports: Viewport[], viewport: Viewport): void {
  if (viewport.rowEnd < viewport.rowStart || viewport.colEnd < viewport.colStart) {
    return
  }
  viewports.push(viewport)
}

function sameRenderTileDeltaConnectionIdentity(left: RenderTileDeltaConnectionIdentity, right: RenderTileDeltaConnectionIdentity): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.freezeCols === right.freezeCols &&
    left.freezeRows === right.freezeRows &&
    left.renderTileSource === right.renderTileSource &&
    left.sheetId === right.sheetId &&
    left.sheetName === right.sheetName &&
    left.sheetOrdinal === right.sheetOrdinal &&
    sameOptionalViewportIdentity(left.residentViewport, right.residentViewport) &&
    sameViewportIdentity(left.viewport, right.viewport)
  )
}

function sameWorkbookDeltaConnectionIdentity(left: WorkbookDeltaConnectionIdentity, right: WorkbookDeltaConnectionIdentity): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.renderTileSource === right.renderTileSource &&
    left.sheetId === right.sheetId &&
    left.sheetOrdinal === right.sheetOrdinal
  )
}

function sameLocalInvalidationConnectionIdentity(
  left: LocalInvalidationConnectionIdentity,
  right: LocalInvalidationConnectionIdentity,
): boolean {
  return (
    left.engine === right.engine &&
    left.needsLocalCellInvalidation === right.needsLocalCellInvalidation &&
    left.sheetName === right.sheetName &&
    sameStringListIdentity(left.visibleAddresses, right.visibleAddresses)
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

function shouldForceLocalTilesForWorkbookDelta(batch: WorkbookDeltaBatchLikeV3): boolean {
  if (batch.source === 'workerAuthoritative') {
    return false
  }
  if (batch.source !== 'localOptimistic') {
    return true
  }
  return !hasAxisOnlyWorkbookDeltaDamage(batch)
}

function hasAxisOnlyWorkbookDeltaDamage(batch: WorkbookDeltaBatchLikeV3): boolean {
  return (batch.dirty.axisX.length > 0 || batch.dirty.axisY.length > 0) && batch.dirty.cellRanges.length === 0
}

function sameViewportIdentity(left: Viewport, right: Viewport): boolean {
  return (
    left.colEnd === right.colEnd && left.colStart === right.colStart && left.rowEnd === right.rowEnd && left.rowStart === right.rowStart
  )
}

function sameOptionalViewportIdentity(left: Viewport | undefined, right: Viewport | undefined): boolean {
  if (left === undefined || right === undefined) {
    return left === right
  }
  return sameViewportIdentity(left, right)
}

function matchesRenderTileSheetIdentity(tile: RenderTileSheetIdentity, expected: RenderTileSheetIdentity): boolean {
  if (expected.sheetId !== undefined && expected.sheetOrdinal !== undefined) {
    return tile.sheetId === expected.sheetId && tile.sheetOrdinal === expected.sheetOrdinal
  }
  if (expected.sheetId !== undefined) {
    return tile.sheetId === expected.sheetId
  }
  if (expected.sheetOrdinal !== undefined) {
    return tile.sheetOrdinal === expected.sheetOrdinal
  }
  return false
}

function buildRetainedFixedRenderTileDataPanesCompatibility(
  input: GridRenderTilePaneRuntimeInput,
): RetainedFixedRenderTileDataPanesCompatibility {
  return {
    dprBucket: input.dprBucket,
    freezeCols: input.freezeCols,
    freezeRows: input.freezeRows,
    frozenColumnWidth: input.frozenColumnWidth,
    frozenRowHeight: input.frozenRowHeight,
    hostClientHeight: input.hostClientHeight,
    hostClientWidth: input.hostClientWidth,
    renderTileViewport: input.renderTileViewport,
    residentViewport: input.residentViewport,
    sceneRevision: input.sceneRevision,
    sheetId: input.sheetId,
    sheetOrdinal: input.sheetOrdinal,
    visibleViewport: input.visibleViewport,
  }
}

function sameRetainedFixedRenderTileDataPanesCompatibility(
  left: RetainedFixedRenderTileDataPanesCompatibility,
  right: RetainedFixedRenderTileDataPanesCompatibility,
): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.freezeCols === right.freezeCols &&
    left.freezeRows === right.freezeRows &&
    left.frozenColumnWidth === right.frozenColumnWidth &&
    left.frozenRowHeight === right.frozenRowHeight &&
    left.hostClientHeight === right.hostClientHeight &&
    left.hostClientWidth === right.hostClientWidth &&
    left.sceneRevision === right.sceneRevision &&
    left.sheetId === right.sheetId &&
    left.sheetOrdinal === right.sheetOrdinal &&
    sameViewportIdentity(left.renderTileViewport, right.renderTileViewport) &&
    sameViewportIdentity(left.residentViewport, right.residentViewport) &&
    sameViewportIdentity(left.visibleViewport, right.visibleViewport)
  )
}

function upsertRenderTileIntoHost(gridRuntimeHost: GridRuntimeHost, tile: GridRenderTile): void {
  gridRuntimeHost.tiles.upsertTile({
    axisSeqX: tile.version.axisX,
    axisSeqY: tile.version.axisY,
    byteSizeCpu: estimateTileCpuBytes(tile),
    byteSizeGpu: tile.rectInstances.byteLength + tile.textMetrics.byteLength,
    colTile: tile.coord.colTile,
    dprBucket: tile.coord.dprBucket,
    freezeSeq: tile.version.freeze,
    key: tile.tileId,
    packet: tile,
    rectSeq: tile.version.styles,
    rowTile: tile.coord.rowTile,
    sheetOrdinal: tile.coord.sheetOrdinal,
    state: 'ready',
    styleSeq: tile.version.styles,
    textSeq: tile.version.text,
    valueSeq: tile.version.values,
  })
}

function resolveRenderTileViewportUnion(tiles: readonly GridRenderTile[]): Viewport | null {
  const first = tiles[0]
  if (!first) {
    return null
  }
  let colStart = first.bounds.colStart
  let colEnd = first.bounds.colEnd
  let rowStart = first.bounds.rowStart
  let rowEnd = first.bounds.rowEnd
  for (let index = 1; index < tiles.length; index += 1) {
    const tile = tiles[index]!
    colStart = Math.min(colStart, tile.bounds.colStart)
    colEnd = Math.max(colEnd, tile.bounds.colEnd)
    rowStart = Math.min(rowStart, tile.bounds.rowStart)
    rowEnd = Math.max(rowEnd, tile.bounds.rowEnd)
  }
  return { colEnd, colStart, rowEnd, rowStart }
}

function estimateTileCpuBytes(tile: GridRenderTile): number {
  let textBytes = 0
  for (const run of tile.textRuns) {
    textBytes += run.text.length * 2
    textBytes += run.font.length * 2
    textBytes += run.color.length * 2
  }
  return tile.rectInstances.byteLength + tile.textMetrics.byteLength + textBytes
}
