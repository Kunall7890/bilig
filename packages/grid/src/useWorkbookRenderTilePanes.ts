import { useEffect, useMemo, useSyncExternalStore } from 'react'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import type { Item } from './gridTypes.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from './renderer-v3/render-tile-pane-state.js'
import type { GridTileReadinessSnapshotV3 } from './runtime/gridTileCoordinator.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export interface WorkbookRenderTilePanesState {
  readonly preloadDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly renderTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly tileReadiness: GridTileReadinessSnapshotV3
}

export function useWorkbookRenderTilePanes(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly dprBucket: number
  readonly editingCell?: Item | null | undefined
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostElement: HTMLDivElement | null
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sceneRevision: number
  readonly selectedCell?: Item | undefined
  readonly selectedCellSnapshot?: CellSnapshot | null | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly visibleAddresses: readonly string[]
  readonly visibleViewport: Viewport
}): WorkbookRenderTilePanesState {
  const {
    columnWidths,
    dprBucket,
    editingCell,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    renderTileSource,
    renderTileViewport,
    residentViewport,
    rowHeights,
    sceneRevision,
    selectedCell,
    selectedCellSnapshot,
    sheetId,
    sheetOrdinal,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    visibleAddresses,
    visibleViewport,
  } = input
  const renderTileBridgeState = useSyncExternalStore(
    (listener) => gridRuntimeHost.subscribeRenderTileBridgeState(listener),
    () => gridRuntimeHost.snapshotRenderTileBridgeState(),
    () => gridRuntimeHost.snapshotRenderTileBridgeState(),
  )
  const renderRevisionSnapshot = engine.getRenderRevisionSnapshot?.()
  const renderRevisionKey = [
    renderRevisionSnapshot?.authoritativeRevision ?? 'none',
    renderRevisionSnapshot?.localRevision ?? 'none',
    renderRevisionSnapshot?.projectedRevision ?? 'none',
    renderRevisionSnapshot?.tileSceneRevision ?? 'none',
    renderRevisionSnapshot?.tileSceneCameraSeq ?? 'none',
  ].join(':')

  useEffect(() => {
    return () => {
      gridRuntimeHost.disconnectRenderTileConnections()
    }
  }, [gridRuntimeHost])

  const state = useMemo<WorkbookRenderTilePanesState & { readonly needsLocalCellInvalidation: boolean }>(() => {
    void renderTileBridgeState.renderTileRevision
    void renderTileBridgeState.localFallbackRevision
    void renderRevisionKey
    return gridRuntimeHost.resolveRenderTilePanes({
      columnWidths,
      dprBucket,
      editingCell,
      engine,
      freezeCols,
      freezeRows,
      forceLocalTiles: renderTileBridgeState.forceLocalTiles,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics,
      hostClientHeight,
      hostClientWidth,
      hostReady: hostElement !== null,
      renderTileSource,
      renderTileViewport,
      residentViewport,
      rowHeights,
      sceneRevision,
      selectedCell,
      selectedCellSnapshot,
      sheetId,
      sheetOrdinal,
      sheetName,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      visibleViewport,
    })
  }, [
    columnWidths,
    dprBucket,
    editingCell,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    renderTileBridgeState,
    renderRevisionKey,
    renderTileSource,
    renderTileViewport,
    residentViewport,
    rowHeights,
    sceneRevision,
    selectedCell,
    selectedCellSnapshot,
    sheetId,
    sheetOrdinal,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    visibleViewport,
  ])

  useEffect(() => {
    gridRuntimeHost.noteRenderTileReadiness(state.tileReadiness)
  }, [gridRuntimeHost, state.tileReadiness])

  useEffect(() => {
    gridRuntimeHost.syncRenderTileConnections({
      dprBucket,
      engine,
      freezeCols,
      freezeRows,
      needsLocalCellInvalidation: state.needsLocalCellInvalidation,
      renderTileSource,
      renderTileViewport,
      residentViewport,
      sheetId,
      sheetOrdinal,
      sheetName,
      visibleAddresses,
      visibleViewport,
    })
  }, [
    dprBucket,
    engine,
    freezeCols,
    freezeRows,
    gridRuntimeHost,
    renderTileSource,
    renderTileViewport,
    residentViewport,
    sheetId,
    sheetName,
    sheetOrdinal,
    state.needsLocalCellInvalidation,
    visibleAddresses,
    visibleViewport,
  ])

  return state
}
