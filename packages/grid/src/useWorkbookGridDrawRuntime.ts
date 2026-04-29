import type { Dispatch, SetStateAction } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { Item } from './gridTypes.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import type { SheetGridViewportSubscription } from './workbookGridSurfaceTypes.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import { useWorkbookGridRenderPanes } from './useWorkbookGridRenderPanes.js'
import { useWorkbookGridViewportRuntime } from './useWorkbookGridViewportRuntime.js'

type MutableRef<T> = {
  current: T
}

type SortedAxisOverrides = readonly (readonly [number, number])[]

export function useWorkbookGridDrawRuntime(input: {
  readonly columnAxis: GridAxisWorldIndex
  readonly columnWidths: Readonly<Record<number, number>>
  readonly dprBucket: number
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridCameraStore: GridCameraStore
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostElement: HTMLDivElement | null
  readonly liveVisibleRegionRef: MutableRef<VisibleRegionState>
  readonly onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly requiresLiveViewportState: boolean
  readonly restoreViewportTarget?:
    | {
        readonly token: number
        readonly viewport: Viewport
      }
    | undefined
  readonly rowAxis: GridAxisWorldIndex
  readonly rowHeights: Readonly<Record<number, number>>
  readonly scrollTransformRef: MutableRef<WorkbookGridScrollSnapshot>
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly scrollViewportRef: MutableRef<HTMLDivElement | null>
  readonly selectedCell: Item
  readonly setVisibleRegion: Dispatch<SetStateAction<VisibleRegionState>>
  readonly sheetId?: number | undefined
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly subscribeViewport?: SheetGridViewportSubscription | undefined
  readonly syncRuntimeAxes: () => void
  readonly visibleRegion: VisibleRegionState
}) {
  const {
    columnAxis,
    columnWidths,
    dprBucket,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridCameraStore,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    liveVisibleRegionRef,
    onVisibleViewportChange,
    renderTileSource,
    requiresLiveViewportState,
    restoreViewportTarget,
    rowAxis,
    rowHeights,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    selectedCell,
    setVisibleRegion,
    sheetId,
    sheetName,
    shouldUseRemoteRenderTileSource,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    subscribeViewport,
    syncRuntimeAxes,
    visibleRegion,
  } = input
  const viewportResidency = useWorkbookGridViewportRuntime({
    columnAxis,
    engine,
    freezeCols,
    freezeRows,
    gridCameraStore,
    gridRuntimeHost,
    gridMetrics,
    hostElement,
    liveVisibleRegionRef,
    onVisibleViewportChange,
    requiresLiveViewportState,
    rowAxis,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    selectedCell,
    setVisibleRegion,
    sheetName,
    shouldUseRemoteRenderTileSource,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    visibleRegion,
    restoreViewportTarget,
  })

  return useWorkbookGridRenderPanes({
    columnWidths,
    dprBucket,
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
    rowHeights,
    sheetId,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    subscribeViewport,
    viewportResidency,
  })
}
