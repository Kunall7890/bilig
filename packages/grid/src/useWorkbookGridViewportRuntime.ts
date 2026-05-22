import type { Dispatch, SetStateAction } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import type { GridMetrics } from './gridMetrics.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { Item } from './gridTypes.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { useWorkbookViewportResidencyState } from './useWorkbookViewportResidencyState.js'
import { useWorkbookViewportScrollRuntime } from './useWorkbookViewportScrollRuntime.js'

type MutableRef<T> = {
  current: T
}

export function useWorkbookGridViewportRuntime(input: {
  readonly columnAxis: GridAxisWorldIndex
  readonly editingCell?: Item | null | undefined
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly gridCameraStore: GridCameraStore
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostElement: HTMLDivElement | null
  readonly liveVisibleRegionRef: MutableRef<VisibleRegionState>
  readonly onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined
  readonly requiresLiveViewportState: boolean
  readonly restoreViewportTarget?:
    | {
        readonly token: number
        readonly viewport: Viewport
      }
    | undefined
  readonly rowAxis: GridAxisWorldIndex
  readonly scrollTransformRef: MutableRef<WorkbookGridScrollSnapshot>
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly scrollViewportRef: MutableRef<HTMLDivElement | null>
  readonly selectedCell: Item
  readonly setVisibleRegion: Dispatch<SetStateAction<VisibleRegionState>>
  readonly sheetName: string
  readonly shouldUseRemoteRenderTileSource: boolean
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly syncRuntimeAxes: () => void
  readonly visibleRegion: VisibleRegionState
}) {
  const {
    columnAxis,
    editingCell,
    engine,
    freezeCols,
    freezeRows,
    gridCameraStore,
    gridMetrics,
    gridRuntimeHost,
    hostElement,
    liveVisibleRegionRef,
    onVisibleViewportChange,
    requiresLiveViewportState,
    restoreViewportTarget,
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
  } = input
  const viewportResidency = useWorkbookViewportResidencyState({
    engine,
    freezeCols,
    freezeRows,
    gridRuntimeHost,
    sheetName,
    shouldUseRemoteRenderTileSource,
    visibleRegion,
  })

  useWorkbookViewportScrollRuntime({
    columnAxis,
    editingCell,
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
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    viewport: viewportResidency.viewport,
    restoreViewportTarget,
    setVisibleRegion,
  })

  return viewportResidency
}
