import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useSyncExternalStore } from 'react'
import type { GridRenderRevisionSnapshot } from '../grid-engine.js'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { Rectangle } from '../gridTypes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
export { TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS, GridDrawSchedulerV3, shouldDeferTypeGpuV3PreloadSync } from './draw-scheduler.js'
export { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import type { WorkbookPaneSurfaceBackendStatusV3 } from './workbook-pane-surface-runtime.js'
import { WorkbookPaneNativeRectLayerV3 } from './WorkbookPaneNativeRectLayerV3.js'
import { packGridTextRunsBufferV3 } from './text-run-buffer.js'
import {
  WorkbookPaneNativeTextLayerV3,
  resolveNativeTextRunSelectionOccludedClipV3,
  type SuppressedNativeTextCellV3,
} from './WorkbookPaneNativeTextLayerV3.js'
import { WorkbookPaneRendererHostRuntimeV3 } from './workbook-pane-renderer-host-runtime.js'

export interface WorkbookPaneRendererV3Props {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly geometry: GridGeometrySnapshot | null
  readonly cameraStore?: GridCameraStore | null
  readonly headerPanes?: readonly GridHeaderPaneState[] | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly preloadTilePanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly renderRevisionSnapshot?: GridRenderRevisionSnapshot | null | undefined
  readonly overlayBuilder?: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | undefined
  readonly overlay?: DynamicGridOverlayBatchV3 | undefined
  readonly onBackendStatusChange?: ((status: WorkbookPaneSurfaceBackendStatusV3) => void) | undefined
  readonly scrollTransformStore?: WorkbookGridScrollStore | null
  readonly selectionOcclusionRanges?: readonly Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>[] | null | undefined
  readonly suppressedTextCell?: SuppressedNativeTextCellV3 | null | undefined
}

export interface WorkbookPaneTextLayerModeV3 {
  readonly hasNativeTextLayerRuns: boolean
  readonly nativeHeaderPanes: readonly GridHeaderPaneState[]
  readonly nativeLayerSource: 'backend-unavailable-live' | 'typegpu-pending-native-text' | 'typegpu-ready-native-visuals'
  readonly nativeTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly showNativeTextLayer: boolean
  readonly showTypeGpuCanvas: boolean
  readonly typeGpuDrawText: boolean
}

export function resolveWorkbookPaneTextLayerModeV3(input: {
  readonly active: boolean
  readonly backendStatus: WorkbookPaneSurfaceBackendStatusV3
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): WorkbookPaneTextLayerModeV3 {
  const showTypeGpuCanvas = input.backendStatus !== 'unavailable'
  const typeGpuDrawText = false
  const nativeHeaderPanes = input.headerPanes
  const nativeTilePanes = input.tilePanes
  const nativeHeaderTextRunCount = countHeaderPaneTextRunsV3(nativeHeaderPanes)
  const nativeTileTextRunCount = countTilePaneTextRunsV3(nativeTilePanes)
  const hasNativeTextLayerRuns = nativeHeaderTextRunCount + nativeTileTextRunCount > 0
  return {
    hasNativeTextLayerRuns,
    nativeHeaderPanes,
    nativeLayerSource:
      input.backendStatus === 'ready'
        ? 'typegpu-ready-native-visuals'
        : showTypeGpuCanvas
          ? 'typegpu-pending-native-text'
          : 'backend-unavailable-live',
    nativeTilePanes,
    showNativeTextLayer: input.active && hasNativeTextLayerRuns,
    showTypeGpuCanvas,
    typeGpuDrawText,
  }
}

export const WorkbookPaneRendererV3 = memo(function WorkbookPaneRendererV3({
  active,
  cameraStore = null,
  geometry,
  headerPanes = [],
  host,
  onBackendStatusChange,
  overlay,
  overlayBuilder,
  preloadTilePanes = [],
  renderRevisionSnapshot = null,
  scrollTransformStore = null,
  selectionOcclusionRanges = null,
  suppressedTextCell = null,
  tilePanes,
}: WorkbookPaneRendererV3Props) {
  const hostRuntimeRef = useRef<WorkbookPaneRendererHostRuntimeV3 | null>(null)
  const hostRuntimeLifetimeRef = useRef(0)
  if (!hostRuntimeRef.current) {
    hostRuntimeRef.current = new WorkbookPaneRendererHostRuntimeV3()
  }
  const hostRuntime = hostRuntimeRef.current
  const backendStatus = useSyncExternalStore(
    hostRuntime.subscribeBackendStatus,
    hostRuntime.getBackendStatusSnapshot,
    hostRuntime.getBackendStatusSnapshot,
  )
  useEffect(() => {
    onBackendStatusChange?.(backendStatus)
  }, [backendStatus, onBackendStatusChange])
  const frameProofStatus = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getFrameProofStatusSnapshot,
    hostRuntime.getFrameProofStatusSnapshot,
  )
  const frameProofSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getFrameProofSignatureSnapshot,
    hostRuntime.getFrameProofSignatureSnapshot,
  )
  const hasPresentedFrame = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getHasPresentedFrameSnapshot,
    hostRuntime.getHasPresentedFrameSnapshot,
  )
  const presentedFrameProofSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedFrameProofSignatureSnapshot,
    hostRuntime.getPresentedFrameProofSignatureSnapshot,
  )
  const presentedVisualFrame = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedVisualFrameSnapshot,
    hostRuntime.getPresentedVisualFrameSnapshot,
  )
  const currentContentSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getCurrentContentSignatureSnapshot,
    hostRuntime.getCurrentContentSignatureSnapshot,
  )
  const currentRectCount = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getCurrentRectCountSnapshot,
    hostRuntime.getCurrentRectCountSnapshot,
  )
  const currentRectSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getCurrentRectSignatureSnapshot,
    hostRuntime.getCurrentRectSignatureSnapshot,
  )
  const currentSceneOwnershipSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getVisibleSceneOwnershipSignatureSnapshot,
    hostRuntime.getVisibleSceneOwnershipSignatureSnapshot,
  )
  const currentSceneOwnershipEpoch = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getVisibleSceneOwnershipEpochSnapshot,
    hostRuntime.getVisibleSceneOwnershipEpochSnapshot,
  )
  const currentSceneOwnershipEpochSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getVisibleSceneOwnershipEpochSignatureSnapshot,
    hostRuntime.getVisibleSceneOwnershipEpochSignatureSnapshot,
  )
  const currentTextRunCount = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getCurrentTextRunCountSnapshot,
    hostRuntime.getCurrentTextRunCountSnapshot,
  )
  const currentTextSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getCurrentTextSignatureSnapshot,
    hostRuntime.getCurrentTextSignatureSnapshot,
  )
  const presentedContentSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedContentSignatureSnapshot,
    hostRuntime.getPresentedContentSignatureSnapshot,
  )
  const presentedRectCount = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedRectCountSnapshot,
    hostRuntime.getPresentedRectCountSnapshot,
  )
  const presentedRectSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedRectSignatureSnapshot,
    hostRuntime.getPresentedRectSignatureSnapshot,
  )
  const presentedSceneOwnershipSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedVisibleSceneOwnershipSignatureSnapshot,
    hostRuntime.getPresentedVisibleSceneOwnershipSignatureSnapshot,
  )
  const presentedSceneOwnershipEpoch = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedVisibleSceneOwnershipEpochSnapshot,
    hostRuntime.getPresentedVisibleSceneOwnershipEpochSnapshot,
  )
  const presentedSceneOwnershipEpochSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedVisibleSceneOwnershipEpochSignatureSnapshot,
    hostRuntime.getPresentedVisibleSceneOwnershipEpochSignatureSnapshot,
  )
  const presentedTextRunCount = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedTextRunCountSnapshot,
    hostRuntime.getPresentedTextRunCountSnapshot,
  )
  const presentedTextSignature = useSyncExternalStore(
    hostRuntime.subscribeFrameProofStatus,
    hostRuntime.getPresentedTextSignatureSnapshot,
    hostRuntime.getPresentedTextSignatureSnapshot,
  )
  const hasPresentedAnyVisibleFrame = presentedFrameProofSignature.length > 0
  const hasPresentedVisibleFrame = hasPresentedFrame && frameProofStatus === 'presented'
  const sceneOwnershipEpochMatchesPresentedFrame =
    currentSceneOwnershipEpochSignature.length > 0 &&
    presentedSceneOwnershipEpochSignature.length > 0 &&
    currentSceneOwnershipEpochSignature === presentedSceneOwnershipEpochSignature

  const setCanvasRef = useCallback(
    (canvas: HTMLCanvasElement | null) => {
      hostRuntime.setCanvas(canvas)
    },
    [hostRuntime],
  )
  const headerTextRunCount = countHeaderPaneTextRunsV3(headerPanes)
  const suppressedTextTilePanes = useMemo(
    () =>
      resolveWorkbookPaneSuppressedTextCellTilePanesV3({
        suppressedTextCell,
        tilePanes,
      }),
    [suppressedTextCell, tilePanes],
  )
  const typeGpuTilePanes = useMemo(
    () =>
      resolveWorkbookPaneSelectionOccludedTilePanesV3({
        geometry,
        selectionOcclusionRanges,
        tilePanes: suppressedTextTilePanes,
      }),
    [geometry, selectionOcclusionRanges, suppressedTextTilePanes],
  )
  const tileTextRunCount = countTilePaneTextRunsV3(typeGpuTilePanes)
  const textLayerMode = resolveWorkbookPaneTextLayerModeV3({
    active,
    backendStatus,
    headerPanes,
    tilePanes: suppressedTextTilePanes,
  })
  const { showTypeGpuCanvas, typeGpuDrawText, nativeLayerSource, nativeHeaderPanes, nativeTilePanes, showNativeTextLayer } = textLayerMode
  const presentedHeaderPanes = presentedVisualFrame?.headerPanes ?? []
  const presentedTilePanes = presentedVisualFrame?.tilePanes ?? []
  const presentedHeaderTextRunCount = countHeaderPaneTextRunsV3(presentedHeaderPanes)
  const presentedTileTextRunCount = countTilePaneTextRunsV3(presentedTilePanes)
  const nativeHeaderTextRunCount = countHeaderPaneTextRunsV3(nativeHeaderPanes)
  const nativeTileTextRunCount = countTilePaneTextRunsV3(nativeTilePanes)

  useLayoutEffect(() => {
    hostRuntime.updateProps({
      active,
      cameraStore,
      drawText: typeGpuDrawText,
      geometry,
      headerPanes,
      host,
      overlay: overlay ?? null,
      overlayBuilder: overlayBuilder ?? null,
      preloadTilePanes,
      renderRevisionSnapshot,
      scrollTransformStore,
      tilePanes: typeGpuTilePanes,
    })
  }, [
    active,
    cameraStore,
    geometry,
    headerPanes,
    host,
    hostRuntime,
    overlay,
    overlayBuilder,
    preloadTilePanes,
    renderRevisionSnapshot,
    scrollTransformStore,
    typeGpuDrawText,
    typeGpuTilePanes,
  ])

  useEffect(() => {
    const lifetime = hostRuntimeLifetimeRef.current + 1
    hostRuntimeLifetimeRef.current = lifetime
    return () => {
      queueMicrotask(() => {
        if (hostRuntimeLifetimeRef.current !== lifetime) {
          return
        }
        hostRuntime.dispose()
        if (hostRuntimeRef.current === hostRuntime) {
          hostRuntimeRef.current = null
        }
      })
    }
  }, [hostRuntime])

  if (!active || !host) {
    return null
  }
  const tileSceneRevision = resolveWorkbookPaneTileSceneRevisionV3(typeGpuTilePanes)
  const tileSceneCameraSeq = resolveWorkbookPaneTileSceneCameraSeqV3(typeGpuTilePanes)
  const visibleRenderRevision = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, tileSceneRevision)
  const visibleRenderCameraSeq = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, tileSceneCameraSeq)
  const visibleProjectedRenderRevision = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, renderRevisionSnapshot?.projectedRevision)
  const visibleLocalRenderRevision = resolveWorkbookPanePresentedRevisionV3(frameProofStatus, renderRevisionSnapshot?.localRevision)
  const visibleAuthoritativeRenderRevision = resolveWorkbookPanePresentedRevisionV3(
    frameProofStatus,
    renderRevisionSnapshot?.authoritativeRevision,
  )

  return (
    <>
      {showTypeGpuCanvas ? (
        <canvas
          aria-hidden="true"
          className="pointer-events-none absolute inset-0 z-10"
          data-pane-renderer="workbook-pane-renderer-v3"
          data-renderer-mode="typegpu-v3"
          data-testid="grid-pane-renderer"
          data-v3-backend-status={backendStatus}
          data-v3-body-world-x={geometry?.camera.bodyWorldX ?? 0}
          data-v3-body-world-y={geometry?.camera.bodyWorldY ?? 0}
          data-v3-canvas-proof-layer="disabled"
          data-v3-current-content-signature={currentContentSignature}
          data-v3-current-fill-handle-revision={currentSceneOwnershipEpoch?.fillHandleRevision ?? ''}
          data-v3-current-rect-count={currentRectCount}
          data-v3-current-rect-signature={currentRectSignature}
          data-v3-current-scene-epoch={currentSceneOwnershipEpoch?.sceneEpoch ?? ''}
          data-v3-current-scene-epoch-signature={currentSceneOwnershipEpochSignature}
          data-v3-current-scene-ownership-signature={currentSceneOwnershipSignature}
          data-v3-current-selection-revision={currentSceneOwnershipEpoch?.selectionRevision ?? ''}
          data-v3-current-semantic-mutation-revision={currentSceneOwnershipEpoch?.semanticMutationRevision ?? ''}
          data-v3-current-text-run-count={currentTextRunCount}
          data-v3-current-text-signature={currentTextSignature}
          data-v3-current-viewport-revision={currentSceneOwnershipEpoch?.viewportRevision ?? ''}
          data-v3-current-workbook-revision={currentSceneOwnershipEpoch?.workbookRevision ?? ''}
          data-v3-draw-text={typeGpuDrawText ? 'true' : 'false'}
          data-v3-frame-proof-status={frameProofStatus}
          data-v3-frame-proof-signature={frameProofSignature}
          data-v3-has-presented-frame={hasPresentedFrame ? 'true' : 'false'}
          data-v3-has-presented-any-frame={hasPresentedAnyVisibleFrame ? 'true' : 'false'}
          data-v3-has-presented-visible-frame={hasPresentedVisibleFrame ? 'true' : 'false'}
          data-v3-header-pane-count={headerPanes.length}
          data-v3-header-text-run-count={headerTextRunCount}
          data-v3-authoritative-render-revision={renderRevisionSnapshot?.authoritativeRevision ?? ''}
          data-v3-local-render-revision={renderRevisionSnapshot?.localRevision ?? ''}
          data-v3-presented-frame-proof-signature={presentedFrameProofSignature}
          data-v3-presented-camera-seq={presentedVisualFrame?.cameraSeq ?? ''}
          data-v3-presented-content-signature={presentedContentSignature}
          data-v3-presented-draw-text={presentedVisualFrame ? (presentedVisualFrame.drawText ? 'true' : 'false') : ''}
          data-v3-presented-fill-handle-revision={presentedSceneOwnershipEpoch?.fillHandleRevision ?? ''}
          data-v3-presented-header-pane-count={presentedHeaderPanes.length}
          data-v3-presented-header-text-run-count={presentedHeaderTextRunCount}
          data-v3-native-layer-source={nativeLayerSource}
          data-v3-native-header-pane-count={nativeHeaderPanes.length}
          data-v3-native-header-text-run-count={nativeHeaderTextRunCount}
          data-v3-presented-overlay-camera-seq={presentedVisualFrame?.overlayCameraSeq ?? ''}
          data-v3-presented-overlay-rect-count={presentedVisualFrame?.overlayRectCount ?? ''}
          data-v3-presented-overlay-rect-signature={presentedVisualFrame?.overlayRectSignature ?? ''}
          data-v3-presented-overlay-seq={presentedVisualFrame?.overlaySeq ?? ''}
          data-v3-presented-rect-count={presentedRectCount}
          data-v3-presented-rect-signature={presentedRectSignature}
          data-v3-presented-render-tx={presentedVisualFrame?.scrollSnapshot.renderTx ?? presentedVisualFrame?.scrollSnapshot.tx ?? ''}
          data-v3-presented-render-ty={presentedVisualFrame?.scrollSnapshot.renderTy ?? presentedVisualFrame?.scrollSnapshot.ty ?? ''}
          data-v3-presented-scene-epoch={presentedSceneOwnershipEpoch?.sceneEpoch ?? ''}
          data-v3-presented-scene-epoch-signature={presentedSceneOwnershipEpochSignature}
          data-v3-presented-scene-ownership-signature={presentedSceneOwnershipSignature}
          data-v3-presented-selection-revision={presentedSceneOwnershipEpoch?.selectionRevision ?? ''}
          data-v3-presented-semantic-mutation-revision={presentedSceneOwnershipEpoch?.semanticMutationRevision ?? ''}
          data-v3-presented-scroll-left={presentedVisualFrame?.scrollSnapshot.scrollLeft ?? ''}
          data-v3-presented-scroll-top={presentedVisualFrame?.scrollSnapshot.scrollTop ?? ''}
          data-v3-presented-text-signature={presentedTextSignature}
          data-v3-presented-text-run-count={presentedTileTextRunCount}
          data-v3-presented-tile-pane-count={presentedTilePanes.length}
          data-v3-presented-viewport-revision={presentedSceneOwnershipEpoch?.viewportRevision ?? ''}
          data-v3-presented-visible-text-run-count={presentedTextRunCount}
          data-v3-presented-workbook-revision={presentedSceneOwnershipEpoch?.workbookRevision ?? ''}
          data-v3-native-text-run-count={nativeTileTextRunCount}
          data-v3-native-tile-pane-count={nativeTilePanes.length}
          data-v3-preload-pane-count={preloadTilePanes.length}
          data-v3-projected-render-revision={renderRevisionSnapshot?.projectedRevision ?? ''}
          data-v3-text-run-count={tileTextRunCount}
          data-v3-tile-scene-camera-seq={tileSceneCameraSeq ?? ''}
          data-v3-tile-scene-revision={tileSceneRevision ?? ''}
          data-v3-tile-pane-count={typeGpuTilePanes.length}
          data-v3-visible-scene-epoch-matches-presented-frame={sceneOwnershipEpochMatchesPresentedFrame ? 'true' : 'false'}
          data-v3-visible-authoritative-render-revision={visibleAuthoritativeRenderRevision ?? ''}
          data-v3-visible-local-render-revision={visibleLocalRenderRevision ?? ''}
          data-v3-visible-projected-render-revision={visibleProjectedRenderRevision ?? ''}
          data-v3-visible-render-camera-seq={visibleRenderCameraSeq ?? ''}
          data-v3-visible-render-revision={visibleRenderRevision ?? ''}
          ref={setCanvasRef}
          style={{ backgroundColor: 'transparent', contain: 'strict', height: '100%', opacity: 1, width: '100%' }}
        />
      ) : null}
      {!showTypeGpuCanvas ? (
        <div
          aria-hidden="true"
          className="pointer-events-none absolute inset-0 z-10"
          data-renderer-mode="typegpu-v3-unavailable"
          data-testid="grid-pane-renderer-unavailable"
          data-v3-backend-status={backendStatus}
          data-v3-canvas-proof-layer="disabled"
        />
      ) : null}
      <WorkbookPaneNativeRectLayerV3
        active={active}
        cameraStore={cameraStore}
        geometry={geometry}
        headerPanes={nativeHeaderPanes}
        presentedScrollSnapshot={presentedVisualFrame?.scrollSnapshot ?? null}
        scrollTransformStore={scrollTransformStore}
        tilePanes={nativeTilePanes}
      />
      {showNativeTextLayer ? (
        <WorkbookPaneNativeTextLayerV3
          active={active}
          cameraStore={cameraStore}
          geometry={geometry}
          headerPanes={nativeHeaderPanes}
          presentedScrollSnapshot={presentedVisualFrame?.scrollSnapshot ?? null}
          scrollTransformStore={scrollTransformStore}
          selectionOcclusionRanges={selectionOcclusionRanges}
          suppressedTextCell={suppressedTextCell}
          tilePanes={nativeTilePanes}
        />
      ) : null}
    </>
  )
})

function countHeaderPaneTextRunsV3(headerPanes: readonly GridHeaderPaneState[]): number {
  return headerPanes.reduce((total, pane) => total + pane.textRuns.length, 0)
}

function countTilePaneTextRunsV3(tilePanes: readonly WorkbookRenderTilePaneState[]): number {
  return tilePanes.reduce((total, pane) => total + pane.tile.textRuns.length, 0)
}

export function resolveWorkbookPaneSelectionOccludedTilePanesV3(input: {
  readonly geometry: GridGeometrySnapshot | null
  readonly selectionOcclusionRanges?: readonly Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>[] | null | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): readonly WorkbookRenderTilePaneState[] {
  const ranges = input.selectionOcclusionRanges ?? []
  if (!input.geometry || ranges.length === 0) {
    return input.tilePanes
  }

  let panesChanged = false
  const nextPanes = input.tilePanes.map((pane) => {
    if (pane.tile.textRuns.length === 0) {
      return pane
    }

    let textRunsChanged = false
    const nextTextRuns: (typeof pane.tile.textRuns)[number][] = []
    for (const run of pane.tile.textRuns) {
      const clip = resolveNativeTextRunSelectionOccludedClipV3({
        clipHeight: run.clipHeight,
        clipWidth: run.clipWidth,
        clipX: run.clipX,
        clipY: run.clipY,
        geometry: input.geometry,
        pane,
        run,
        selectionOcclusionRanges: ranges,
      })
      if (!clip) {
        textRunsChanged = true
        continue
      }
      if (clip.clipX !== run.clipX || clip.clipY !== run.clipY || clip.clipWidth !== run.clipWidth || clip.clipHeight !== run.clipHeight) {
        textRunsChanged = true
        nextTextRuns.push({
          ...run,
          clipHeight: clip.clipHeight,
          clipWidth: clip.clipWidth,
          clipX: clip.clipX,
          clipY: clip.clipY,
        })
        continue
      }
      nextTextRuns.push(run)
    }

    if (!textRunsChanged) {
      return pane
    }

    panesChanged = true
    return {
      ...pane,
      tile: {
        ...pane.tile,
        dirty: undefined,
        textCount: nextTextRuns.length,
        textRuns: nextTextRuns,
        textSignature: undefined,
      },
    }
  })

  return panesChanged ? nextPanes : input.tilePanes
}

export function resolveWorkbookPaneSuppressedTextCellTilePanesV3(input: {
  readonly suppressedTextCell?: SuppressedNativeTextCellV3 | null | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): readonly WorkbookRenderTilePaneState[] {
  const suppressedTextCell = input.suppressedTextCell ?? null
  if (!suppressedTextCell) {
    return input.tilePanes
  }

  let panesChanged = false
  const nextPanes = input.tilePanes.map((pane) => {
    if (pane.tile.textRuns.length === 0) {
      return pane
    }

    const nextTextRuns = pane.tile.textRuns.filter((run) => run.row !== suppressedTextCell.row || run.col !== suppressedTextCell.col)
    if (nextTextRuns.length === pane.tile.textRuns.length) {
      return pane
    }

    panesChanged = true
    const textBuffer = packGridTextRunsBufferV3(nextTextRuns)
    return {
      ...pane,
      tile: {
        ...pane.tile,
        dirty: undefined,
        textCount: textBuffer.textCount,
        textMetrics: textBuffer.textMetrics,
        textRuns: textBuffer.textRuns,
        textSignature: textBuffer.textSignature,
      },
    }
  })

  return panesChanged ? nextPanes : input.tilePanes
}

export function resolveWorkbookPaneTileSceneRevisionV3(tilePanes: readonly WorkbookRenderTilePaneState[]): number | null {
  return maxTilePaneField(tilePanes, (pane) => pane.tile.lastBatchId)
}

export function resolveWorkbookPanePresentedRevisionV3(
  frameProofStatus: 'idle' | 'pending' | 'presented',
  revision: number | null | undefined,
): number | null {
  return frameProofStatus === 'presented' && revision !== null && revision !== undefined ? revision : null
}

export function resolveWorkbookPaneTileSceneCameraSeqV3(tilePanes: readonly WorkbookRenderTilePaneState[]): number | null {
  return maxTilePaneField(tilePanes, (pane) => pane.tile.lastCameraSeq)
}

function maxTilePaneField(
  tilePanes: readonly WorkbookRenderTilePaneState[],
  readValue: (pane: WorkbookRenderTilePaneState) => number,
): number | null {
  let result: number | null = null
  for (const pane of tilePanes) {
    const value = readValue(pane)
    if (!Number.isFinite(value)) {
      continue
    }
    result = result === null ? value : Math.max(result, value)
  }
  return result
}
