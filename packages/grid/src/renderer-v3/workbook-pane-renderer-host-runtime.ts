import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridRenderRevisionSnapshot } from '../grid-engine.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { resolveGridTextTileRevisionKeyV3 } from './typegpu-tile-resource-revisions.js'
import {
  WorkbookPaneRendererRuntimeV3,
  resolveWorkbookPaneRendererGeometryV3,
  resolveTypeGpuV3DrawScrollSnapshot,
  type TypeGpuSurfaceSizeV3,
  type WorkbookPaneFrameResultV3,
  type WorkbookPanePresentedVisualFrameV3,
} from './workbook-pane-renderer-runtime.js'
import {
  resolveWorkbookPaneVisibleSceneProofV3,
  type WorkbookPaneVisiblePayloadProofV3,
  type WorkbookPaneVisibleSceneOwnershipEpochV3,
} from './workbook-pane-visible-scene-proof.js'
import {
  EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3,
  WorkbookPaneSurfaceRuntimeV3,
  type WorkbookPaneSurfaceBackendStatusV3,
  type WorkbookPaneSurfaceSnapshotV3,
} from './workbook-pane-surface-runtime.js'

export interface WorkbookPaneRendererHostPropsV3 {
  readonly active: boolean
  readonly cameraStore: GridCameraStore | null
  readonly drawText: boolean
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly host: HTMLDivElement | null
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly overlayBuilder: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | null
  readonly preloadTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly renderRevisionSnapshot: GridRenderRevisionSnapshot | null
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

export interface WorkbookPaneRendererHostRuntimeOptionsV3 {
  readonly rendererRuntime?: WorkbookPaneRendererRuntimeV3 | undefined
  readonly surfaceRuntime?: WorkbookPaneSurfaceRuntimeV3 | undefined
}

export type WorkbookPaneFrameProofStatusV3 = 'idle' | 'pending' | 'presented'

const EMPTY_HOST_PROPS: WorkbookPaneRendererHostPropsV3 = Object.freeze({
  active: false,
  cameraStore: null,
  drawText: true,
  geometry: null,
  headerPanes: [],
  host: null,
  overlay: null,
  overlayBuilder: null,
  preloadTilePanes: [],
  renderRevisionSnapshot: null,
  scrollTransformStore: null,
  tilePanes: [],
})

const EMPTY_VISIBLE_PAYLOAD_PROOF: WorkbookPaneVisiblePayloadProofV3 = Object.freeze({
  contentSignature: '',
  rectCount: 0,
  rectSignature: '',
  textRunCount: 0,
  textSignature: '',
})

export class WorkbookPaneRendererHostRuntimeV3 {
  private canvas: HTMLCanvasElement | null = null
  private disposed = false
  private readonly backendStatusListeners = new Set<() => void>()
  private readonly frameProofListeners = new Set<() => void>()
  private frameProofSignature = ''
  private frameProofStatus: WorkbookPaneFrameProofStatusV3 = 'idle'
  private hasPresentedFrame = false
  private payloadProof: WorkbookPaneVisiblePayloadProofV3 = EMPTY_VISIBLE_PAYLOAD_PROOF
  private presentedFrameProofSignature = ''
  private presentedPayloadProof: WorkbookPaneVisiblePayloadProofV3 = EMPTY_VISIBLE_PAYLOAD_PROOF
  private presentedVisibleSceneOwnershipEpoch: WorkbookPaneVisibleSceneOwnershipEpochV3 | null = null
  private presentedVisibleSceneOwnershipEpochSignature = ''
  private presentedVisibleSceneOwnershipSignature = ''
  private presentedVisualFrame: WorkbookPanePresentedVisualFrameV3 | null = null
  private props: WorkbookPaneRendererHostPropsV3 = EMPTY_HOST_PROPS
  private readonly rendererRuntime: WorkbookPaneRendererRuntimeV3
  private surfaceBackendStatus: WorkbookPaneSurfaceBackendStatusV3
  private surfaceSnapshot: WorkbookPaneSurfaceSnapshotV3 = EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3
  private readonly surfaceRuntime: WorkbookPaneSurfaceRuntimeV3
  private readonly unsubscribeSurface: () => void
  private visibleSceneOwnershipEpoch: WorkbookPaneVisibleSceneOwnershipEpochV3 | null = null
  private visibleSceneOwnershipEpochSignature = ''
  private visibleSceneOwnershipSignature = ''

  constructor(options: WorkbookPaneRendererHostRuntimeOptionsV3 = {}) {
    this.rendererRuntime = options.rendererRuntime ?? new WorkbookPaneRendererRuntimeV3()
    this.rendererRuntime.setInputSignalListener(() => this.handleInputSignal())
    this.rendererRuntime.setFrameResultListener((result) => this.handleFrameResult(result))
    this.surfaceRuntime = options.surfaceRuntime ?? new WorkbookPaneSurfaceRuntimeV3()
    this.surfaceBackendStatus = this.surfaceRuntime.getSnapshot().backendStatus
    this.unsubscribeSurface = this.surfaceRuntime.subscribe((snapshot) => {
      this.surfaceSnapshot = snapshot
      this.syncFrameProofSignature(this.props)
      if (this.surfaceBackendStatus !== snapshot.backendStatus) {
        this.surfaceBackendStatus = snapshot.backendStatus
        this.emitBackendStatus()
      }
      if (snapshot.backendStatus !== 'ready') {
        this.setFrameProofStatus(this.frameProofSignature ? 'pending' : 'idle')
      }
      this.applyRendererState()
      this.requestRenderDraw()
    })
  }

  readonly getBackendStatusSnapshot = (): WorkbookPaneSurfaceBackendStatusV3 => this.surfaceBackendStatus
  readonly getFrameProofSignatureSnapshot = (): string => this.frameProofSignature
  readonly getFrameProofStatusSnapshot = (): WorkbookPaneFrameProofStatusV3 => this.frameProofStatus
  readonly getHasPresentedFrameSnapshot = (): boolean => this.hasPresentedFrame
  readonly getPresentedFrameProofSignatureSnapshot = (): string => this.presentedFrameProofSignature
  readonly getCurrentContentSignatureSnapshot = (): string => this.payloadProof.contentSignature
  readonly getCurrentRectCountSnapshot = (): number => this.payloadProof.rectCount
  readonly getCurrentRectSignatureSnapshot = (): string => this.payloadProof.rectSignature
  readonly getCurrentTextRunCountSnapshot = (): number => this.payloadProof.textRunCount
  readonly getCurrentTextSignatureSnapshot = (): string => this.payloadProof.textSignature
  readonly getPresentedContentSignatureSnapshot = (): string => this.presentedPayloadProof.contentSignature
  readonly getPresentedRectCountSnapshot = (): number => this.presentedPayloadProof.rectCount
  readonly getPresentedRectSignatureSnapshot = (): string => this.presentedPayloadProof.rectSignature
  readonly getPresentedTextRunCountSnapshot = (): number => this.presentedPayloadProof.textRunCount
  readonly getPresentedTextSignatureSnapshot = (): string => this.presentedPayloadProof.textSignature
  readonly getPresentedVisibleSceneOwnershipEpochSnapshot = (): WorkbookPaneVisibleSceneOwnershipEpochV3 | null =>
    this.presentedVisibleSceneOwnershipEpoch
  readonly getPresentedVisibleSceneOwnershipEpochSignatureSnapshot = (): string => this.presentedVisibleSceneOwnershipEpochSignature
  readonly getPresentedVisibleSceneOwnershipSignatureSnapshot = (): string => this.presentedVisibleSceneOwnershipSignature
  readonly getPresentedVisualFrameSnapshot = (): WorkbookPanePresentedVisualFrameV3 | null => this.presentedVisualFrame
  readonly getVisibleSceneOwnershipEpochSnapshot = (): WorkbookPaneVisibleSceneOwnershipEpochV3 | null => this.visibleSceneOwnershipEpoch
  readonly getVisibleSceneOwnershipEpochSignatureSnapshot = (): string => this.visibleSceneOwnershipEpochSignature
  readonly getVisibleSceneOwnershipSignatureSnapshot = (): string => this.visibleSceneOwnershipSignature

  readonly subscribeBackendStatus = (listener: () => void): (() => void) => {
    if (this.disposed) {
      return () => {}
    }
    this.backendStatusListeners.add(listener)
    return () => {
      this.backendStatusListeners.delete(listener)
    }
  }

  readonly subscribeFrameProofStatus = (listener: () => void): (() => void) => {
    if (this.disposed) {
      return () => {}
    }
    this.frameProofListeners.add(listener)
    return () => {
      this.frameProofListeners.delete(listener)
    }
  }

  updateProps(props: WorkbookPaneRendererHostPropsV3): void {
    if (this.disposed) {
      return
    }
    this.syncFrameProofSignature(props)
    this.props = props
    this.surfaceRuntime.setHost(props.host)
    this.surfaceRuntime.setActive(props.active)
    this.syncCanvasTarget()
    this.applyRendererState()
    this.requestRenderDraw()
  }

  setCanvas(canvas: HTMLCanvasElement | null): void {
    if (this.canvas === canvas) {
      return
    }
    this.canvas = canvas
    this.syncCanvasTarget()
  }

  dispose(): void {
    if (this.disposed) {
      return
    }
    this.disposed = true
    this.unsubscribeSurface()
    this.backendStatusListeners.clear()
    this.frameProofListeners.clear()
    this.rendererRuntime.setInputSignalListener(null)
    this.rendererRuntime.setFrameResultListener(null)
    const canvas = this.canvas
    this.canvas = null
    this.surfaceRuntime.dispose()
    this.rendererRuntime.dispose()
    if (canvas) {
      canvas.width = 0
      canvas.height = 0
    }
  }

  private applyRendererState(): void {
    this.rendererRuntime.updateState({
      active: this.props.active,
      backend: this.surfaceSnapshot.backend,
      cameraStore: this.props.cameraStore,
      drawText: this.props.drawText,
      frameProofSignature: this.frameProofSignature,
      geometry: this.props.geometry,
      headerPanes: this.props.headerPanes,
      overlay: this.props.overlay,
      overlayBuilder: this.props.overlayBuilder,
      preloadTilePanes: this.props.preloadTilePanes,
      renderRevisionSnapshot: this.props.renderRevisionSnapshot,
      scrollTransformStore: this.props.scrollTransformStore,
      surface: this.surfaceSnapshot.surface,
      tilePanes: this.props.tilePanes,
      visibleSceneOwnershipEpoch: this.visibleSceneOwnershipEpoch,
      visibleSceneOwnershipEpochSignature: this.visibleSceneOwnershipEpochSignature,
      visibleSceneOwnershipSignature: this.visibleSceneOwnershipSignature,
      webGpuReady: this.surfaceSnapshot.webGpuReady,
    })
  }

  private requestRenderDraw(): void {
    this.rendererRuntime.requestDraw()
  }

  private emitBackendStatus(): void {
    this.backendStatusListeners.forEach((listener) => listener())
  }

  private emitFrameProofStatus(): void {
    this.frameProofListeners.forEach((listener) => listener())
  }

  private handleFrameResult(result: WorkbookPaneFrameResultV3): void {
    const signature = result.frameProofSignature
    const sceneOwnershipSignature = result.visibleSceneOwnershipSignature
    const sceneOwnershipEpochSignature = result.visibleSceneOwnershipEpochSignature
    if (
      !result.submitted ||
      !signature ||
      signature !== this.frameProofSignature ||
      !sceneOwnershipSignature ||
      sceneOwnershipSignature !== this.visibleSceneOwnershipSignature ||
      !sceneOwnershipEpochSignature ||
      sceneOwnershipEpochSignature !== this.visibleSceneOwnershipEpochSignature ||
      !result.visualFrame
    ) {
      this.adoptRejectedFrameVisibleSceneProof(result)
      return
    }
    this.setPresentedVisualFrame(result.visualFrame)
    this.setPresentedFrameProofSignature(signature)
    this.setPresentedVisibleSceneOwnershipEpoch(result.visibleSceneOwnershipEpoch)
    this.setPresentedVisibleSceneOwnershipEpochSignature(sceneOwnershipEpochSignature)
    this.setPresentedVisibleSceneOwnershipSignature(sceneOwnershipSignature)
    this.setPresentedPayloadProof(this.payloadProof)
    this.setHasPresentedFrame(true)
    this.setFrameProofStatus('presented')
  }

  private adoptRejectedFrameVisibleSceneProof(result: WorkbookPaneFrameResultV3): void {
    if (!result.visibleSceneOwnershipEpoch || !result.visibleSceneOwnershipSignature || !result.visibleSceneOwnershipEpochSignature) {
      return
    }
    this.setVisibleSceneProof({
      ownershipEpoch: result.visibleSceneOwnershipEpoch,
      ownershipEpochSignature: result.visibleSceneOwnershipEpochSignature,
      ownershipSignature: result.visibleSceneOwnershipSignature,
      payload: result.visibleScenePayloadProof,
    })
    this.syncFrameProofSignature(this.props)
    this.applyRendererState()
    this.setHasPresentedFrame(false)
    this.setFrameProofStatus(this.frameProofSignature ? 'pending' : 'idle')
    this.requestRenderDraw()
  }

  private handleInputSignal(): void {
    if (this.disposed || !this.frameProofSignature || this.surfaceBackendStatus !== 'ready') {
      return
    }
    this.syncFrameProofSignature(this.props)
    this.applyRendererState()
    this.setHasPresentedFrame(false)
    this.setFrameProofStatus('pending')
  }

  private setFrameProofStatus(status: WorkbookPaneFrameProofStatusV3): void {
    if (this.frameProofStatus === status) {
      return
    }
    this.frameProofStatus = status
    this.emitFrameProofStatus()
  }

  private setHasPresentedFrame(value: boolean): void {
    if (this.hasPresentedFrame === value) {
      return
    }
    this.hasPresentedFrame = value
    this.emitFrameProofStatus()
  }

  private setPresentedFrameProofSignature(signature: string): void {
    if (this.presentedFrameProofSignature === signature) {
      return
    }
    this.presentedFrameProofSignature = signature
    this.emitFrameProofStatus()
  }

  private setPresentedPayloadProof(payload: WorkbookPaneVisiblePayloadProofV3): void {
    if (this.presentedPayloadProof === payload || areVisiblePayloadProofsEqual(this.presentedPayloadProof, payload)) {
      return
    }
    this.presentedPayloadProof = payload
    this.emitFrameProofStatus()
  }

  private setPresentedVisibleSceneOwnershipSignature(signature: string): void {
    if (this.presentedVisibleSceneOwnershipSignature === signature) {
      return
    }
    this.presentedVisibleSceneOwnershipSignature = signature
    this.emitFrameProofStatus()
  }

  private setPresentedVisualFrame(frame: WorkbookPanePresentedVisualFrameV3 | null): void {
    if (this.presentedVisualFrame === frame) {
      return
    }
    this.presentedVisualFrame = frame
    this.emitFrameProofStatus()
  }

  private syncFrameProofSignature(props: WorkbookPaneRendererHostPropsV3): void {
    const overlay = resolveWorkbookPaneHostOverlayProofV3(props)
    const geometry = resolveWorkbookPaneRendererGeometryV3({
      cameraStore: props.cameraStore,
      geometry: props.geometry,
    })
    const visibleSceneProof = resolveWorkbookPaneVisibleSceneProofV3({
      drawText: props.drawText,
      geometry,
      headerPanes: props.headerPanes,
      overlay,
      renderRevisionSnapshot: props.renderRevisionSnapshot ?? null,
      scrollSnapshot: resolveTypeGpuV3DrawScrollSnapshot({
        fallback: props.scrollTransformStore?.getSnapshot() ?? { tx: 0, ty: 0 },
        geometry,
        panes: props.tilePanes,
      }),
      surface: this.surfaceSnapshot.surface,
      tilePanes: props.tilePanes,
    })
    const visibleSceneChanged = this.setVisibleSceneProof(visibleSceneProof)
    const signature = resolveWorkbookPaneFrameProofSignatureV3({
      ...props,
      overlay,
      surface: this.surfaceSnapshot.surface,
    })
    if (this.frameProofSignature !== signature) {
      this.frameProofSignature = signature
      this.emitFrameProofStatus()
    }
    if (!signature) {
      this.setPresentedFrameProofSignature('')
      this.setPresentedVisibleSceneOwnershipEpoch(null)
      this.setPresentedVisibleSceneOwnershipEpochSignature('')
      this.setPresentedVisibleSceneOwnershipSignature('')
      this.setPresentedPayloadProof(EMPTY_VISIBLE_PAYLOAD_PROOF)
      this.setPresentedVisualFrame(null)
      this.setHasPresentedFrame(false)
      this.setFrameProofStatus('idle')
      return
    }
    const hasPresentedCurrentSignature =
      this.presentedFrameProofSignature === signature &&
      this.presentedVisibleSceneOwnershipEpochSignature === this.visibleSceneOwnershipEpochSignature &&
      this.presentedVisibleSceneOwnershipSignature === this.visibleSceneOwnershipSignature &&
      areVisiblePayloadProofsEqual(this.presentedPayloadProof, this.payloadProof)
    if (visibleSceneChanged && !hasPresentedCurrentSignature) {
      this.setHasPresentedFrame(false)
      this.setFrameProofStatus('pending')
      return
    }
    this.setHasPresentedFrame(hasPresentedCurrentSignature)
    this.setFrameProofStatus(hasPresentedCurrentSignature ? 'presented' : 'pending')
  }

  private setVisibleSceneProof(proof: {
    readonly ownershipEpoch: WorkbookPaneVisibleSceneOwnershipEpochV3
    readonly ownershipEpochSignature: string
    readonly ownershipSignature: string
    readonly payload: WorkbookPaneVisiblePayloadProofV3
  }): boolean {
    let changed = false
    if (this.visibleSceneOwnershipEpochSignature !== proof.ownershipEpochSignature) {
      this.visibleSceneOwnershipEpochSignature = proof.ownershipEpochSignature
      changed = true
    }
    if (!areVisibleSceneOwnershipEpochsEqual(this.visibleSceneOwnershipEpoch, proof.ownershipEpoch)) {
      this.visibleSceneOwnershipEpoch = proof.ownershipEpoch
      changed = true
    }
    if (this.visibleSceneOwnershipSignature !== proof.ownershipSignature) {
      this.visibleSceneOwnershipSignature = proof.ownershipSignature
      changed = true
    }
    if (!areVisiblePayloadProofsEqual(this.payloadProof, proof.payload)) {
      this.payloadProof = proof.payload
      changed = true
    }
    if (changed) {
      this.emitFrameProofStatus()
    }
    return changed
  }

  private setPresentedVisibleSceneOwnershipEpoch(epoch: WorkbookPaneVisibleSceneOwnershipEpochV3 | null): void {
    if (areVisibleSceneOwnershipEpochsEqual(this.presentedVisibleSceneOwnershipEpoch, epoch)) {
      return
    }
    this.presentedVisibleSceneOwnershipEpoch = epoch
    this.emitFrameProofStatus()
  }

  private setPresentedVisibleSceneOwnershipEpochSignature(signature: string): void {
    if (this.presentedVisibleSceneOwnershipEpochSignature === signature) {
      return
    }
    this.presentedVisibleSceneOwnershipEpochSignature = signature
    this.emitFrameProofStatus()
  }

  private syncCanvasTarget(): void {
    if (this.disposed) {
      return
    }
    this.surfaceRuntime.setCanvas(this.props.active && this.props.host ? this.canvas : null)
  }
}

function resolveWorkbookPaneHostOverlayProofV3(props: WorkbookPaneRendererHostPropsV3): DynamicGridOverlayBatchV3 | null {
  if (!props.overlayBuilder) {
    return props.overlay
  }
  const geometry = resolveWorkbookPaneRendererGeometryV3({
    cameraStore: props.cameraStore,
    geometry: props.geometry,
  })
  if (!geometry) {
    return props.overlay
  }
  return props.overlayBuilder(geometry) ?? null
}

export function resolveWorkbookPaneFrameProofSignatureV3(props: {
  readonly drawText?: boolean | undefined
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly renderRevisionSnapshot?: GridRenderRevisionSnapshot | null | undefined
  readonly surface?: TypeGpuSurfaceSizeV3 | null | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): string {
  const textOwnershipSignature = `drawText:${props.drawText === false ? 'gpu-text-off' : 'gpu-text-on'}`
  const surfaceSignature = props.surface
    ? ['surface', props.surface.width, props.surface.height, props.surface.pixelWidth, props.surface.pixelHeight, props.surface.dpr].join(
        ':',
      )
    : ''
  const renderRevisionSignature = props.renderRevisionSnapshot
    ? [
        props.renderRevisionSnapshot.authoritativeRevision ?? 'none',
        props.renderRevisionSnapshot.localRevision ?? 'none',
        props.renderRevisionSnapshot.projectedRevision,
        props.renderRevisionSnapshot.tileSceneCameraSeq ?? 'none',
        props.renderRevisionSnapshot.tileSceneRevision ?? 'none',
      ].join(':')
    : ''
  const tileSignature = props.tilePanes
    .map((pane) => {
      const tile = pane.tile
      return [
        pane.paneId,
        pane.generation,
        pane.drawVisible === false ? 'hidden' : 'visible',
        pane.frame.x,
        pane.frame.y,
        pane.frame.width,
        pane.frame.height,
        pane.contentOffset.x,
        pane.contentOffset.y,
        pane.surfaceSize.width,
        pane.surfaceSize.height,
        pane.viewport.rowStart,
        pane.viewport.rowEnd,
        pane.viewport.colStart,
        pane.viewport.colEnd,
        pane.scrollAxes.x ? 'scroll-x' : 'fixed-x',
        pane.scrollAxes.y ? 'scroll-y' : 'fixed-y',
        tile.tileId,
        tile.textCount,
        tile.textSignature ?? resolveGridTextTileRevisionKeyV3(tile).textSignature,
        tile.rectCount,
        tile.rectSignature ?? '',
        tile.version.axisX,
        tile.version.axisY,
        tile.version.freeze,
        tile.version.styles,
        tile.version.text,
        tile.version.values,
      ].join(':')
    })
    .join('|')
  const headerSignature = props.headerPanes
    .map((pane) =>
      [
        pane.paneId,
        pane.frame.x,
        pane.frame.y,
        pane.frame.width,
        pane.frame.height,
        pane.contentOffset.x,
        pane.contentOffset.y,
        pane.surfaceSize.width,
        pane.surfaceSize.height,
        pane.scrollAxes.x ? 'scroll-x' : 'fixed-x',
        pane.scrollAxes.y ? 'scroll-y' : 'fixed-y',
        pane.rectSignature,
        pane.textSignature,
        pane.rectCount,
        pane.textCount,
      ].join(':'),
    )
    .join('|')
  const overlaySignature = props.overlay
    ? [
        props.overlay.sheetName,
        props.overlay.seq,
        props.overlay.cameraSeq,
        props.overlay.surfaceSize.width,
        props.overlay.surfaceSize.height,
        props.overlay.rectCount,
        props.overlay.fillRectCount,
        props.overlay.borderRectCount,
        props.overlay.rectSignature,
      ].join(':')
    : ''
  return [textOwnershipSignature, surfaceSignature, tileSignature, headerSignature, overlaySignature, renderRevisionSignature]
    .filter(Boolean)
    .join('#')
}

function areVisiblePayloadProofsEqual(left: WorkbookPaneVisiblePayloadProofV3, right: WorkbookPaneVisiblePayloadProofV3): boolean {
  return (
    left.contentSignature === right.contentSignature &&
    left.rectCount === right.rectCount &&
    left.rectSignature === right.rectSignature &&
    left.textRunCount === right.textRunCount &&
    left.textSignature === right.textSignature
  )
}

function areVisibleSceneOwnershipEpochsEqual(
  left: WorkbookPaneVisibleSceneOwnershipEpochV3 | null,
  right: WorkbookPaneVisibleSceneOwnershipEpochV3 | null,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  return (
    left.fillHandleRevision === right.fillHandleRevision &&
    left.sceneEpoch === right.sceneEpoch &&
    left.selectionRevision === right.selectionRevision &&
    left.semanticMutationRevision === right.semanticMutationRevision &&
    left.viewportRevision === right.viewportRevision &&
    left.workbookRevision === right.workbookRevision
  )
}
