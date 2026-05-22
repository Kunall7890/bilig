export type WorkbookVisibleRenderBackendStatus = 'idle' | 'initializing' | 'ready' | 'unavailable'
export type WorkbookVisibleRenderFrameProofStatus = 'idle' | 'pending' | 'presented'

export interface WorkbookVisibleRenderProof {
  readonly mode: 'typegpu-v3' | 'typegpu-v3-unavailable'
  readonly backendStatus: WorkbookVisibleRenderBackendStatus
  readonly frameProofStatus: WorkbookVisibleRenderFrameProofStatus
  readonly hasPresentedFrame: boolean
  readonly hasPresentedVisibleFrame: boolean
  readonly frameProofSignature: string
  readonly presentedFrameProofSignature: string
  readonly currentSceneOwnershipSignature: string
  readonly presentedSceneOwnershipSignature: string
  readonly authoritativeRevision: number | null
  readonly localRevision: number | null
  readonly projectedRevision: number | null
  readonly visibleRenderRevision: number | null
  readonly tileSceneRevision: number | null
  readonly tileSceneCameraSeq: number | null
  readonly currentTilePaneCount: number
  readonly currentHeaderPaneCount: number
  readonly presentedTilePaneCount: number
  readonly presentedHeaderPaneCount: number
  readonly surfaceWidth: number
  readonly surfaceHeight: number
  readonly surfacePixelWidth: number
  readonly surfacePixelHeight: number
  readonly devicePixelRatio: number
  readonly capturedAtUnixMs: number
}
