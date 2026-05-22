import type { GridRenderRevisionSnapshot } from '../grid-engine.js'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { WorkbookGridScrollSnapshot } from '../workbookGridScrollStore.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { resolveGridTextTileRevisionKeyV3 } from './typegpu-tile-resource-revisions.js'
import type { TypeGpuSurfaceSizeV3 } from './workbook-pane-renderer-runtime.js'

export interface WorkbookPaneVisiblePayloadProofV3 {
  readonly contentSignature: string
  readonly rectCount: number
  readonly rectSignature: string
  readonly textRunCount: number
  readonly textSignature: string
}

export interface WorkbookPaneVisibleSceneProofV3 {
  readonly ownershipSignature: string
  readonly payload: WorkbookPaneVisiblePayloadProofV3
}

export function resolveWorkbookPaneVisiblePayloadProofV3(input: {
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): WorkbookPaneVisiblePayloadProofV3 {
  const textRunCount =
    input.headerPanes.reduce((total, pane) => total + pane.textCount, 0) +
    input.tilePanes.reduce((total, pane) => total + pane.tile.textCount, 0)
  const rectCount =
    input.headerPanes.reduce((total, pane) => total + pane.rectCount, 0) +
    input.tilePanes.reduce((total, pane) => total + pane.tile.rectCount, 0) +
    (input.overlay?.rectCount ?? 0)
  const textSignature = [
    ...input.headerPanes.map((pane) => ['header-text', pane.paneId, pane.textCount, pane.textSignature].join(':')),
    ...input.tilePanes.map((pane) =>
      [
        'tile-text',
        pane.paneId,
        pane.tile.tileId,
        pane.tile.textCount,
        pane.tile.textSignature ?? resolveGridTextTileRevisionKeyV3(pane.tile).textSignature,
        pane.tile.lastBatchId,
        pane.tile.lastCameraSeq,
        pane.tile.version.axisX,
        pane.tile.version.axisY,
        pane.tile.version.freeze,
        pane.tile.version.styles,
        pane.tile.version.text,
        pane.tile.version.values,
      ].join(':'),
    ),
  ].join('|')
  const rectSignature = [
    ...input.headerPanes.map((pane) => ['header-rect', pane.paneId, pane.rectCount, pane.rectSignature].join(':')),
    ...input.tilePanes.map((pane) =>
      [
        'tile-rect',
        pane.paneId,
        pane.tile.tileId,
        pane.tile.rectCount,
        pane.tile.rectSignature ?? '',
        pane.tile.lastBatchId,
        pane.tile.lastCameraSeq,
        pane.tile.version.axisX,
        pane.tile.version.axisY,
        pane.tile.version.freeze,
        pane.tile.version.styles,
        pane.tile.version.text,
        pane.tile.version.values,
      ].join(':'),
    ),
    input.overlay
      ? ['overlay-rect', input.overlay.sheetName, input.overlay.seq, input.overlay.rectCount, input.overlay.rectSignature].join(':')
      : '',
  ]
    .filter(Boolean)
    .join('|')
  return {
    contentSignature: ['content-v3', textRunCount, textSignature, rectCount, rectSignature].join('#'),
    rectCount,
    rectSignature: rectSignature || 'rects:none',
    textRunCount,
    textSignature: textSignature || 'text:none',
  }
}

export function resolveWorkbookPaneVisibleSceneProofV3(input: {
  readonly drawText: boolean
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly renderRevisionSnapshot: GridRenderRevisionSnapshot | null
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuSurfaceSizeV3 | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): WorkbookPaneVisibleSceneProofV3 {
  const payload = resolveWorkbookPaneVisiblePayloadProofV3({
    headerPanes: input.headerPanes,
    overlay: input.overlay,
    tilePanes: input.tilePanes,
  })
  const surfaceSignature = input.surface
    ? ['surface', input.surface.width, input.surface.height, input.surface.pixelWidth, input.surface.pixelHeight, input.surface.dpr].join(
        ':',
      )
    : 'surface:none'
  const renderRevisionSignature = input.renderRevisionSnapshot
    ? [
        'revision',
        input.renderRevisionSnapshot.authoritativeRevision ?? 'none',
        input.renderRevisionSnapshot.localRevision ?? 'none',
        input.renderRevisionSnapshot.projectedRevision,
        input.renderRevisionSnapshot.tileSceneCameraSeq ?? 'none',
        input.renderRevisionSnapshot.tileSceneRevision ?? 'none',
      ].join(':')
    : 'revision:none'
  const cameraSignature = input.geometry
    ? [
        'camera',
        input.geometry.camera.sheetName,
        input.geometry.camera.seq,
        input.geometry.camera.bodyScrollX,
        input.geometry.camera.bodyScrollY,
        input.geometry.camera.bodyWorldX,
        input.geometry.camera.bodyWorldY,
        input.geometry.camera.frozenWidth,
        input.geometry.camera.frozenHeight,
      ].join(':')
    : 'camera:none'
  const scrollSignature = [
    'scroll',
    input.scrollSnapshot.tx ?? 'none',
    input.scrollSnapshot.ty ?? 'none',
    input.scrollSnapshot.renderTx ?? 'none',
    input.scrollSnapshot.renderTy ?? 'none',
    input.scrollSnapshot.scrollLeft ?? 'none',
    input.scrollSnapshot.scrollTop ?? 'none',
  ].join(':')
  const tileOwnershipSignature = input.tilePanes
    .map((pane) =>
      [
        'tile',
        pane.paneId,
        pane.generation,
        pane.drawVisible === false ? 'hidden' : 'visible',
        pane.viewport.rowStart,
        pane.viewport.rowEnd,
        pane.viewport.colStart,
        pane.viewport.colEnd,
        pane.frame.x,
        pane.frame.y,
        pane.frame.width,
        pane.frame.height,
        pane.contentOffset.x,
        pane.contentOffset.y,
        pane.surfaceSize.width,
        pane.surfaceSize.height,
        pane.tile.tileId,
        pane.tile.coord.sheetId,
        pane.tile.coord.sheetOrdinal,
        pane.tile.coord.paneKind,
        pane.tile.coord.rowTile,
        pane.tile.coord.colTile,
        pane.tile.lastBatchId,
        pane.tile.lastCameraSeq,
        pane.tile.version.axisX,
        pane.tile.version.axisY,
        pane.tile.version.freeze,
        pane.tile.version.styles,
        pane.tile.version.text,
        pane.tile.version.values,
      ].join(':'),
    )
    .join('|')
  const headerOwnershipSignature = input.headerPanes
    .map((pane) =>
      [
        'header',
        pane.paneId,
        pane.drawVisible === false ? 'hidden' : 'visible',
        pane.frame.x,
        pane.frame.y,
        pane.frame.width,
        pane.frame.height,
        pane.contentOffset.x,
        pane.contentOffset.y,
        pane.surfaceSize.width,
        pane.surfaceSize.height,
        pane.rectCount,
        pane.textCount,
        pane.rectSignature,
        pane.textSignature,
      ].join(':'),
    )
    .join('|')
  const overlaySignature = input.overlay
    ? [
        'overlay',
        input.overlay.sheetName,
        input.overlay.seq,
        input.overlay.cameraSeq,
        input.overlay.surfaceSize.width,
        input.overlay.surfaceSize.height,
        input.overlay.rectCount,
        input.overlay.fillRectCount,
        input.overlay.borderRectCount,
        input.overlay.rectSignature,
      ].join(':')
    : 'overlay:none'
  return {
    ownershipSignature: [
      'visible-scene-v3',
      input.drawText ? 'gpu-text-on' : 'gpu-text-off',
      surfaceSignature,
      renderRevisionSignature,
      cameraSignature,
      scrollSignature,
      payload.contentSignature,
      tileOwnershipSignature,
      headerOwnershipSignature,
      overlaySignature,
    ].join('#'),
    payload,
  }
}
