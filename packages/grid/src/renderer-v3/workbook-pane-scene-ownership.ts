import type { GridRenderRevisionSnapshot } from '../grid-engine.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import type { TypeGpuSurfaceSizeV3 } from './workbook-pane-renderer-runtime.js'

export interface WorkbookPaneSceneOwnershipInputV3 {
  readonly drawText?: boolean | undefined
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly renderRevisionSnapshot?: GridRenderRevisionSnapshot | null | undefined
  readonly surface?: TypeGpuSurfaceSizeV3 | null | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

export function resolveWorkbookPaneSceneOwnershipSignatureV3(input: WorkbookPaneSceneOwnershipInputV3): string {
  return [
    `drawText:${input.drawText === false ? 'off' : 'on'}`,
    resolveSurfaceOwnershipSignature(input.surface),
    resolveRevisionOwnershipSignature(input.renderRevisionSnapshot),
    resolveHeaderOwnershipSignature(input.headerPanes),
    resolveTileOwnershipSignature(input.tilePanes),
    resolveOverlayOwnershipSignature(input.overlay),
  ]
    .filter(Boolean)
    .join('#')
}

function resolveSurfaceOwnershipSignature(surface: TypeGpuSurfaceSizeV3 | null | undefined): string {
  if (!surface) {
    return ''
  }
  return ['surface', surface.width, surface.height, surface.pixelWidth, surface.pixelHeight, surface.dpr].join(':')
}

function resolveRevisionOwnershipSignature(revision: GridRenderRevisionSnapshot | null | undefined): string {
  if (!revision) {
    return ''
  }
  return [
    'revision',
    revision.authoritativeRevision ?? 'none',
    revision.localRevision ?? 'none',
    revision.projectedRevision,
    revision.tileSceneRevision ?? 'none',
    revision.tileSceneCameraSeq ?? 'none',
  ].join(':')
}

function resolveHeaderOwnershipSignature(headerPanes: readonly GridHeaderPaneState[]): string {
  if (headerPanes.length === 0) {
    return ''
  }
  return headerPanes
    .map((pane) =>
      [
        'header',
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
}

function resolveTileOwnershipSignature(tilePanes: readonly WorkbookRenderTilePaneState[]): string {
  if (tilePanes.length === 0) {
    return ''
  }
  return tilePanes
    .map((pane) => {
      const tile = pane.tile
      return [
        'tile',
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
        tile.coord.sheetId,
        tile.coord.sheetOrdinal,
        tile.coord.paneKind,
        tile.coord.rowTile,
        tile.coord.colTile,
        tile.coord.dprBucket,
        tile.bounds.rowStart,
        tile.bounds.rowEnd,
        tile.bounds.colStart,
        tile.bounds.colEnd,
        tile.lastBatchId,
        tile.lastCameraSeq,
        tile.version.axisX,
        tile.version.axisY,
        tile.version.freeze,
        tile.version.styles,
        tile.version.text,
        tile.version.values,
      ].join(':')
    })
    .join('|')
}

function resolveOverlayOwnershipSignature(overlay: DynamicGridOverlayBatchV3 | null): string {
  if (!overlay) {
    return ''
  }
  return [
    'overlay',
    overlay.sheetName,
    overlay.seq,
    overlay.cameraSeq,
    overlay.surfaceSize.width,
    overlay.surfaceSize.height,
    overlay.rectCount,
    overlay.fillRectCount,
    overlay.borderRectCount,
    overlay.rectSignature,
  ].join(':')
}
