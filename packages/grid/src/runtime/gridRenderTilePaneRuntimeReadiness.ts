import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import type { TileKey53 } from '../renderer-v3/tile-key.js'
import type { GridRuntimeHost } from './gridRuntimeHost.js'
import type { GridTileReadinessSnapshotV3 } from './gridTileCoordinator.js'
import type { GridEngineLike } from '../grid-engine.js'
import { tileSatisfiesRequiredProjectedRevision } from './gridRenderTileRevision.js'

interface DirtyPaneTileKeyInput {
  readonly engine: GridEngineLike
  readonly gridRuntimeHost: GridRuntimeHost
}

export function resolveAcknowledgedVisibleDirtyPaneTileKeys(
  panes: readonly WorkbookRenderTilePaneState[],
  input: DirtyPaneTileKeyInput,
): readonly TileKey53[] {
  const keys: number[] = []
  const seen = new Set<number>()
  for (const pane of panes) {
    if (pane.drawVisible === false || (pane.tile.dirtyMasks?.length ?? 0) === 0 || seen.has(pane.tile.tileId)) {
      continue
    }
    const requiredProjectedRevision = input.gridRuntimeHost.tiles.dirtyTiles.getRequiredProjectedRevision(pane.tile.tileId)
    if (!tileSatisfiesRequiredProjectedRevision(pane.tile, input.engine, requiredProjectedRevision)) {
      continue
    }
    seen.add(pane.tile.tileId)
    keys.push(pane.tile.tileId)
  }
  return keys
}

export function hasUnacknowledgedVisibleDirtyTileKeys(
  readiness: GridTileReadinessSnapshotV3,
  acknowledgedVisibleTileKeys: readonly TileKey53[],
): boolean {
  if (readiness.visibleDirtyTileKeys.length === 0) {
    return false
  }
  const acknowledged = new Set(acknowledgedVisibleTileKeys)
  return readiness.visibleDirtyTileKeys.some((key) => !acknowledged.has(key))
}
