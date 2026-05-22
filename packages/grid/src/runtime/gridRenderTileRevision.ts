import type { GridEngineLike } from '../grid-engine.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import type { GridRenderTilePaneRuntimeInput } from './gridRenderTilePaneRuntime.js'

export function normalizeNonNegativeInteger(value: number | null | undefined): number | null {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : null
}

export function tileProjectionRevisionIsBehind(tile: GridRenderTile | null, engine: GridEngineLike): boolean {
  if (!tile) {
    return false
  }
  const renderRevision = engine.getRenderRevisionSnapshot?.()
  const projectedRevision = normalizeNonNegativeInteger(renderRevision?.projectedRevision)
  if (projectedRevision === null) {
    return false
  }
  const authoritativeRevision = normalizeNonNegativeInteger(renderRevision?.authoritativeRevision)
  const localRevision = normalizeNonNegativeInteger(renderRevision?.localRevision)
  const hasPendingLocalProjection =
    localRevision !== null && localRevision > 0 && (authoritativeRevision === null || projectedRevision > authoritativeRevision)
  return tile.lastBatchId < projectedRevision || hasPendingLocalProjection
}

export function resolveLocalRenderGeneration(input: GridRenderTilePaneRuntimeInput): number {
  const renderRevision = input.engine.getRenderRevisionSnapshot?.()
  return Math.max(
    input.sceneRevision,
    normalizeNonNegativeInteger(renderRevision?.localRevision) ?? 0,
    normalizeNonNegativeInteger(renderRevision?.projectedRevision) ?? 0,
  )
}

export function tileSatisfiesRequiredProjectedRevision(
  tile: GridRenderTile,
  engine: GridEngineLike,
  requiredProjectedRevision: number | null,
): boolean {
  if (requiredProjectedRevision === null) {
    return true
  }
  const renderRevision = engine.getRenderRevisionSnapshot?.()
  if (!renderRevision) {
    return true
  }
  const projectedRevision = normalizeNonNegativeInteger(renderRevision.projectedRevision)
  if (projectedRevision === null || projectedRevision < requiredProjectedRevision) {
    return false
  }
  return (
    tile.version.styles >= requiredProjectedRevision &&
    tile.version.text >= requiredProjectedRevision &&
    tile.version.values >= requiredProjectedRevision
  )
}
