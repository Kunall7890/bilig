import { formatAddress } from '@bilig/formula'
import type { Viewport } from '@bilig/protocol'

import type { GridEngineLike } from '../grid-engine.js'
import { snapshotToRenderCell } from '../gridCells.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

interface VisibleTextRefreshCacheInput {
  readonly engine: GridEngineLike
  readonly sceneRevision: number
  readonly sheetName: string
  readonly visibleViewport: Viewport
}

interface VisibleTextRefreshCacheEntry extends VisibleTextRefreshCacheInput {
  readonly needsLocalRefresh: boolean
  readonly renderRevisionKey: string
  readonly tile: GridRenderTile
  readonly visibleColEnd: number
  readonly visibleColStart: number
  readonly visibleRowEnd: number
  readonly visibleRowStart: number
}

export class GridVisibleTextRefreshCache {
  private readonly entries = new Map<number, VisibleTextRefreshCacheEntry>()

  needsLocalRefresh(tileKey: number, tile: GridRenderTile | null, input: VisibleTextRefreshCacheInput): boolean {
    if (!tile) {
      this.entries.delete(tileKey)
      return false
    }
    const visibleRowStart = Math.max(tile.bounds.rowStart, input.visibleViewport.rowStart)
    const visibleRowEnd = Math.min(tile.bounds.rowEnd, input.visibleViewport.rowEnd)
    const visibleColStart = Math.max(tile.bounds.colStart, input.visibleViewport.colStart)
    const visibleColEnd = Math.min(tile.bounds.colEnd, input.visibleViewport.colEnd)
    if (visibleRowStart > visibleRowEnd || visibleColStart > visibleColEnd) {
      this.entries.delete(tileKey)
      return false
    }

    const renderRevisionKey = resolveRenderRevisionKey(input.engine)
    const cached = this.entries.get(tileKey)
    if (
      cached &&
      cached.tile === tile &&
      cached.engine === input.engine &&
      cached.sheetName === input.sheetName &&
      cached.sceneRevision === input.sceneRevision &&
      cached.renderRevisionKey === renderRevisionKey &&
      cached.visibleRowStart === visibleRowStart &&
      cached.visibleRowEnd === visibleRowEnd &&
      cached.visibleColStart === visibleColStart &&
      cached.visibleColEnd === visibleColEnd
    ) {
      return cached.needsLocalRefresh
    }

    const needsLocalRefresh = tileVisibleTextNeedsLocalRefresh(tile, input, {
      visibleColEnd,
      visibleColStart,
      visibleRowEnd,
      visibleRowStart,
    })
    this.entries.set(tileKey, {
      ...input,
      needsLocalRefresh,
      renderRevisionKey,
      tile,
      visibleColEnd,
      visibleColStart,
      visibleRowEnd,
      visibleRowStart,
    })
    return needsLocalRefresh
  }
}

function tileVisibleTextNeedsLocalRefresh(
  tile: GridRenderTile,
  input: Pick<VisibleTextRefreshCacheInput, 'engine' | 'sheetName'>,
  visibleBounds: {
    readonly visibleColEnd: number
    readonly visibleColStart: number
    readonly visibleRowEnd: number
    readonly visibleRowStart: number
  },
): boolean {
  if (tile.textRuns.length !== tile.textCount) {
    return true
  }

  const textRunsByCell = new Map<string, string>()
  for (const run of tile.textRuns) {
    if (run.text.length === 0) {
      continue
    }
    const row = run.row
    const col = run.col
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(col) ||
      row === undefined ||
      col === undefined ||
      row < tile.bounds.rowStart ||
      row > tile.bounds.rowEnd ||
      col < tile.bounds.colStart ||
      col > tile.bounds.colEnd
    ) {
      return true
    }
    if (
      row >= visibleBounds.visibleRowStart &&
      row <= visibleBounds.visibleRowEnd &&
      col >= visibleBounds.visibleColStart &&
      col <= visibleBounds.visibleColEnd
    ) {
      const key = `${row}:${col}`
      if (textRunsByCell.has(key)) {
        return true
      }
      textRunsByCell.set(key, run.text)
    }
  }

  for (let row = visibleBounds.visibleRowStart; row <= visibleBounds.visibleRowEnd; row += 1) {
    for (let col = visibleBounds.visibleColStart; col <= visibleBounds.visibleColEnd; col += 1) {
      const snapshot = input.engine.getCell(input.sheetName, formatAddress(row, col))
      const renderCell = snapshotToRenderCell(snapshot, input.engine.getCellStyle(snapshot.styleId))
      const visibleTileText = textRunsByCell.get(`${row}:${col}`) ?? ''
      if (visibleTileText !== renderCell.displayText) {
        return true
      }
    }
  }
  return false
}

function resolveRenderRevisionKey(engine: GridEngineLike): string {
  const revision = engine.getRenderRevisionSnapshot?.()
  if (!revision) {
    return 'untracked'
  }
  return [
    revision.authoritativeRevision ?? 'none',
    revision.localRevision ?? 'none',
    revision.projectedRevision,
    revision.tileSceneRevision ?? 'none',
    revision.tileSceneCameraSeq ?? 'none',
  ].join(':')
}
