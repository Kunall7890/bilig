import { VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, ValueTag, type CellSnapshot, type Viewport } from '@bilig/protocol'
import { indexToColumn } from '@bilig/formula'
import { buildGridGpuScene } from '../gridGpuScene.js'
import { getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset, type GridMetrics } from '../gridMetrics.js'
import { buildGridTextScene } from '../gridTextScene.js'
import type { GridEngineLike } from '../grid-engine.js'
import { snapshotToRenderCell } from '../gridCells.js'
import { CompactSelection, type GridSelection, type Item, type Rectangle } from '../gridTypes.js'
import { collectViewportItems } from '../gridViewportItems.js'
import { resolveColumnOffset } from '../workbookGridViewport.js'
import { packGridRectBufferV3 } from './rect-instance-buffer.js'
import type { GridRenderTile } from './render-tile-source.js'
import { resolveGridRenderTileDirtySpansV3 } from './render-tile-dirty-spans.js'
import { createGridTilePacketV3 } from './tile-packet-v3.js'
import { packTileKey53 } from './tile-key.js'
import { packGridTextBufferV3, packGridTextRunsBufferV3, type PackedGridTextBufferV3 } from './text-run-buffer.js'
import { DirtyMaskV3 } from './tile-damage-index.js'

export interface GridTileMaterializerAxisInputV3 {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
}

export interface MaterializeGridRenderTileInputV3 extends GridTileMaterializerAxisInputV3 {
  readonly engine: GridEngineLike
  readonly editingCell?: Item | null | undefined
  readonly sheetName: string
  readonly sheetId: number
  readonly sheetOrdinal: number
  readonly viewport: Viewport
  readonly dprBucket: number
  readonly packetSeq: number
  readonly materializedAtSeq: number
  readonly cameraSeq: number
  readonly valueSeq: number
  readonly styleSeq: number
  readonly textSeq: number
  readonly rectSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly glyphAtlasSeq?: number | undefined
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
  readonly reuseStaticGridRectsFrom?: GridRenderTile | undefined
  readonly reuseTextRunsFrom?: GridRenderTile | undefined
  readonly selectedCell?: Item | undefined
  readonly selectedCellSnapshot?: CellSnapshot | null | undefined
}

const STATIC_TILE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_TILE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})

export function materializeGridRenderTileV3(input: MaterializeGridRenderTileInputV3): GridRenderTile {
  const rowTile = Math.floor(input.viewport.rowStart / VIEWPORT_TILE_ROW_COUNT)
  const colTile = Math.floor(input.viewport.colStart / VIEWPORT_TILE_COLUMN_COUNT)
  const tileId = packTileKey53({
    colTile,
    dprBucket: input.dprBucket,
    rowTile,
    sheetOrdinal: input.sheetOrdinal,
  })
  const surfaceSize = resolveTileSurfaceSizeV3(input.viewport, input)
  const visibleItems = collectViewportItems(input.viewport)
  const inboundTextSourceItems = collectInboundTextSpillSourceItemsV3(input)
  const visibleRegion = {
    range: {
      x: input.viewport.colStart,
      y: input.viewport.rowStart,
      width: input.viewport.colEnd - input.viewport.colStart + 1,
      height: input.viewport.rowEnd - input.viewport.rowStart + 1,
    },
    tx: 0,
    ty: 0,
    freezeRows: 0,
    freezeCols: 0,
  }
  const getCellBounds = createTileCellBoundsResolverV3(input)
  const getTextCellBounds = createTileCellBoundsResolverV3({
    ...input,
    minCol: inboundTextSourceItems.reduce((min, [col]) => Math.min(min, col), input.viewport.colStart),
  })
  const selectedCell = input.selectedCell ?? STATIC_TILE_SELECTED_CELL
  const selectedCellSnapshot = input.selectedCellSnapshot ?? null
  const reuseRectTile = canReuseStaticGridRectsForTile(input, tileId) ? input.reuseStaticGridRectsFrom : undefined
  const reuseTextTile = canReuseTextRunsForTile(input, tileId) ? input.reuseTextRunsFrom : undefined
  const patchedTextBuffer = reuseTextTile ? null : tryPatchTextRunsForTile(input, tileId, getTextCellBounds, visibleRegion, surfaceSize)
  const rectBuffer = reuseRectTile
    ? {
        rectCount: reuseRectTile.rectCount,
        rectInstances: reuseRectTile.rectInstances,
        rectSignature: reuseRectTile.rectSignature,
      }
    : packGridRectBufferV3(
        buildGridGpuScene({
          activeHeaderDrag: null,
          columnWidths: input.columnWidths,
          contentMode: 'data',
          engine: input.engine,
          getCellBounds,
          gridMetrics: input.gridMetrics,
          gridSelection: STATIC_TILE_GRID_SELECTION,
          hostBounds: { left: 0, top: 0 },
          hoveredCell: null,
          hoveredHeader: null,
          includeLeadingGridLines: false,
          resizeGuideColumn: null,
          resizeGuideRow: null,
          rowHeights: input.rowHeights,
          selectedCell,
          selectionRange: null,
          sheetName: input.sheetName,
          visibleItems,
          visibleRegion,
        }),
        surfaceSize,
      )
  const textBuffer = reuseTextTile
    ? {
        textCount: reuseTextTile.textCount,
        textMetrics: reuseTextTile.textMetrics,
        textRuns: reuseTextTile.textRuns,
        textSignature: reuseTextTile.textSignature ?? '',
      }
    : patchedTextBuffer
      ? patchedTextBuffer
      : packGridTextBufferV3(
          buildGridTextScene({
            activeHeaderDrag: null,
            columnWidths: input.columnWidths,
            contentMode: 'data',
            editingCell: input.editingCell ?? null,
            engine: input.engine,
            getCellBounds: getTextCellBounds,
            gridMetrics: input.gridMetrics,
            hostBounds: {
              height: surfaceSize.height,
              left: 0,
              top: 0,
              width: surfaceSize.width,
            },
            hoveredHeader: null,
            resizeGuideColumn: null,
            rowHeights: input.rowHeights,
            selectedCell,
            selectedCellSnapshot,
            selectionRange: null,
            sheetName: input.sheetName,
            visibleItems: inboundTextSourceItems.length > 0 ? [...visibleItems, ...inboundTextSourceItems] : visibleItems,
            visibleRegion,
          }),
        )
  const packet = createGridTilePacketV3({
    axisSeqX: input.axisSeqX,
    axisSeqY: input.axisSeqY,
    cellCount: visibleItems.length,
    dirtyLocalCols: input.dirtyLocalCols,
    dirtyLocalRows: input.dirtyLocalRows,
    dirtyMasks: input.dirtyMasks,
    freezeSeq: input.freezeSeq,
    glyphAtlasSeq: input.glyphAtlasSeq ?? 0,
    materializedAtSeq: input.materializedAtSeq,
    packetSeq: input.packetSeq,
    rectInstanceCount: rectBuffer.rectCount,
    rectInstances: rectBuffer.rectInstances,
    rectSeq: input.rectSeq,
    sheetId: input.sheetId,
    styleSeq: input.styleSeq,
    textRunCount: textBuffer.textCount,
    textRuns: textBuffer.textMetrics,
    textSeq: input.textSeq,
    tileKey: tileId,
    valueSeq: input.valueSeq,
  })
  const tile: GridRenderTile = {
    bounds: {
      colEnd: packet.colEnd,
      colStart: packet.colStart,
      rowEnd: packet.rowEnd,
      rowStart: packet.rowStart,
    },
    coord: {
      colTile,
      dprBucket: input.dprBucket,
      paneKind: 'body',
      rowTile,
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetOrdinal,
    },
    lastBatchId: packet.packetSeq,
    lastCameraSeq: input.cameraSeq,
    packet,
    rectCount: rectBuffer.rectCount,
    rectInstances: rectBuffer.rectInstances,
    rectSignature: rectBuffer.rectSignature,
    textCount: textBuffer.textCount,
    textMetrics: textBuffer.textMetrics,
    textRuns: textBuffer.textRuns,
    textSignature: textBuffer.textSignature,
    dirty: {
      glyphSpans: [],
      rectSpans: [],
      textSpans: [],
    },
    dirtyLocalCols: packet.dirtyLocalCols,
    dirtyLocalRows: packet.dirtyLocalRows,
    dirtyMasks: packet.dirtyMasks,
    tileId,
    version: {
      axisX: packet.axisSeqX,
      axisY: packet.axisSeqY,
      freeze: packet.freezeSeq,
      styles: packet.styleSeq,
      text: packet.textSeq,
      values: packet.valueSeq,
    },
  }
  return {
    ...tile,
    dirty: resolveGridRenderTileDirtySpansV3(tile),
  }
}

function canReuseStaticGridRectsForTile(input: MaterializeGridRenderTileInputV3, tileId: number): boolean {
  const baseTile = input.reuseStaticGridRectsFrom
  if (!baseTile || baseTile.tileId !== tileId) {
    return false
  }
  if (
    baseTile.bounds.rowStart !== input.viewport.rowStart ||
    baseTile.bounds.rowEnd !== input.viewport.rowEnd ||
    baseTile.bounds.colStart !== input.viewport.colStart ||
    baseTile.bounds.colEnd !== input.viewport.colEnd
  ) {
    return false
  }
  const dirtyMasks = input.dirtyMasks
  if (!dirtyMasks || dirtyMasks.length === 0) {
    return false
  }
  const rectDirtyMask = DirtyMaskV3.Rect | DirtyMaskV3.Border | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
  for (const mask of dirtyMasks) {
    if ((mask & rectDirtyMask) !== 0) {
      return false
    }
  }
  return true
}

function tryPatchTextRunsForTile(
  input: MaterializeGridRenderTileInputV3,
  tileId: number,
  getCellBounds: (col: number, row: number) => Rectangle | undefined,
  visibleRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows: number
    readonly freezeCols: number
  },
  surfaceSize: { readonly width: number; readonly height: number },
): PackedGridTextBufferV3 | null {
  const baseTile = input.reuseTextRunsFrom
  if (!baseTile || baseTile.tileId !== tileId || !sameTileBounds(baseTile, input.viewport)) {
    return null
  }
  if (input.editingCell !== undefined && input.editingCell !== null) {
    return null
  }
  if (input.selectedCellSnapshot !== undefined && input.selectedCellSnapshot !== null) {
    return null
  }
  if (baseTile.textRuns.length !== baseTile.textCount) {
    return null
  }
  const dirtyItems = collectPatchableTextDirtyItems(input)
  if (!dirtyItems || dirtyItems.length === 0) {
    return null
  }
  const dirtyKeys = new Set(dirtyItems.map(([col, row]) => textRunCellKey(row, col)))
  for (const run of baseTile.textRuns) {
    if (run.row === undefined || run.col === undefined) {
      return null
    }
  }
  const dirtyTextBuffer = packGridTextBufferV3(
    buildGridTextScene({
      activeHeaderDrag: null,
      columnWidths: input.columnWidths,
      contentMode: 'data',
      editingCell: null,
      engine: input.engine,
      getCellBounds,
      gridMetrics: input.gridMetrics,
      hostBounds: {
        height: surfaceSize.height,
        left: 0,
        top: 0,
        width: surfaceSize.width,
      },
      hoveredHeader: null,
      resizeGuideColumn: null,
      rowHeights: input.rowHeights,
      selectedCell: STATIC_TILE_SELECTED_CELL,
      selectedCellSnapshot: null,
      selectionRange: null,
      sheetName: input.sheetName,
      visibleItems: dirtyItems,
      visibleRegion,
    }),
  )
  const replacementRuns = new Map<string, (typeof baseTile.textRuns)[number]>()
  for (const run of dirtyTextBuffer.textRuns) {
    if (run.row === undefined || run.col === undefined) {
      return null
    }
    const key = textRunCellKey(run.row, run.col)
    if (!dirtyKeys.has(key) || replacementRuns.has(key)) {
      return null
    }
    replacementRuns.set(key, run)
  }
  const consumedReplacementKeys = new Set<string>()
  const patchedRuns = baseTile.textRuns.flatMap((run) => {
    const key = textRunCellKey(run.row!, run.col!)
    if (!dirtyKeys.has(key)) {
      return [run]
    }
    const replacement = replacementRuns.get(key)
    if (!replacement) {
      return []
    }
    consumedReplacementKeys.add(key)
    return [replacement]
  })
  dirtyItems.forEach(([col, row]) => {
    const key = textRunCellKey(row, col)
    const replacement = replacementRuns.get(key)
    if (replacement && !consumedReplacementKeys.has(key)) {
      patchedRuns.push(replacement)
      consumedReplacementKeys.add(key)
    }
  })
  return packGridTextRunsBufferV3(patchedRuns.toSorted(compareTextRunsByCell))
}

function collectPatchableTextDirtyItems(input: MaterializeGridRenderTileInputV3): readonly Item[] | null {
  const dirtyMasks = input.dirtyMasks
  const dirtyLocalRows = input.dirtyLocalRows
  const dirtyLocalCols = input.dirtyLocalCols
  if (
    !dirtyMasks ||
    !dirtyLocalRows ||
    !dirtyLocalCols ||
    dirtyLocalRows.length !== dirtyMasks.length * 2 ||
    dirtyLocalCols.length !== dirtyMasks.length * 2
  ) {
    return null
  }
  const structuralMask = DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
  const textDirtyMask = DirtyMaskV3.Value | DirtyMaskV3.Text
  const items: Item[] = []
  const seen = new Set<string>()
  for (let index = 0; index < dirtyMasks.length; index += 1) {
    const mask = dirtyMasks[index] ?? 0
    if ((mask & structuralMask) !== 0) {
      return null
    }
    if ((mask & textDirtyMask) === 0) {
      continue
    }
    const offset = index * 2
    const rowStart = clampDirtyLocalIndex(dirtyLocalRows[offset] ?? 0, input.viewport.rowEnd - input.viewport.rowStart + 1)
    const rowEnd = clampDirtyLocalIndex(dirtyLocalRows[offset + 1] ?? rowStart, input.viewport.rowEnd - input.viewport.rowStart + 1)
    const colStart = clampDirtyLocalIndex(dirtyLocalCols[offset] ?? 0, input.viewport.colEnd - input.viewport.colStart + 1)
    const colEnd = clampDirtyLocalIndex(dirtyLocalCols[offset + 1] ?? colStart, input.viewport.colEnd - input.viewport.colStart + 1)
    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        const globalRow = input.viewport.rowStart + row
        const globalCol = input.viewport.colStart + col
        const key = textRunCellKey(globalRow, globalCol)
        if (seen.has(key)) {
          continue
        }
        seen.add(key)
        items.push([globalCol, globalRow])
        if (items.length > 16) {
          return null
        }
      }
    }
  }
  return items
}

function sameTileBounds(tile: GridRenderTile, viewport: Viewport): boolean {
  return (
    tile.bounds.rowStart === viewport.rowStart &&
    tile.bounds.rowEnd === viewport.rowEnd &&
    tile.bounds.colStart === viewport.colStart &&
    tile.bounds.colEnd === viewport.colEnd
  )
}

function clampDirtyLocalIndex(value: number, length: number): number {
  return Math.max(0, Math.min(length - 1, Math.floor(value)))
}

function textRunCellKey(row: number, col: number): string {
  return `${String(row)}:${String(col)}`
}

function compareTextRunsByCell(
  left: { readonly col?: number | undefined; readonly row?: number | undefined },
  right: { readonly col?: number | undefined; readonly row?: number | undefined },
): number {
  return (left.row ?? 0) - (right.row ?? 0) || (left.col ?? 0) - (right.col ?? 0)
}

function canReuseTextRunsForTile(input: MaterializeGridRenderTileInputV3, tileId: number): boolean {
  const baseTile = input.reuseTextRunsFrom
  if (!baseTile || baseTile.tileId !== tileId) {
    return false
  }
  if (input.editingCell !== undefined && input.editingCell !== null) {
    return false
  }
  if (input.selectedCellSnapshot !== undefined && input.selectedCellSnapshot !== null) {
    return false
  }
  if (
    baseTile.bounds.rowStart !== input.viewport.rowStart ||
    baseTile.bounds.rowEnd !== input.viewport.rowEnd ||
    baseTile.bounds.colStart !== input.viewport.colStart ||
    baseTile.bounds.colEnd !== input.viewport.colEnd
  ) {
    return false
  }
  const dirtyMasks = input.dirtyMasks
  if (!dirtyMasks || dirtyMasks.length === 0) {
    return false
  }
  const textDirtyMask = DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
  for (const mask of dirtyMasks) {
    if ((mask & textDirtyMask) !== 0) {
      return false
    }
  }
  return true
}

export function createTileCellBoundsResolverV3(
  input: GridTileMaterializerAxisInputV3 & {
    readonly viewport: Viewport
    readonly minCol?: number | undefined
    readonly maxCol?: number | undefined
    readonly minRow?: number | undefined
    readonly maxRow?: number | undefined
  },
): (col: number, row: number) => Rectangle | undefined {
  const baseX = resolveColumnOffset(input.viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth)
  const baseY = resolveRowOffset(input.viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight)
  const minCol = input.minCol ?? input.viewport.colStart
  const maxCol = input.maxCol ?? input.viewport.colEnd
  const minRow = input.minRow ?? input.viewport.rowStart
  const maxRow = input.maxRow ?? input.viewport.rowEnd
  return (col, row) => {
    if (col < minCol || col > maxCol || row < minRow || row > maxRow) {
      return undefined
    }
    return {
      height: getResolvedRowHeight(input.rowHeights, row, input.gridMetrics.rowHeight),
      width: getResolvedColumnWidth(input.columnWidths, col, input.gridMetrics.columnWidth),
      x: resolveColumnOffset(col, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth) - baseX,
      y: resolveRowOffset(row, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight) - baseY,
    }
  }
}

function collectInboundTextSpillSourceItemsV3(input: MaterializeGridRenderTileInputV3): readonly Item[] {
  if (input.viewport.colStart <= 0) {
    return []
  }
  const items: Item[] = []
  const minSourceCol = Math.max(0, input.viewport.colStart - VIEWPORT_TILE_COLUMN_COUNT)
  for (let row = input.viewport.rowStart; row <= input.viewport.rowEnd; row += 1) {
    const firstTargetSnapshot = input.engine.getCell(input.sheetName, `${indexToColumn(input.viewport.colStart)}${row + 1}`)
    if (firstTargetSnapshot.value.tag !== ValueTag.Empty) {
      continue
    }
    for (let col = input.viewport.colStart - 1; col >= minSourceCol; col -= 1) {
      const snapshot = input.engine.getCell(input.sheetName, `${indexToColumn(col)}${row + 1}`)
      if (snapshot.value.tag === ValueTag.Empty) {
        continue
      }
      const renderCell = snapshotToRenderCell(snapshot, input.engine.getCellStyle(snapshot.styleId))
      if (
        renderCell.displayText.length > 0 &&
        !renderCell.wrap &&
        renderCell.align === 'left' &&
        (renderCell.kind === 'string' || renderCell.kind === 'error')
      ) {
        items.push([col, row])
      }
      break
    }
  }
  return items
}

export function resolveTileSurfaceSizeV3(
  viewport: Viewport,
  input: GridTileMaterializerAxisInputV3,
): { readonly width: number; readonly height: number } {
  return {
    height:
      resolveRowOffset(viewport.rowEnd + 1, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight) -
      resolveRowOffset(viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight),
    width:
      resolveColumnOffset(viewport.colEnd + 1, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth) -
      resolveColumnOffset(viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth),
  }
}
