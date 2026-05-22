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
import type { GridRenderTile, GridRenderTileTextRun } from './render-tile-source.js'
import { resolveGridRenderTileDirtySpansV3 } from './render-tile-dirty-spans.js'
import { createGridTilePacketV3 } from './tile-packet-v3.js'
import { packTileKey53 } from './tile-key.js'
import { packGridTextBufferV3, packGridTextRunsV3, type PackedGridTextBufferV3 } from './text-run-buffer.js'
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
  readonly selectedCell?: Item | undefined
  readonly selectedCellSnapshot?: CellSnapshot | null | undefined
}

const STATIC_TILE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_TILE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})
const TEXT_ONLY_DIRTY_MASK_V3 = DirtyMaskV3.Value | DirtyMaskV3.Text
const MAX_FAST_TEXT_PATCH_CELLS_V3 = 16

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
  const textBuffer =
    reuseRectTile &&
    materializeSmallTextOnlyDirtyPatchV3({
      baseTile: reuseRectTile,
      getTextCellBounds,
      input,
      selectedCell,
      selectedCellSnapshot,
      surfaceSize,
      visibleRegion,
    })
  const resolvedTextBuffer =
    textBuffer ||
    packGridTextBufferV3(
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
    textRunCount: resolvedTextBuffer.textCount,
    textRuns: resolvedTextBuffer.textMetrics,
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
    textCount: resolvedTextBuffer.textCount,
    textMetrics: resolvedTextBuffer.textMetrics,
    textRuns: resolvedTextBuffer.textRuns,
    textSignature: resolvedTextBuffer.textSignature,
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

interface DirtyTextCellV3 {
  readonly col: number
  readonly row: number
}

function materializeSmallTextOnlyDirtyPatchV3(input: {
  readonly baseTile: GridRenderTile
  readonly getTextCellBounds: (col: number, row: number) => Rectangle | undefined
  readonly input: MaterializeGridRenderTileInputV3
  readonly selectedCell: Item
  readonly selectedCellSnapshot: CellSnapshot | null
  readonly surfaceSize: { readonly width: number; readonly height: number }
  readonly visibleRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows: number
    readonly freezeCols: number
  }
}): PackedGridTextBufferV3 | undefined {
  if (input.input.viewport.colStart > 0) {
    return undefined
  }
  if (input.baseTile.textRuns.length !== input.baseTile.textCount) {
    return undefined
  }
  if (input.baseTile.textRuns.some((run) => run.row === undefined || run.col === undefined)) {
    return undefined
  }

  const dirtyCells = collectSmallTextOnlyDirtyCellsV3(input.input)
  if (!dirtyCells || dirtyCells.length === 0) {
    return undefined
  }
  const dirtyCellKeys = new Set(dirtyCells.map((cell) => dirtyTextCellKeyV3(cell.col, cell.row)))
  const editingCellKey = input.input.editingCell ? dirtyTextCellKeyV3(input.input.editingCell[0], input.input.editingCell[1]) : null
  const editingCellBlankRun =
    editingCellKey === null
      ? undefined
      : input.baseTile.textRuns.find((run) => dirtyTextCellKeyV3(run.col ?? -1, run.row ?? -1) === editingCellKey)
  if (input.selectedCellSnapshot && !dirtyCellKeys.has(dirtyTextCellKeyV3(input.selectedCell[0], input.selectedCell[1]))) {
    return undefined
  }
  if (hasPreservedTextRunSpillingAcrossDirtyCellV3(input.baseTile.textRuns, dirtyCells)) {
    return undefined
  }

  const snapshotCache = new Map<string, CellSnapshot>()
  const cachedEngine = createTextPatchCachedEngineV3(input.input.engine, input.input.sheetName, snapshotCache)
  for (const cell of dirtyCells) {
    const address = `${indexToColumn(cell.col)}${cell.row + 1}`
    if (input.input.engine.getMergeRange?.(input.input.sheetName, address)) {
      return undefined
    }
    const cellKey = dirtyTextCellKeyV3(cell.col, cell.row)
    if (cellKey === editingCellKey) {
      continue
    }
    const snapshot = resolveCachedTextPatchSnapshotV3(input.input.engine, input.input.sheetName, address, snapshotCache)
    const renderSnapshot =
      input.selectedCellSnapshot?.address === address && input.selectedCell[0] === cell.col && input.selectedCell[1] === cell.row
        ? input.selectedCellSnapshot
        : snapshot
    if (
      !canPatchDirtyTextCellWithoutSpillScanV3({
        col: cell.col,
        engine: input.input.engine,
        row: cell.row,
        sheetName: input.input.sheetName,
        snapshot: renderSnapshot,
        snapshotCache,
        viewport: input.input.viewport,
      })
    ) {
      return undefined
    }
  }

  const dirtyTextBuffer = packGridTextBufferV3(
    buildGridTextScene({
      activeHeaderDrag: null,
      columnWidths: input.input.columnWidths,
      contentMode: 'data',
      editingCell: input.input.editingCell ?? null,
      engine: cachedEngine,
      getCellBounds: input.getTextCellBounds,
      gridMetrics: input.input.gridMetrics,
      hostBounds: {
        height: input.surfaceSize.height,
        left: 0,
        top: 0,
        width: input.surfaceSize.width,
      },
      hoveredHeader: null,
      resizeGuideColumn: null,
      rowHeights: input.input.rowHeights,
      selectedCell: input.selectedCell,
      selectedCellSnapshot: input.selectedCellSnapshot,
      selectionRange: null,
      sheetName: input.input.sheetName,
      visibleItems: dirtyCells.map((cell) => [cell.col, cell.row] as const),
      visibleRegion: input.visibleRegion,
    }),
  )

  const replacementRuns = new Map<string, GridRenderTileTextRun>()
  for (const run of dirtyTextBuffer.textRuns) {
    if (run.row === undefined || run.col === undefined) {
      return undefined
    }
    const key = dirtyTextCellKeyV3(run.col, run.row)
    if (!dirtyCellKeys.has(key)) {
      return undefined
    }
    replacementRuns.set(key, run)
  }
  if (editingCellKey && dirtyCellKeys.has(editingCellKey) && editingCellBlankRun) {
    replacementRuns.set(editingCellKey, { ...editingCellBlankRun, text: '' })
  }
  const expectedReplacementCount = dirtyCells.reduce((count, cell) => {
    const key = dirtyTextCellKeyV3(cell.col, cell.row)
    return count + (key === editingCellKey && !editingCellBlankRun ? 0 : 1)
  }, 0)
  if (replacementRuns.size !== expectedReplacementCount) {
    return undefined
  }

  const consumedReplacementKeys = new Set<string>()
  const nextRuns: GridRenderTileTextRun[] = []
  for (const run of input.baseTile.textRuns) {
    const key = dirtyTextCellKeyV3(run.col ?? -1, run.row ?? -1)
    if (!dirtyCellKeys.has(key)) {
      nextRuns.push(run)
      continue
    }
    const replacement = replacementRuns.get(key)
    if (replacement) {
      nextRuns.push(replacement)
      consumedReplacementKeys.add(key)
    }
  }
  for (const cell of dirtyCells) {
    const key = dirtyTextCellKeyV3(cell.col, cell.row)
    if (!consumedReplacementKeys.has(key)) {
      const replacement = replacementRuns.get(key)
      if (!replacement && key === editingCellKey) {
        continue
      }
      if (!replacement) {
        return undefined
      }
      nextRuns.push(replacement)
    }
  }

  return packGridTextRunsV3(nextRuns)
}

function collectSmallTextOnlyDirtyCellsV3(input: MaterializeGridRenderTileInputV3): readonly DirtyTextCellV3[] | null {
  const dirtyMasks = input.dirtyMasks
  const dirtyLocalRows = input.dirtyLocalRows
  const dirtyLocalCols = input.dirtyLocalCols
  if (!dirtyMasks || !dirtyLocalRows || !dirtyLocalCols) {
    return null
  }
  if (dirtyLocalRows.length !== dirtyMasks.length * 2 || dirtyLocalCols.length !== dirtyMasks.length * 2) {
    return null
  }
  const rowCount = input.viewport.rowEnd - input.viewport.rowStart + 1
  const colCount = input.viewport.colEnd - input.viewport.colStart + 1
  const cells: DirtyTextCellV3[] = []
  const seen = new Set<string>()
  for (let index = 0; index < dirtyMasks.length; index += 1) {
    const mask = dirtyMasks[index] ?? 0
    if ((mask & ~TEXT_ONLY_DIRTY_MASK_V3) !== 0) {
      return null
    }
    if ((mask & TEXT_ONLY_DIRTY_MASK_V3) === 0) {
      continue
    }
    const offset = index * 2
    const rowStart = dirtyLocalRows[offset] ?? 0
    const rowEnd = dirtyLocalRows[offset + 1] ?? rowStart
    const colStart = dirtyLocalCols[offset] ?? 0
    const colEnd = dirtyLocalCols[offset + 1] ?? colStart
    if (rowStart < 0 || rowEnd < rowStart || rowEnd >= rowCount || colStart < 0 || colEnd < colStart || colEnd >= colCount) {
      return null
    }
    for (let localRow = rowStart; localRow <= rowEnd; localRow += 1) {
      for (let localCol = colStart; localCol <= colEnd; localCol += 1) {
        const col = input.viewport.colStart + localCol
        const row = input.viewport.rowStart + localRow
        const key = dirtyTextCellKeyV3(col, row)
        if (seen.has(key)) {
          continue
        }
        seen.add(key)
        cells.push({ col, row })
        if (cells.length > MAX_FAST_TEXT_PATCH_CELLS_V3) {
          return null
        }
      }
    }
  }
  return cells
}

function createTextPatchCachedEngineV3(
  engine: GridEngineLike,
  sheetName: string,
  snapshotCache: ReadonlyMap<string, CellSnapshot>,
): GridEngineLike {
  const cachedEngine: GridEngineLike = {
    getCell: (requestedSheetName, address) =>
      requestedSheetName === sheetName && snapshotCache.has(address)
        ? snapshotCache.get(address)!
        : engine.getCell(requestedSheetName, address),
    getCellStyle: engine.getCellStyle.bind(engine),
    subscribeCells: engine.subscribeCells.bind(engine),
    workbook: engine.workbook,
  }
  if (engine.getMergeRange) {
    cachedEngine.getMergeRange = engine.getMergeRange.bind(engine)
  }
  if (engine.getRenderRevisionSnapshot) {
    cachedEngine.getRenderRevisionSnapshot = engine.getRenderRevisionSnapshot.bind(engine)
  }
  if (engine.listMergeRanges) {
    cachedEngine.listMergeRanges = engine.listMergeRanges.bind(engine)
  }
  if (engine.subscribeSheetChannel) {
    cachedEngine.subscribeSheetChannel = engine.subscribeSheetChannel.bind(engine)
  }
  return cachedEngine
}

function resolveCachedTextPatchSnapshotV3(
  engine: GridEngineLike,
  sheetName: string,
  address: string,
  snapshotCache: Map<string, CellSnapshot>,
): CellSnapshot {
  const cached = snapshotCache.get(address)
  if (cached) {
    return cached
  }
  const snapshot = engine.getCell(sheetName, address)
  snapshotCache.set(address, snapshot)
  return snapshot
}

function canPatchDirtyTextCellWithoutSpillScanV3(input: {
  readonly col: number
  readonly engine: GridEngineLike
  readonly row: number
  readonly sheetName: string
  readonly snapshot: CellSnapshot
  readonly snapshotCache: Map<string, CellSnapshot>
  readonly viewport: Viewport
}): boolean {
  const { engine, snapshot } = input
  if (snapshot.value.tag === ValueTag.Boolean) {
    return false
  }
  const renderCell = snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId))
  if (renderCell.displayText.length === 0) {
    return false
  }
  if (renderCell.wrap || renderCell.align !== 'left' || (renderCell.kind !== 'string' && renderCell.kind !== 'error')) {
    return true
  }
  if (input.col >= input.viewport.colEnd) {
    return false
  }
  const blockerAddress = `${indexToColumn(input.col + 1)}${input.row + 1}`
  const blockerSnapshot = resolveCachedTextPatchSnapshotV3(engine, input.sheetName, blockerAddress, input.snapshotCache)
  const blockerRenderCell = snapshotToRenderCell(blockerSnapshot, engine.getCellStyle(blockerSnapshot.styleId))
  return blockerRenderCell.displayText.length > 0
}

function hasPreservedTextRunSpillingAcrossDirtyCellV3(
  runs: readonly GridRenderTileTextRun[],
  dirtyCells: readonly DirtyTextCellV3[],
): boolean {
  for (const run of runs) {
    if (run.row === undefined || run.col === undefined || run.spillColEnd === undefined) {
      continue
    }
    for (const cell of dirtyCells) {
      if (run.row === cell.row && run.col < cell.col && run.spillColEnd >= cell.col) {
        return true
      }
    }
  }
  return false
}

function dirtyTextCellKeyV3(col: number, row: number): string {
  return `${row}:${col}`
}

function canReuseStaticGridRectsForTile(input: MaterializeGridRenderTileInputV3, tileId: number): boolean {
  const baseTile = input.reuseStaticGridRectsFrom
  if (!baseTile || baseTile.tileId !== tileId) {
    return false
  }
  const contentRevision = resolveMaterializerContentRevision(input.engine)
  if (contentRevision !== null && baseTile.lastBatchId < contentRevision) {
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
  const rectDirtyMask =
    DirtyMaskV3.Style | DirtyMaskV3.Rect | DirtyMaskV3.Border | DirtyMaskV3.AxisX | DirtyMaskV3.AxisY | DirtyMaskV3.Freeze
  for (const mask of dirtyMasks) {
    if ((mask & rectDirtyMask) !== 0) {
      return false
    }
  }
  return true
}

function resolveMaterializerContentRevision(engine: GridEngineLike): number | null {
  const revision = engine.getRenderRevisionSnapshot?.()
  if (!revision) {
    return null
  }
  let result: number | null = null
  for (const value of [revision.authoritativeRevision, revision.localRevision, revision.projectedRevision]) {
    if (value === null || value === undefined || !Number.isFinite(value)) {
      continue
    }
    result = result === null ? value : Math.max(result, value)
  }
  return result
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
