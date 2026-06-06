import { parseCellAddress } from '@bilig/formula'
import type { WorkbookSnapshot } from '@bilig/protocol'
import type { SheetRecord } from '../workbook-store.js'

type WorkbookSnapshotCell = WorkbookSnapshot['sheets'][number]['cells'][number]
type WorkbookSnapshotSheet = WorkbookSnapshot['sheets'][number]
type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: unknown
}

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')

interface SnapshotCellCoordinate {
  readonly row: number
  readonly col: number
}

interface RuntimeImageSheetCellsSnapshotLike {
  readonly coords: readonly SnapshotCellCoordinate[]
  readonly coordinateOrder?: 'dense-row-major'
  readonly dimensions?: {
    readonly width: number
    readonly height: number
  }
  readonly cellCount?: number
}

interface FreshRuntimeLogicalSheetInternals {
  readonly deferVisibleCellPageRebuild?: () => void
  readonly setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred?: (
    firstCellIndex: number,
    rowIds: readonly string[],
    colIds: readonly string[],
  ) => void
}

interface DenseSnapshotSheetRestoreRun {
  readonly startIndex: number
  readonly rows: readonly number[]
  readonly cols: readonly number[]
  readonly cellCount: number
}

interface DenseRuntimeSheetRestorePlan {
  readonly width: number
  readonly height: number
}

function hasSnapshotCoordinate(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function readSnapshotCellCoordinates(sheetName: string, cell: WorkbookSnapshotCell): SnapshotCellCoordinate {
  if (hasSnapshotCoordinate(cell.row) && hasSnapshotCoordinate(cell.col)) {
    return {
      row: cell.row,
      col: cell.col,
    }
  }
  const parsed = parseCellAddress(cell.address, sheetName)
  return {
    row: parsed.row,
    col: parsed.col,
  }
}

function isFreshRuntimeLogicalSheetInternals(value: unknown): value is FreshRuntimeLogicalSheetInternals {
  return typeof value === 'object' && value !== null
}

function sameNumberList(left: readonly number[], right: readonly number[]): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

export function attachDenseFreshRuntimeCellIdentities(
  sheet: SheetRecord,
  firstCellIndex: number,
  rowIds: readonly string[],
  colIds: readonly string[],
): boolean {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshRuntimeLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const attachDenseFreshVisibleCellIdentities = logical?.setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred?.bind(logical)
  if (!attachDenseFreshVisibleCellIdentities) {
    return false
  }
  logical?.deferVisibleCellPageRebuild?.()
  attachDenseFreshVisibleCellIdentities(firstCellIndex, rowIds, colIds)
  return true
}

export function attachDenseFreshRuntimeCells(
  sheet: SheetRecord,
  firstCellIndex: number,
  rowStart: number,
  colStart: number,
  rowIds: readonly string[],
  colIds: readonly string[],
): boolean {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshRuntimeLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const attachDenseFreshVisibleCellIdentities = logical?.setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred?.bind(logical)
  if (!attachDenseFreshVisibleCellIdentities) {
    return false
  }
  logical?.deferVisibleCellPageRebuild?.()
  attachDenseFreshVisibleCellIdentities(firstCellIndex, rowIds, colIds)
  sheet.grid.setDenseRowMajor(rowStart, colStart, rowIds.length, colIds.length, firstCellIndex)
  return true
}

export function getDenseRuntimeSheetRestorePlan(
  sheet: WorkbookSnapshot['sheets'][number],
  sheetCells: RuntimeImageSheetCellsSnapshotLike | undefined,
): DenseRuntimeSheetRestorePlan | undefined {
  const dimensions = sheetCells?.dimensions
  if (!dimensions) {
    return undefined
  }
  const { width, height } = dimensions
  const cellCount = sheetCells.cellCount ?? sheetCells.coords.length
  const hasDenseCoordinateOrder = sheetCells.coordinateOrder === 'dense-row-major'
  if (
    !Number.isInteger(width) ||
    !Number.isInteger(height) ||
    width <= 0 ||
    height <= 0 ||
    cellCount !== sheet.cells.length ||
    width * height !== sheet.cells.length
  ) {
    return undefined
  }
  if (hasDenseCoordinateOrder) {
    return { width, height }
  }
  if (sheetCells.coords.length !== sheet.cells.length) {
    return undefined
  }
  for (let index = 0; index < sheetCells.coords.length; index += 1) {
    const coords = sheetCells.coords[index]!
    if (coords.row !== Math.floor(index / width) || coords.col !== index % width) {
      return undefined
    }
  }
  return { width, height }
}

export function getDenseSnapshotSheetRestoreRuns(
  sheet: WorkbookSnapshot['sheets'][number],
): readonly DenseSnapshotSheetRestoreRun[] | undefined {
  if (sheet.cells.length <= 1) {
    return undefined
  }
  const runs: DenseSnapshotSheetRestoreRun[] = []
  let currentRow = -1
  let previousCol = -1
  let currentRowCols: number[] = []
  let currentRowStartIndex = 0
  let runStartIndex = 0
  let runRows: number[] = []
  let runCols: readonly number[] | undefined

  const flushRun = (): void => {
    if (runCols === undefined || runRows.length === 0) {
      return
    }
    runs.push({
      startIndex: runStartIndex,
      rows: runRows,
      cols: runCols,
      cellCount: runRows.length * runCols.length,
    })
  }

  const finishRow = (): void => {
    if (currentRow < 0 || currentRowCols.length === 0) {
      return
    }
    if (runCols !== undefined && sameNumberList(runCols, currentRowCols)) {
      runRows.push(currentRow)
      return
    }
    flushRun()
    runStartIndex = currentRowStartIndex
    runRows = [currentRow]
    runCols = currentRowCols
  }

  for (let index = 0; index < sheet.cells.length; index += 1) {
    const coords = readSnapshotCellCoordinates(sheet.name, sheet.cells[index]!)
    if (coords.row < currentRow || (coords.row === currentRow && coords.col <= previousCol)) {
      return undefined
    }
    if (coords.row !== currentRow) {
      finishRow()
      currentRow = coords.row
      previousCol = -1
      currentRowCols = []
      currentRowStartIndex = index
    }
    currentRowCols.push(coords.col)
    previousCol = coords.col
  }
  finishRow()
  flushRun()
  if (runs.length === 0 || runs.length >= sheet.cells.length) {
    return undefined
  }
  return runs
}

export function shouldReleaseImportedSourceSnapshotCells(snapshot: WorkbookSnapshot): boolean {
  const source = (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
  return source !== undefined && !(source instanceof Uint8Array)
}

export function releaseRestoredSheetCells(sheet: WorkbookSnapshotSheet, shouldRelease: boolean): void {
  if (shouldRelease) {
    sheet.cells = []
  }
}
