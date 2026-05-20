import { attachRuntimeImage } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'

export interface ImportedRuntimeCellCoordinate {
  readonly row: number
  readonly col: number
}

export interface ImportedRuntimeSheetCells {
  readonly sheetName: string
  readonly coords: readonly ImportedRuntimeCellCoordinate[]
  readonly coordinateOrder?: 'dense-row-major'
  readonly dimensions: {
    readonly width: number
    readonly height: number
  }
  readonly cellCount: number
}

function importedRuntimeCoordinatesAreDenseRowMajor(
  coords: readonly ImportedRuntimeCellCoordinate[],
  width: number,
  height: number,
): boolean {
  if (width <= 0 || height <= 0 || coords.length !== width * height) {
    return false
  }
  for (let index = 0; index < coords.length; index += 1) {
    const coord = coords[index]!
    if (coord.row !== Math.floor(index / width) || coord.col !== index % width) {
      return false
    }
  }
  return true
}

export function createImportedRuntimeSheetCells(args: {
  readonly sheetName: string
  readonly coords: readonly ImportedRuntimeCellCoordinate[]
  readonly width: number
  readonly height: number
}): ImportedRuntimeSheetCells {
  const coordinateOrder = importedRuntimeCoordinatesAreDenseRowMajor(args.coords, args.width, args.height) ? 'dense-row-major' : undefined
  return {
    sheetName: args.sheetName,
    coords: args.coords,
    ...(coordinateOrder ? { coordinateOrder } : {}),
    dimensions: { width: args.width, height: args.height },
    cellCount: args.coords.length,
  }
}

export function pushImportedSnapshotCell(
  cells: WorkbookSnapshot['sheets'][number]['cells'],
  coords: ImportedRuntimeCellCoordinate[],
  cell: WorkbookSnapshot['sheets'][number]['cells'][number],
  row: number,
  col: number,
): void {
  cell.row = row
  cell.col = col
  cells.push(cell)
  coords.push({ row, col })
}

export function attachImportedRuntimeCoordinates(
  snapshot: WorkbookSnapshot,
  sheetCells: readonly ImportedRuntimeSheetCells[],
): WorkbookSnapshot {
  return attachRuntimeImage(snapshot, {
    version: 1,
    templateBank: [],
    formulaInstances: [],
    formulaValues: [],
    sheetCells,
  })
}
