export interface CellAxisIdentity {
  readonly sheetId: number
  readonly rowId: string
  readonly colId: string
}

interface DenseRowMajorIdentityBlock {
  readonly firstCellIndex: number
  readonly cellCount: number
  readonly colCount: number
  readonly sheetId: number
  readonly rowIds: readonly string[]
  readonly colIds: readonly string[]
}

export class CellAxisIdentityStore {
  private readonly sheetIds: Array<number | undefined> = []
  private readonly rowIds: Array<string | undefined> = []
  private readonly colIds: Array<string | undefined> = []
  private readonly denseRowMajorBlocks: DenseRowMajorIdentityBlock[] = []
  private readonly denseDeletedCellIndices = new Set<number>()

  get(cellIndex: number): CellAxisIdentity | undefined {
    const sheetId = this.sheetIds[cellIndex]
    const rowId = this.rowIds[cellIndex]
    const colId = this.colIds[cellIndex]
    if (sheetId !== undefined && rowId !== undefined && colId !== undefined) {
      return {
        sheetId,
        rowId,
        colId,
      }
    }
    return this.getDenseRowMajorIdentity(cellIndex)
  }

  set(cellIndex: number, identity: CellAxisIdentity): void {
    this.setParts(cellIndex, identity.sheetId, identity.rowId, identity.colId)
  }

  setParts(cellIndex: number, sheetId: number, rowId: string, colId: string): void {
    this.denseDeletedCellIndices.delete(cellIndex)
    this.sheetIds[cellIndex] = sheetId
    this.rowIds[cellIndex] = rowId
    this.colIds[cellIndex] = colId
  }

  setDenseRowMajorParts(firstCellIndex: number, sheetId: number, rowIds: readonly string[], colIds: readonly string[]): void {
    const colCount = colIds.length
    const cellCount = rowIds.length * colCount
    if (cellCount === 0) {
      return
    }
    this.denseRowMajorBlocks.push({
      firstCellIndex,
      cellCount,
      colCount,
      sheetId,
      rowIds,
      colIds,
    })
  }

  delete(cellIndex: number): boolean {
    const hasStoredIdentity =
      this.sheetIds[cellIndex] !== undefined && this.rowIds[cellIndex] !== undefined && this.colIds[cellIndex] !== undefined
    const hasDenseIdentity = this.getDenseRowMajorIdentity(cellIndex) !== undefined
    if (!hasStoredIdentity && !hasDenseIdentity) {
      return false
    }
    this.sheetIds[cellIndex] = undefined
    this.rowIds[cellIndex] = undefined
    this.colIds[cellIndex] = undefined
    if (hasDenseIdentity) {
      this.denseDeletedCellIndices.add(cellIndex)
    }
    return true
  }

  clear(): void {
    this.sheetIds.length = 0
    this.rowIds.length = 0
    this.colIds.length = 0
    this.denseRowMajorBlocks.length = 0
    this.denseDeletedCellIndices.clear()
  }

  forEach(callback: (identity: CellAxisIdentity, cellIndex: number) => void): void {
    for (let cellIndex = 0; cellIndex < this.rowIds.length; cellIndex += 1) {
      const sheetId = this.sheetIds[cellIndex]
      const rowId = this.rowIds[cellIndex]
      const colId = this.colIds[cellIndex]
      if (sheetId === undefined || rowId === undefined || colId === undefined) {
        continue
      }
      callback({ sheetId, rowId, colId }, cellIndex)
    }
    const emittedDenseCellIndices = new Set<number>()
    for (let blockIndex = this.denseRowMajorBlocks.length - 1; blockIndex >= 0; blockIndex -= 1) {
      const block = this.denseRowMajorBlocks[blockIndex]!
      for (let offset = 0; offset < block.cellCount; offset += 1) {
        const cellIndex = block.firstCellIndex + offset
        if (
          emittedDenseCellIndices.has(cellIndex) ||
          this.denseDeletedCellIndices.has(cellIndex) ||
          (this.sheetIds[cellIndex] !== undefined && this.rowIds[cellIndex] !== undefined && this.colIds[cellIndex] !== undefined)
        ) {
          continue
        }
        emittedDenseCellIndices.add(cellIndex)
        const rowOffset = Math.floor(offset / block.colCount)
        const colOffset = offset - rowOffset * block.colCount
        callback(
          {
            sheetId: block.sheetId,
            rowId: block.rowIds[rowOffset]!,
            colId: block.colIds[colOffset]!,
          },
          cellIndex,
        )
      }
    }
  }

  entries(): Array<readonly [number, CellAxisIdentity]> {
    const entries: Array<readonly [number, CellAxisIdentity]> = []
    this.forEach((identity, cellIndex) => {
      entries.push([cellIndex, identity])
    })
    return entries
  }

  private getDenseRowMajorIdentity(cellIndex: number): CellAxisIdentity | undefined {
    if (this.denseDeletedCellIndices.has(cellIndex)) {
      return undefined
    }
    for (let blockIndex = this.denseRowMajorBlocks.length - 1; blockIndex >= 0; blockIndex -= 1) {
      const block = this.denseRowMajorBlocks[blockIndex]!
      const offset = cellIndex - block.firstCellIndex
      if (offset < 0 || offset >= block.cellCount) {
        continue
      }
      const rowOffset = Math.floor(offset / block.colCount)
      const colOffset = offset - rowOffset * block.colCount
      return {
        sheetId: block.sheetId,
        rowId: block.rowIds[rowOffset]!,
        colId: block.colIds[colOffset]!,
      }
    }
    return undefined
  }
}
