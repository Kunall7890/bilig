export interface AxisResidentCellIdentity {
  readonly sheetId?: number
  readonly rowId: string
  readonly colId: string
}

export type AxisResidentCellIndexRebuildSource = (callback: (cellIndex: number, identity: AxisResidentCellIdentity) => void) => void

interface DenseRowMajorResidentCellBlock {
  readonly firstCellIndex: number
  readonly cellCount: number
  readonly colCount: number
  readonly rowIds: readonly string[]
  readonly colIds: readonly string[]
  rowOffsets?: ReadonlyMap<string, number>
  colOffsets?: ReadonlyMap<string, number>
}

export class AxisResidentCellIndex {
  private readonly sheetIds: Array<number | undefined> = []
  private readonly rowIds: Array<string | undefined> = []
  private readonly colIds: Array<string | undefined> = []
  private readonly byRow = new Map<string, Set<number>>()
  private readonly byColumn = new Map<string, Set<number>>()
  private readonly denseRowMajorBlocks: DenseRowMajorResidentCellBlock[] = []
  private readonly denseDeletedCellIndices = new Set<number>()
  private primaryIndexDirty = false
  private secondaryIndexesDirty = false

  constructor(private readonly rebuildSource?: AxisResidentCellIndexRebuildSource) {}

  set(cellIndex: number, identity: AxisResidentCellIdentity): void {
    this.ensurePrimaryIndexUnlessDense()
    const existingRowId = this.rowIds[cellIndex]
    const existingColId = this.colIds[cellIndex]
    if (existingRowId !== undefined && existingColId !== undefined && !this.secondaryIndexesDirty) {
      deleteFromSetMap(this.byRow, existingRowId, cellIndex)
      deleteFromSetMap(this.byColumn, existingColId, cellIndex)
    }
    this.setPrimaryParts(cellIndex, identity.sheetId, identity.rowId, identity.colId)
    if (!this.secondaryIndexesDirty) {
      addToSetMap(this.byRow, identity.rowId, cellIndex)
      addToSetMap(this.byColumn, identity.colId, cellIndex)
    }
  }

  add(cellIndex: number, identity: AxisResidentCellIdentity): void {
    this.ensurePrimaryIndexUnlessDense()
    const existingRowId = this.rowIds[cellIndex]
    const existingColId = this.colIds[cellIndex]
    if (existingRowId !== undefined && existingColId !== undefined && !this.secondaryIndexesDirty) {
      deleteFromSetMap(this.byRow, existingRowId, cellIndex)
      deleteFromSetMap(this.byColumn, existingColId, cellIndex)
    }
    this.setPrimaryParts(cellIndex, identity.sheetId, identity.rowId, identity.colId)
    if (!this.secondaryIndexesDirty) {
      addToSetMap(this.byRow, identity.rowId, cellIndex)
      addToSetMap(this.byColumn, identity.colId, cellIndex)
    }
  }

  addDeferred(cellIndex: number, identity: AxisResidentCellIdentity): void {
    this.ensurePrimaryIndexUnlessDense()
    this.setPrimaryParts(cellIndex, identity.sheetId, identity.rowId, identity.colId)
    this.secondaryIndexesDirty = true
  }

  addDeferredParts(cellIndex: number, rowId: string, colId: string): void {
    this.ensurePrimaryIndexUnlessDense()
    this.setPrimaryParts(cellIndex, undefined, rowId, colId)
    this.secondaryIndexesDirty = true
  }

  addDenseRowMajorDeferredParts(firstCellIndex: number, rowIds: readonly string[], colIds: readonly string[]): void {
    const colCount = colIds.length
    const cellCount = rowIds.length * colCount
    if (cellCount === 0) {
      return
    }
    this.primaryIndexDirty = false
    this.denseRowMajorBlocks.push({
      firstCellIndex,
      cellCount,
      colCount,
      rowIds,
      colIds,
    })
    this.secondaryIndexesDirty = true
  }

  deferRebuild(): void {
    this.sheetIds.length = 0
    this.rowIds.length = 0
    this.colIds.length = 0
    this.byRow.clear()
    this.byColumn.clear()
    this.denseRowMajorBlocks.length = 0
    this.denseDeletedCellIndices.clear()
    this.primaryIndexDirty = true
    this.secondaryIndexesDirty = false
  }

  get(cellIndex: number): AxisResidentCellIdentity | undefined {
    this.ensurePrimaryIndexUnlessDense()
    const rowId = this.rowIds[cellIndex]
    const colId = this.colIds[cellIndex]
    if (rowId === undefined || colId === undefined) {
      return this.getDenseRowMajorIdentity(cellIndex)
    }
    const sheetId = this.sheetIds[cellIndex]
    return sheetId === undefined ? { rowId, colId } : { sheetId, rowId, colId }
  }

  delete(cellIndex: number): boolean {
    this.ensurePrimaryIndexUnlessDense()
    const existingRowId = this.rowIds[cellIndex]
    const existingColId = this.colIds[cellIndex]
    const existingDenseIdentity = this.getDenseRowMajorIdentity(cellIndex)
    if (existingRowId === undefined || existingColId === undefined) {
      if (!existingDenseIdentity) {
        return false
      }
      this.denseDeletedCellIndices.add(cellIndex)
      return true
    }
    this.sheetIds[cellIndex] = undefined
    this.rowIds[cellIndex] = undefined
    this.colIds[cellIndex] = undefined
    if (this.secondaryIndexesDirty) {
      return true
    }
    deleteFromSetMap(this.byRow, existingRowId, cellIndex)
    deleteFromSetMap(this.byColumn, existingColId, cellIndex)
    return true
  }

  clear(): void {
    this.sheetIds.length = 0
    this.rowIds.length = 0
    this.colIds.length = 0
    this.byRow.clear()
    this.byColumn.clear()
    this.denseRowMajorBlocks.length = 0
    this.denseDeletedCellIndices.clear()
    this.primaryIndexDirty = false
    this.secondaryIndexesDirty = false
  }

  cellsInRow(rowId: string): number[] {
    this.ensureSecondaryIndexes()
    const cells = new Set(this.byRow.get(rowId))
    this.addDenseCellsInRow(rowId, cells)
    return sortedCells(cells)
  }

  cellsInColumn(colId: string): number[] {
    this.ensureSecondaryIndexes()
    const cells = new Set(this.byColumn.get(colId))
    this.addDenseCellsInColumn(colId, cells)
    return sortedCells(cells)
  }

  forEachCellInRow(rowId: string, callback: (cellIndex: number) => void): void {
    this.ensureSecondaryIndexes()
    this.byRow.get(rowId)?.forEach(callback)
    this.forEachDenseCellInRow(rowId, callback)
  }

  forEachCellInColumn(colId: string, callback: (cellIndex: number) => void): void {
    this.ensureSecondaryIndexes()
    this.byColumn.get(colId)?.forEach(callback)
    this.forEachDenseCellInColumn(colId, callback)
  }

  cellsInRows(rowIds: readonly string[]): number[] {
    return sortedUniqueCells(rowIds.flatMap((rowId) => this.cellsInRow(rowId)))
  }

  cellsInColumns(colIds: readonly string[]): number[] {
    return sortedUniqueCells(colIds.flatMap((colId) => this.cellsInColumn(colId)))
  }

  cellsInRowsUnordered(rowIds: readonly string[]): number[] {
    this.ensureSecondaryIndexes()
    return uniqueCells(rowIds.flatMap((rowId) => this.unorderedCellsInRow(rowId)))
  }

  cellsInColumnsUnordered(colIds: readonly string[]): number[] {
    this.ensureSecondaryIndexes()
    return uniqueCells(colIds.flatMap((colId) => this.unorderedCellsInColumn(colId)))
  }

  private ensurePrimaryIndexUnlessDense(): void {
    if (!this.primaryIndexDirty || this.denseRowMajorBlocks.length > 0) {
      return
    }
    this.ensurePrimaryIndex()
  }

  private ensurePrimaryIndex(): void {
    if (!this.primaryIndexDirty) {
      return
    }
    this.sheetIds.length = 0
    this.rowIds.length = 0
    this.colIds.length = 0
    this.rebuildSource?.((cellIndex, identity) => {
      this.setPrimaryParts(cellIndex, identity.sheetId, identity.rowId, identity.colId)
    })
    this.primaryIndexDirty = false
    this.secondaryIndexesDirty = true
  }

  private ensureSecondaryIndexes(): void {
    this.ensurePrimaryIndexUnlessDense()
    if (!this.secondaryIndexesDirty) {
      return
    }
    this.byRow.clear()
    this.byColumn.clear()
    for (let cellIndex = 0; cellIndex < this.rowIds.length; cellIndex += 1) {
      const rowId = this.rowIds[cellIndex]
      const colId = this.colIds[cellIndex]
      if (rowId === undefined || colId === undefined) {
        continue
      }
      addToSetMap(this.byRow, rowId, cellIndex)
      addToSetMap(this.byColumn, colId, cellIndex)
    }
    this.secondaryIndexesDirty = false
  }

  private setPrimaryParts(cellIndex: number, sheetId: number | undefined, rowId: string, colId: string): void {
    this.denseDeletedCellIndices.delete(cellIndex)
    this.sheetIds[cellIndex] = sheetId
    this.rowIds[cellIndex] = rowId
    this.colIds[cellIndex] = colId
  }

  private unorderedCellsInRow(rowId: string): number[] {
    const cells = [...(this.byRow.get(rowId) ?? [])]
    this.forEachDenseCellInRow(rowId, (cellIndex) => {
      cells.push(cellIndex)
    })
    return cells
  }

  private unorderedCellsInColumn(colId: string): number[] {
    const cells = [...(this.byColumn.get(colId) ?? [])]
    this.forEachDenseCellInColumn(colId, (cellIndex) => {
      cells.push(cellIndex)
    })
    return cells
  }

  private addDenseCellsInRow(rowId: string, cells: Set<number>): void {
    this.forEachDenseCellInRow(rowId, (cellIndex) => {
      cells.add(cellIndex)
    })
  }

  private addDenseCellsInColumn(colId: string, cells: Set<number>): void {
    this.forEachDenseCellInColumn(colId, (cellIndex) => {
      cells.add(cellIndex)
    })
  }

  private forEachDenseCellInRow(rowId: string, callback: (cellIndex: number) => void): void {
    for (let blockIndex = this.denseRowMajorBlocks.length - 1; blockIndex >= 0; blockIndex -= 1) {
      const block = this.denseRowMajorBlocks[blockIndex]!
      block.rowOffsets ??= indexedAxisIds(block.rowIds)
      const rowOffset = block.rowOffsets.get(rowId)
      if (rowOffset === undefined) {
        continue
      }
      const rowCellStart = block.firstCellIndex + rowOffset * block.colCount
      for (let colOffset = 0; colOffset < block.colCount; colOffset += 1) {
        const cellIndex = rowCellStart + colOffset
        if (this.isDenseCellShadowed(cellIndex)) {
          continue
        }
        callback(cellIndex)
      }
    }
  }

  private forEachDenseCellInColumn(colId: string, callback: (cellIndex: number) => void): void {
    for (let blockIndex = this.denseRowMajorBlocks.length - 1; blockIndex >= 0; blockIndex -= 1) {
      const block = this.denseRowMajorBlocks[blockIndex]!
      block.colOffsets ??= indexedAxisIds(block.colIds)
      const colOffset = block.colOffsets.get(colId)
      if (colOffset === undefined) {
        continue
      }
      for (let rowOffset = 0; rowOffset < block.rowIds.length; rowOffset += 1) {
        const cellIndex = block.firstCellIndex + rowOffset * block.colCount + colOffset
        if (this.isDenseCellShadowed(cellIndex)) {
          continue
        }
        callback(cellIndex)
      }
    }
  }

  private getDenseRowMajorIdentity(cellIndex: number): AxisResidentCellIdentity | undefined {
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
        rowId: block.rowIds[rowOffset]!,
        colId: block.colIds[colOffset]!,
      }
    }
    return undefined
  }

  private isDenseCellShadowed(cellIndex: number): boolean {
    return this.denseDeletedCellIndices.has(cellIndex) || (this.rowIds[cellIndex] !== undefined && this.colIds[cellIndex] !== undefined)
  }
}

function indexedAxisIds(axisIds: readonly string[]): ReadonlyMap<string, number> {
  const offsets = new Map<string, number>()
  for (let index = 0; index < axisIds.length; index += 1) {
    offsets.set(axisIds[index]!, index)
  }
  return offsets
}

function addToSetMap(map: Map<string, Set<number>>, key: string, value: number): void {
  let values = map.get(key)
  if (!values) {
    values = new Set<number>()
    map.set(key, values)
  }
  values.add(value)
}

function deleteFromSetMap(map: Map<string, Set<number>>, key: string, value: number): void {
  const values = map.get(key)
  if (!values) {
    return
  }
  values.delete(value)
  if (values.size === 0) {
    map.delete(key)
  }
}

function sortedCells(values: ReadonlySet<number> | undefined): number[] {
  return values ? [...values].toSorted((left, right) => left - right) : []
}

function sortedUniqueCells(values: readonly number[]): number[] {
  return [...new Set(values)].toSorted((left, right) => left - right)
}

function uniqueCells(values: readonly number[]): number[] {
  return [...new Set(values)]
}
