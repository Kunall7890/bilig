export interface AxisResidentCellIdentity {
  readonly sheetId?: number
  readonly rowId: string
  readonly colId: string
}

export type AxisResidentCellIndexRebuildSource = (callback: (cellIndex: number, identity: AxisResidentCellIdentity) => void) => void

export class AxisResidentCellIndex {
  private readonly sheetIds: Array<number | undefined> = []
  private readonly rowIds: Array<string | undefined> = []
  private readonly colIds: Array<string | undefined> = []
  private readonly byRow = new Map<string, Set<number>>()
  private readonly byColumn = new Map<string, Set<number>>()
  private primaryIndexDirty = false
  private secondaryIndexesDirty = false

  constructor(private readonly rebuildSource?: AxisResidentCellIndexRebuildSource) {}

  set(cellIndex: number, identity: AxisResidentCellIdentity): void {
    this.ensurePrimaryIndex()
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
    this.ensurePrimaryIndex()
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
    this.ensurePrimaryIndex()
    this.setPrimaryParts(cellIndex, identity.sheetId, identity.rowId, identity.colId)
    this.secondaryIndexesDirty = true
  }

  addDeferredParts(cellIndex: number, rowId: string, colId: string): void {
    this.ensurePrimaryIndex()
    this.setPrimaryParts(cellIndex, undefined, rowId, colId)
    this.secondaryIndexesDirty = true
  }

  addDenseRowMajorDeferredParts(firstCellIndex: number, rowIds: readonly string[], colIds: readonly string[]): void {
    this.ensurePrimaryIndex()
    const colCount = colIds.length
    const cellCount = rowIds.length * colCount
    if (cellCount === 0) {
      return
    }
    const endCellIndex = firstCellIndex + cellCount
    this.sheetIds.length = Math.max(this.sheetIds.length, endCellIndex)
    this.rowIds.length = Math.max(this.rowIds.length, endCellIndex)
    this.colIds.length = Math.max(this.colIds.length, endCellIndex)

    let cellIndex = firstCellIndex
    for (let rowOffset = 0; rowOffset < rowIds.length; rowOffset += 1) {
      const rowId = rowIds[rowOffset]!
      for (let colOffset = 0; colOffset < colCount; colOffset += 1) {
        this.sheetIds[cellIndex] = undefined
        this.rowIds[cellIndex] = rowId
        this.colIds[cellIndex] = colIds[colOffset]!
        cellIndex += 1
      }
    }
    this.secondaryIndexesDirty = true
  }

  deferRebuild(): void {
    this.sheetIds.length = 0
    this.rowIds.length = 0
    this.colIds.length = 0
    this.byRow.clear()
    this.byColumn.clear()
    this.primaryIndexDirty = true
    this.secondaryIndexesDirty = false
  }

  get(cellIndex: number): AxisResidentCellIdentity | undefined {
    this.ensurePrimaryIndex()
    const rowId = this.rowIds[cellIndex]
    const colId = this.colIds[cellIndex]
    if (rowId === undefined || colId === undefined) {
      return undefined
    }
    const sheetId = this.sheetIds[cellIndex]
    return sheetId === undefined ? { rowId, colId } : { sheetId, rowId, colId }
  }

  delete(cellIndex: number): boolean {
    this.ensurePrimaryIndex()
    const existingRowId = this.rowIds[cellIndex]
    const existingColId = this.colIds[cellIndex]
    if (existingRowId === undefined || existingColId === undefined) {
      return false
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
    this.primaryIndexDirty = false
    this.secondaryIndexesDirty = false
  }

  cellsInRow(rowId: string): number[] {
    this.ensureSecondaryIndexes()
    return sortedCells(this.byRow.get(rowId))
  }

  cellsInColumn(colId: string): number[] {
    this.ensureSecondaryIndexes()
    return sortedCells(this.byColumn.get(colId))
  }

  forEachCellInRow(rowId: string, callback: (cellIndex: number) => void): void {
    this.ensureSecondaryIndexes()
    this.byRow.get(rowId)?.forEach(callback)
  }

  forEachCellInColumn(colId: string, callback: (cellIndex: number) => void): void {
    this.ensureSecondaryIndexes()
    this.byColumn.get(colId)?.forEach(callback)
  }

  cellsInRows(rowIds: readonly string[]): number[] {
    return sortedUniqueCells(rowIds.flatMap((rowId) => this.cellsInRow(rowId)))
  }

  cellsInColumns(colIds: readonly string[]): number[] {
    return sortedUniqueCells(colIds.flatMap((colId) => this.cellsInColumn(colId)))
  }

  cellsInRowsUnordered(rowIds: readonly string[]): number[] {
    this.ensureSecondaryIndexes()
    return uniqueCells(rowIds.flatMap((rowId) => [...(this.byRow.get(rowId) ?? [])]))
  }

  cellsInColumnsUnordered(colIds: readonly string[]): number[] {
    this.ensureSecondaryIndexes()
    return uniqueCells(colIds.flatMap((colId) => [...(this.byColumn.get(colId) ?? [])]))
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
    this.ensurePrimaryIndex()
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
    this.sheetIds[cellIndex] = sheetId
    this.rowIds[cellIndex] = rowId
    this.colIds[cellIndex] = colId
  }
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
