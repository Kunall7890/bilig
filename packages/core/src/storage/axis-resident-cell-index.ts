export interface AxisResidentCellIdentity {
  readonly rowId: string
  readonly colId: string
}

export class AxisResidentCellIndex {
  private readonly byCell = new Map<number, AxisResidentCellIdentity>()
  private readonly byRow = new Map<string, Set<number>>()
  private readonly byColumn = new Map<string, Set<number>>()
  private secondaryIndexesDirty = false

  set(cellIndex: number, identity: AxisResidentCellIdentity): void {
    const existing = this.byCell.get(cellIndex)
    if (existing && !this.secondaryIndexesDirty) {
      deleteFromSetMap(this.byRow, existing.rowId, cellIndex)
      deleteFromSetMap(this.byColumn, existing.colId, cellIndex)
    }
    this.byCell.set(cellIndex, identity)
    if (!this.secondaryIndexesDirty) {
      addToSetMap(this.byRow, identity.rowId, cellIndex)
      addToSetMap(this.byColumn, identity.colId, cellIndex)
    }
  }

  add(cellIndex: number, identity: AxisResidentCellIdentity): void {
    const existing = this.byCell.get(cellIndex)
    if (existing && !this.secondaryIndexesDirty) {
      deleteFromSetMap(this.byRow, existing.rowId, cellIndex)
      deleteFromSetMap(this.byColumn, existing.colId, cellIndex)
    }
    this.byCell.set(cellIndex, identity)
    if (!this.secondaryIndexesDirty) {
      addToSetMap(this.byRow, identity.rowId, cellIndex)
      addToSetMap(this.byColumn, identity.colId, cellIndex)
    }
  }

  addDeferred(cellIndex: number, identity: AxisResidentCellIdentity): void {
    this.byCell.set(cellIndex, identity)
    this.secondaryIndexesDirty = true
  }

  get(cellIndex: number): AxisResidentCellIdentity | undefined {
    return this.byCell.get(cellIndex)
  }

  delete(cellIndex: number): boolean {
    const existing = this.byCell.get(cellIndex)
    if (!existing) {
      return false
    }
    this.byCell.delete(cellIndex)
    if (this.secondaryIndexesDirty) {
      return true
    }
    deleteFromSetMap(this.byRow, existing.rowId, cellIndex)
    deleteFromSetMap(this.byColumn, existing.colId, cellIndex)
    return true
  }

  clear(): void {
    this.byCell.clear()
    this.byRow.clear()
    this.byColumn.clear()
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

  private ensureSecondaryIndexes(): void {
    if (!this.secondaryIndexesDirty) {
      return
    }
    this.byRow.clear()
    this.byColumn.clear()
    this.byCell.forEach((identity, cellIndex) => {
      addToSetMap(this.byRow, identity.rowId, cellIndex)
      addToSetMap(this.byColumn, identity.colId, cellIndex)
    })
    this.secondaryIndexesDirty = false
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
