import type { WorkbookAxisEntryRecord } from './workbook-metadata-types.js'

type WorkbookAxis = 'row' | 'column'

export class WorkbookIdAllocator {
  private nextRowAxisId = 1
  private nextColumnAxisId = 1
  private nextLogicalRowAxisId = 1
  private nextLogicalColumnAxisId = 1
  private nextStyleId = 1
  private nextFormatId = 1

  reset(): void {
    this.nextRowAxisId = 1
    this.nextColumnAxisId = 1
    this.nextLogicalRowAxisId = 1
    this.nextLogicalColumnAxisId = 1
    this.nextStyleId = 1
    this.nextFormatId = 1
  }

  bumpStyleId(id: string): void {
    const match = /^style-(\d+)$/.exec(id)
    if (!match) {
      return
    }
    const numericId = Number.parseInt(match[1]!, 10)
    if (Number.isFinite(numericId)) {
      this.nextStyleId = Math.max(this.nextStyleId, numericId + 1)
    }
  }

  bumpFormatId(id: string): void {
    const match = /^format-(\d+)$/.exec(id)
    if (!match) {
      return
    }
    const numericId = Number.parseInt(match[1]!, 10)
    if (Number.isFinite(numericId)) {
      this.nextFormatId = Math.max(this.nextFormatId, numericId + 1)
    }
  }

  bumpAxisId(axis: WorkbookAxis, id: string): void {
    const prefix = axis === 'row' ? 'row' : 'column'
    const match = new RegExp(`^${prefix}-(\\d+)$`).exec(id)
    if (!match) {
      return
    }
    const numericId = Number.parseInt(match[1]!, 10)
    if (!Number.isFinite(numericId)) {
      return
    }
    if (axis === 'row') {
      this.nextRowAxisId = Math.max(this.nextRowAxisId, numericId + 1)
      return
    }
    this.nextColumnAxisId = Math.max(this.nextColumnAxisId, numericId + 1)
  }

  createAxisEntry(axis: WorkbookAxis): WorkbookAxisEntryRecord {
    return {
      id: this.createAxisId(axis),
      size: null,
      hidden: null,
      filterHidden: null,
    }
  }

  createLogicalAxisId(axis: WorkbookAxis): string {
    return axis === 'row' ? `lr${this.nextLogicalRowAxisId++}` : `lc${this.nextLogicalColumnAxisId++}`
  }

  createLogicalAxisIds(axis: WorkbookAxis, count: number): string[] {
    if (count <= 0) {
      return []
    }
    const ids: string[] = []
    ids.length = count
    if (axis === 'row') {
      let next = this.nextLogicalRowAxisId
      this.nextLogicalRowAxisId += count
      for (let index = 0; index < count; index += 1) {
        ids[index] = `lr${next++}`
      }
      return ids
    }
    let next = this.nextLogicalColumnAxisId
    this.nextLogicalColumnAxisId += count
    for (let index = 0; index < count; index += 1) {
      ids[index] = `lc${next++}`
    }
    return ids
  }

  private createAxisId(axis: WorkbookAxis): string {
    return axis === 'row' ? `row-${this.nextRowAxisId++}` : `column-${this.nextColumnAxisId++}`
  }
}
