export type AxisKind = 'row' | 'column'

export interface AxisEntrySnapshot {
  readonly id: string
  readonly index: number
}

const DENSE_REVERSE_INDEX_DEFER_THRESHOLD = 32

function createAxisEntrySnapshot(id: string, index: number): AxisEntrySnapshot {
  return { id, index }
}

export class AxisMap {
  private readonly entries: Array<string | undefined> = []
  private readonly idToIndex = new Map<string, number>()
  private reverseIndexDirty = false

  get(index: number): string | undefined {
    return this.entries[index]
  }

  getId(index: number): string | undefined {
    return this.get(index)
  }

  set(index: number, id: string): void {
    const previous = this.entries[index]
    if (!this.reverseIndexDirty && previous !== undefined) {
      this.idToIndex.delete(previous)
    }
    this.entries[index] = id
    if (!this.reverseIndexDirty) {
      this.idToIndex.set(id, index)
    }
  }

  setId(index: number, id: string): void {
    this.set(index, id)
  }

  ensure(index: number, createId: () => string): string {
    const existing = this.entries[index]
    if (existing !== undefined) {
      return existing
    }
    if (index >= this.entries.length) {
      this.entries.length = index + 1
    }
    const id = createId()
    this.entries[index] = id
    if (!this.reverseIndexDirty) {
      this.idToIndex.set(id, index)
    }
    return id
  }

  ensureId(index: number, createId: () => string): string {
    return this.ensure(index, createId)
  }

  ensureDenseIds(start: number, count: number, createId: () => string): string[] {
    if (count <= 0) {
      return []
    }
    const ids: string[] = []
    ids.length = count
    const end = start + count
    if (this.entries.length < end) {
      this.entries.length = end
    }
    for (let offset = 0; offset < count; offset += 1) {
      const index = start + offset
      let id = this.entries[index]
      if (id === undefined) {
        id = createId()
        this.entries[index] = id
        if (count > DENSE_REVERSE_INDEX_DEFER_THRESHOLD) {
          this.reverseIndexDirty = true
        } else if (!this.reverseIndexDirty) {
          this.idToIndex.set(id, index)
        }
      }
      ids[offset] = id
    }
    return ids
  }

  ensureDenseIdsFrom(start: number, count: number, createIds: (count: number) => readonly string[]): string[] {
    if (count <= 0) {
      return []
    }
    const end = start + count
    if (this.entries.length <= start) {
      if (this.entries.length < end) {
        this.entries.length = end
      }
      const created = createIds(count)
      if (created.length !== count) {
        throw new Error(`Expected ${String(count)} dense axis ids, got ${String(created.length)}`)
      }
      const createdIds = Array.isArray(created) ? created : [...created]
      const deferReverseIndex = count > DENSE_REVERSE_INDEX_DEFER_THRESHOLD
      if (deferReverseIndex) {
        this.reverseIndexDirty = true
      }
      for (let offset = 0; offset < count; offset += 1) {
        const index = start + offset
        const id = createdIds[offset]!
        this.entries[index] = id
        if (!deferReverseIndex && !this.reverseIndexDirty) {
          this.idToIndex.set(id, index)
        }
      }
      return createdIds
    }
    if (this.entries.length < end) {
      this.entries.length = end
    }
    let missingCount = 0
    for (let offset = 0; offset < count; offset += 1) {
      if (this.entries[start + offset] === undefined) {
        missingCount += 1
      }
    }
    const created = missingCount > 0 ? createIds(missingCount) : []
    if (created.length !== missingCount) {
      throw new Error(`Expected ${String(missingCount)} dense axis ids, got ${String(created.length)}`)
    }
    const deferReverseIndex = missingCount > DENSE_REVERSE_INDEX_DEFER_THRESHOLD
    if (deferReverseIndex) {
      this.reverseIndexDirty = true
    }
    const ids: string[] = []
    ids.length = count
    let createdIndex = 0
    for (let offset = 0; offset < count; offset += 1) {
      const index = start + offset
      let id = this.entries[index]
      if (id === undefined) {
        id = created[createdIndex]!
        createdIndex += 1
        this.entries[index] = id
        if (!deferReverseIndex && !this.reverseIndexDirty) {
          this.idToIndex.set(id, index)
        }
      }
      ids[offset] = id
    }
    return ids
  }

  indexOf(id: string): number {
    this.ensureReverseIndex()
    return this.idToIndex.get(id) ?? -1
  }

  get length(): number {
    return this.entries.length
  }

  list(): AxisEntrySnapshot[] {
    const snapshots: AxisEntrySnapshot[] = []
    for (let index = 0; index < this.entries.length; index += 1) {
      const id = this.entries[index]
      if (id === undefined) {
        continue
      }
      snapshots.push(createAxisEntrySnapshot(id, index))
    }
    return snapshots
  }

  snapshot(start: number, count: number): AxisEntrySnapshot[] {
    if (count <= 0) {
      return []
    }
    const snapshots: AxisEntrySnapshot[] = []
    for (let offset = 0; offset < count; offset += 1) {
      const index = start + offset
      const id = this.entries[index]
      if (id === undefined) {
        continue
      }
      snapshots.push(createAxisEntrySnapshot(id, index))
    }
    return snapshots
  }

  replaceRange(start: number, entries: readonly AxisEntrySnapshot[]): void {
    if (entries.length === 1) {
      const entry = entries[0]!
      if (entry.index < start) {
        return
      }
      this.entries[entry.index] = entry.id
      this.reverseIndexDirty = true
      return
    }
    for (const entry of entries) {
      if (entry.index < start) {
        continue
      }
      this.entries[entry.index] = entry.id
    }
    if (entries.length > 0) {
      this.reverseIndexDirty = true
    }
  }

  splice(start: number, deleteCount: number, entries: readonly AxisEntrySnapshot[]): AxisEntrySnapshot[]
  splice(start: number, deleteCount: number, insertCount: number, entries: readonly AxisEntrySnapshot[]): AxisEntrySnapshot[]
  splice(
    start: number,
    deleteCount: number,
    insertCountOrEntries: number | readonly AxisEntrySnapshot[],
    maybeEntries?: readonly AxisEntrySnapshot[],
  ): AxisEntrySnapshot[] {
    const entries = typeof insertCountOrEntries === 'number' ? (maybeEntries ?? []) : insertCountOrEntries
    const explicitInsertCount = typeof insertCountOrEntries === 'number' ? insertCountOrEntries : entries.length
    if (entries.length === 0 && this.entries.length <= start) {
      return []
    }
    if (this.entries.length < start) {
      this.entries.length = start
    }
    if (deleteCount === 0 && explicitInsertCount === 1 && entries.length <= 1) {
      const entry = entries[0]
      const insertedId = entry?.index === start ? entry.id : undefined
      this.entries.splice(start, 0, insertedId)
      this.reverseIndexDirty = true
      return []
    }
    const insertLength = Math.max(
      explicitInsertCount,
      entries.reduce((max, entry) => Math.max(max, entry.index - start + 1), 0),
    )
    const inserted: Array<string | undefined> = Array.from({ length: insertLength }, () => undefined)
    for (const entry of entries) {
      const offset = entry.index - start
      if (offset < 0 || offset >= insertLength) {
        continue
      }
      inserted[offset] = entry.id
    }
    const removed = this.entries.splice(start, deleteCount, ...inserted)
    this.reverseIndexDirty = true
    return removed.flatMap((id, index) => (id === undefined ? [] : [createAxisEntrySnapshot(id, start + index)]))
  }

  move(start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return
    }
    const moved = this.entries.splice(start, count)
    this.entries.splice(target, 0, ...moved)
    this.reverseIndexDirty = true
  }

  private ensureReverseIndex(): void {
    if (!this.reverseIndexDirty) {
      return
    }
    this.rebuildIndex()
  }

  private rebuildIndex(): void {
    this.idToIndex.clear()
    this.rebuildIndexFrom(0)
    this.reverseIndexDirty = false
  }

  private rebuildIndexFrom(start: number): void {
    for (let index = start; index < this.entries.length; index += 1) {
      const id = this.entries[index]
      if (id !== undefined) {
        this.idToIndex.set(id, index)
      }
    }
  }
}
