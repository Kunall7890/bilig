import type { Viewport } from '@bilig/protocol'
import type { GridRenderTileTextRun } from './render-tile-source.js'

export interface TextOverflowSourceV3 {
  readonly sheetOrdinal: number
  readonly row: number
  readonly col: number
  readonly spillColEnd: number
}

export class TextOverflowIndexV3 {
  private readonly sources = new Map<string, TextOverflowSourceV3>()
  private readonly cellDependencies = new Map<string, Set<string>>()
  private readonly columnDependencies = new Map<string, Set<string>>()

  replaceTileRuns(input: {
    readonly sheetOrdinal: number
    readonly viewport: Viewport
    readonly textRuns: readonly GridRenderTileTextRun[]
  }): void {
    this.clearSourcesInViewport(input.sheetOrdinal, input.viewport)
    input.textRuns.forEach((run) => {
      if (run.row === undefined || run.col === undefined || run.spillColEnd === undefined || run.spillColEnd <= run.col) {
        return
      }
      this.updateSpill({
        col: run.col,
        row: run.row,
        sheetOrdinal: input.sheetOrdinal,
        spillColEnd: run.spillColEnd,
      })
    })
  }

  updateSpill(source: TextOverflowSourceV3): void {
    const normalized = {
      ...source,
      spillColEnd: Math.max(source.col, source.spillColEnd),
    }
    const id = sourceId(normalized)
    this.deleteSource(id)
    this.sources.set(id, normalized)
    for (let col = normalized.col; col <= normalized.spillColEnd; col += 1) {
      addDependency(this.columnDependencies, columnDependencyId(normalized.sheetOrdinal, col), id)
      if (col > normalized.col) {
        addDependency(this.cellDependencies, cellDependencyId(normalized.sheetOrdinal, normalized.row, col), id)
      }
    }
  }

  markDependenciesForCell(
    input: { readonly sheetOrdinal: number; readonly row: number; readonly col: number },
    callback: (source: TextOverflowSourceV3) => void,
  ): void {
    const ids = new Set<string>()
    const directSource = this.sources.get(sourceId({ ...input, spillColEnd: input.col }))
    if (directSource) {
      ids.add(sourceId(directSource))
    }
    this.cellDependencies.get(cellDependencyId(input.sheetOrdinal, input.row, input.col))?.forEach((id) => ids.add(id))
    ids.forEach((id) => {
      const source = this.sources.get(id)
      if (source) {
        callback(source)
      }
    })
  }

  markDependenciesForCellRange(
    input: {
      readonly sheetOrdinal: number
      readonly rowStart: number
      readonly rowEnd: number
      readonly colStart: number
      readonly colEnd: number
    },
    callback: (source: TextOverflowSourceV3) => void,
  ): void {
    const ids = new Set<string>()
    for (let col = input.colStart; col <= input.colEnd; col += 1) {
      this.columnDependencies.get(columnDependencyId(input.sheetOrdinal, col))?.forEach((id) => {
        const source = this.sources.get(id)
        if (!source || source.row < input.rowStart || source.row > input.rowEnd) {
          return
        }
        ids.add(id)
      })
    }
    ids.forEach((id) => {
      const source = this.sources.get(id)
      if (source) {
        callback(source)
      }
    })
  }

  markDependenciesForAxisX(
    input: { readonly sheetOrdinal: number; readonly colStart: number; readonly colEnd: number },
    callback: (source: TextOverflowSourceV3) => void,
  ): void {
    const ids = new Set<string>()
    for (let col = input.colStart; col <= input.colEnd; col += 1) {
      this.columnDependencies.get(columnDependencyId(input.sheetOrdinal, col))?.forEach((id) => ids.add(id))
    }
    ids.forEach((id) => {
      const source = this.sources.get(id)
      if (source) {
        callback(source)
      }
    })
  }

  clear(): void {
    this.sources.clear()
    this.cellDependencies.clear()
    this.columnDependencies.clear()
  }

  private clearSourcesInViewport(sheetOrdinal: number, viewport: Viewport): void {
    const ids: string[] = []
    this.sources.forEach((source, id) => {
      if (
        source.sheetOrdinal === sheetOrdinal &&
        source.row >= viewport.rowStart &&
        source.row <= viewport.rowEnd &&
        source.col >= viewport.colStart &&
        source.col <= viewport.colEnd
      ) {
        ids.push(id)
      }
    })
    ids.forEach((id) => this.deleteSource(id))
  }

  private deleteSource(id: string): void {
    const source = this.sources.get(id)
    if (!source) {
      return
    }
    this.sources.delete(id)
    for (let col = source.col; col <= source.spillColEnd; col += 1) {
      removeDependency(this.columnDependencies, columnDependencyId(source.sheetOrdinal, col), id)
      if (col > source.col) {
        removeDependency(this.cellDependencies, cellDependencyId(source.sheetOrdinal, source.row, col), id)
      }
    }
  }
}

function addDependency(map: Map<string, Set<string>>, key: string, source: string): void {
  const values = map.get(key) ?? new Set<string>()
  values.add(source)
  map.set(key, values)
}

function removeDependency(map: Map<string, Set<string>>, key: string, source: string): void {
  const values = map.get(key)
  if (!values) {
    return
  }
  values.delete(source)
  if (values.size === 0) {
    map.delete(key)
  }
}

function sourceId(source: TextOverflowSourceV3): string {
  return `${source.sheetOrdinal}:${source.row}:${source.col}`
}

function cellDependencyId(sheetOrdinal: number, row: number, col: number): string {
  return `${sheetOrdinal}:${row}:${col}`
}

function columnDependencyId(sheetOrdinal: number, col: number): string {
  return `${sheetOrdinal}:${col}`
}
