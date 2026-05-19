import type { CompiledFormula } from '@bilig/formula'
import {
  appendTrackedReverseEdge,
  parseQualifiedDependencySheetName,
  removeTrackedReverseEdge,
} from './formula-binding-dependency-helpers.js'

export interface FormulaBindingSheetIndex {
  readonly clear: () => void
  readonly trackFormula: (cellIndex: number, ownerSheetName: string, compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'>) => void
  readonly untrackFormula: (
    cellIndex: number,
    ownerSheetName: string | undefined,
    compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'> | undefined,
  ) => void
  readonly moveSheetName: (
    oldSheetName: string,
    newSheetName: string,
  ) => {
    readonly owners: ReadonlySet<number>
    readonly references: ReadonlySet<number>
  }
  readonly getOwnedBySheetSet: (sheetName: string) => Set<number> | undefined
  readonly getReferencingSheetSet: (sheetName: string) => Set<number> | undefined
  readonly appendOwner: (sheetName: string, cellIndex: number) => void
  readonly removeOwner: (sheetName: string, cellIndex: number) => void
  readonly appendReference: (sheetName: string, cellIndex: number) => void
  readonly removeReference: (sheetName: string, cellIndex: number) => void
  readonly collectOwnedBySheet: (sheetName: string) => number[]
  readonly collectReferencingSheet: (sheetName: string) => number[]
}

function referencedSheetsForCompiled(compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'>): string[] {
  const parsedDeps = compiled.parsedDeps
  let sheets: Set<string> | undefined
  if (parsedDeps !== undefined && parsedDeps.length === compiled.deps.length) {
    for (let index = 0; index < parsedDeps.length; index += 1) {
      const dependency = parsedDeps[index]!
      if (dependency.sheetName) {
        ;(sheets ??= new Set()).add(dependency.sheetName)
      }
      if (dependency.kind === 'range' && dependency.sheetEndName) {
        ;(sheets ??= new Set()).add(dependency.sheetEndName)
      }
    }
    return sheets === undefined ? [] : [...sheets]
  }
  sheets = new Set<string>()
  parsedDeps?.forEach((dependency) => {
    if (dependency.sheetName) {
      sheets.add(dependency.sheetName)
    }
    if (dependency.kind === 'range' && dependency.sheetEndName) {
      sheets.add(dependency.sheetEndName)
    }
  })
  compiled.deps.forEach((dependency, index) => {
    const parsedDependency = compiled.parsedDeps?.[index]
    if (parsedDependency?.kind === 'range' && parsedDependency.sheetEndName) {
      return
    }
    const sheetName = parseQualifiedDependencySheetName(dependency)
    if (sheetName) {
      sheets.add(sheetName)
    }
  })
  return [...sheets]
}

const EMPTY_MOVED_CELL_SET: ReadonlySet<number> = new Set()

function moveSet<Key extends string>(registry: Map<Key, Set<number>>, oldKey: Key, newKey: Key): ReadonlySet<number> {
  const candidates = registry.get(oldKey)
  if (!candidates || candidates.size === 0) {
    return EMPTY_MOVED_CELL_SET
  }
  const existing = registry.get(newKey)
  if (existing) {
    candidates.forEach((cellIndex) => {
      existing.add(cellIndex)
    })
  } else {
    registry.set(newKey, candidates)
  }
  registry.delete(oldKey)
  return candidates
}

export function createFormulaBindingSheetIndex(): FormulaBindingSheetIndex {
  const ownerSheetCells = new Map<string, Set<number>>()
  const referencedSheetCells = new Map<string, Set<number>>()

  return {
    clear() {
      ownerSheetCells.clear()
      referencedSheetCells.clear()
    },
    trackFormula(cellIndex, ownerSheetName, compiled) {
      appendTrackedReverseEdge(ownerSheetCells, ownerSheetName, cellIndex)
      referencedSheetsForCompiled(compiled).forEach((sheetName) => {
        appendTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
      })
    },
    untrackFormula(cellIndex, ownerSheetName, compiled) {
      if (ownerSheetName) {
        removeTrackedReverseEdge(ownerSheetCells, ownerSheetName, cellIndex)
      }
      if (!compiled) {
        return
      }
      referencedSheetsForCompiled(compiled).forEach((sheetName) => {
        removeTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
      })
    },
    moveSheetName(oldSheetName, newSheetName) {
      return {
        owners: moveSet(ownerSheetCells, oldSheetName, newSheetName),
        references: moveSet(referencedSheetCells, oldSheetName, newSheetName),
      }
    },
    getOwnedBySheetSet(sheetName) {
      return ownerSheetCells.get(sheetName)
    },
    getReferencingSheetSet(sheetName) {
      return referencedSheetCells.get(sheetName)
    },
    appendOwner(sheetName, cellIndex) {
      appendTrackedReverseEdge(ownerSheetCells, sheetName, cellIndex)
    },
    removeOwner(sheetName, cellIndex) {
      removeTrackedReverseEdge(ownerSheetCells, sheetName, cellIndex)
    },
    appendReference(sheetName, cellIndex) {
      appendTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
    },
    removeReference(sheetName, cellIndex) {
      removeTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
    },
    collectOwnedBySheet(sheetName) {
      return [...(ownerSheetCells.get(sheetName) ?? [])]
    },
    collectReferencingSheet(sheetName) {
      return [...(referencedSheetCells.get(sheetName) ?? [])]
    },
  }
}
