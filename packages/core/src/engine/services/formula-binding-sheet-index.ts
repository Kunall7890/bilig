import type { CompiledFormula } from '@bilig/formula'
import {
  appendTrackedReverseEdge,
  parseQualifiedDependencySheetName,
  removeTrackedReverseEdge,
} from './formula-binding-dependency-helpers.js'

export interface FormulaBindingSheetIndex {
  readonly clear: () => void
  readonly trackFormula: (cellIndex: number, ownerSheetName: string, compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'>) => void
  readonly trackFormulaOwnerRun: (ownerSheetName: string, cellIndices: readonly number[] | Uint32Array) => void
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
  readonly moveOwnerSheetName: (oldSheetName: string, newSheetName: string) => ReadonlySet<number>
  readonly moveReferenceSheetName: (oldSheetName: string, newSheetName: string) => ReadonlySet<number>
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
  const ownerSheetRuns = new Map<string, (readonly number[] | Uint32Array)[]>()
  const referencedSheetCells = new Map<string, Set<number>>()

  const materializeOwnerSheetSet = (sheetName: string): Set<number> | undefined => {
    const runs = ownerSheetRuns.get(sheetName)
    let owners = ownerSheetCells.get(sheetName)
    if (!runs || runs.length === 0) {
      return owners
    }
    if (!owners) {
      owners = new Set()
      ownerSheetCells.set(sheetName, owners)
    }
    for (const run of runs) {
      for (let index = 0; index < run.length; index += 1) {
        owners.add(run[index]!)
      }
    }
    ownerSheetRuns.delete(sheetName)
    return owners
  }

  const removeOwner = (sheetName: string, cellIndex: number): void => {
    materializeOwnerSheetSet(sheetName)
    removeTrackedReverseEdge(ownerSheetCells, sheetName, cellIndex)
  }

  const moveOwnerSheetName = (oldSheetName: string, newSheetName: string): ReadonlySet<number> => {
    const ownerRuns = ownerSheetRuns.get(oldSheetName)
    if (ownerRuns && ownerRuns.length > 0) {
      const existingOwnerRuns = ownerSheetRuns.get(newSheetName)
      if (existingOwnerRuns) {
        existingOwnerRuns.push(...ownerRuns)
      } else {
        ownerSheetRuns.set(newSheetName, ownerRuns)
      }
      ownerSheetRuns.delete(oldSheetName)
    }
    return moveSet(ownerSheetCells, oldSheetName, newSheetName)
  }

  const moveReferenceSheetName = (oldSheetName: string, newSheetName: string): ReadonlySet<number> => {
    return moveSet(referencedSheetCells, oldSheetName, newSheetName)
  }

  return {
    clear() {
      ownerSheetCells.clear()
      ownerSheetRuns.clear()
      referencedSheetCells.clear()
    },
    trackFormula(cellIndex, ownerSheetName, compiled) {
      appendTrackedReverseEdge(ownerSheetCells, ownerSheetName, cellIndex)
      referencedSheetsForCompiled(compiled).forEach((sheetName) => {
        appendTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
      })
    },
    trackFormulaOwnerRun(ownerSheetName, cellIndices) {
      const runs = ownerSheetRuns.get(ownerSheetName)
      if (runs) {
        runs.push(cellIndices)
      } else {
        ownerSheetRuns.set(ownerSheetName, [cellIndices])
      }
    },
    untrackFormula(cellIndex, ownerSheetName, compiled) {
      if (ownerSheetName) {
        removeOwner(ownerSheetName, cellIndex)
      }
      if (!compiled) {
        return
      }
      referencedSheetsForCompiled(compiled).forEach((sheetName) => {
        removeTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
      })
    },
    moveOwnerSheetName(oldSheetName, newSheetName) {
      return moveOwnerSheetName(oldSheetName, newSheetName)
    },
    moveReferenceSheetName(oldSheetName, newSheetName) {
      return moveReferenceSheetName(oldSheetName, newSheetName)
    },
    moveSheetName(oldSheetName, newSheetName) {
      return {
        owners: moveOwnerSheetName(oldSheetName, newSheetName),
        references: moveReferenceSheetName(oldSheetName, newSheetName),
      }
    },
    getOwnedBySheetSet(sheetName) {
      return materializeOwnerSheetSet(sheetName)
    },
    getReferencingSheetSet(sheetName) {
      return referencedSheetCells.get(sheetName)
    },
    appendOwner(sheetName, cellIndex) {
      appendTrackedReverseEdge(ownerSheetCells, sheetName, cellIndex)
    },
    removeOwner(sheetName, cellIndex) {
      removeOwner(sheetName, cellIndex)
    },
    appendReference(sheetName, cellIndex) {
      appendTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
    },
    removeReference(sheetName, cellIndex) {
      removeTrackedReverseEdge(referencedSheetCells, sheetName, cellIndex)
    },
    collectOwnedBySheet(sheetName) {
      return [...(materializeOwnerSheetSet(sheetName) ?? [])]
    },
    collectReferencingSheet(sheetName) {
      return [...(referencedSheetCells.get(sheetName) ?? [])]
    },
  }
}
