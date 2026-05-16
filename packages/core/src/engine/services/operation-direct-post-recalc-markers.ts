import { ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { makeCellEntity } from '../../entity-ids.js'
import type { RuntimeDirectLookupDescriptor, RuntimeDirectScalarDescriptor } from '../runtime-state.js'
import type { DirectFormulaIndexCollection, DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import type { OperationDirectLookupCurrentService } from './operation-direct-lookup-current.js'
import {
  directScalarDeltaFromNumbers,
  directScalarDeltaFromValues,
  directScalarValueNumber,
  singleInputAffineDirectScalar,
} from './direct-scalar-helpers.js'

export interface OperationDirectPostRecalcFormula {
  readonly compiled: {
    readonly deps: readonly string[]
    readonly volatile: boolean
    readonly producesSpill: boolean
  }
  readonly directAggregate: object | undefined
  readonly directCriteria: object | undefined
  readonly directLookup: RuntimeDirectLookupDescriptor | undefined
  readonly directScalar: RuntimeDirectScalarDescriptor | undefined
}

export interface OperationDirectPostRecalcMarkerState {
  readonly workbook: {
    readonly cellStore: {
      readonly flags: ArrayLike<number | undefined>
      readonly tags: ArrayLike<ValueTag | undefined>
      readonly getValue: (cellIndex: number, readString: (stringId: number) => string) => CellValue
    }
  }
  readonly formulas: {
    readonly get: (cellIndex: number) => OperationDirectPostRecalcFormula | undefined
  }
  readonly strings: {
    readonly get: (id: number) => string
  }
}

export function createOperationDirectPostRecalcMarkers(args: {
  readonly state: OperationDirectPostRecalcMarkerState
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly getEntityDependents: (entityId: number) => Uint32Array
  readonly hasNoCellDependents: (cellIndex: number) => boolean
  readonly canSkipDirectFormulaColumnVersion: (cellIndex: number) => boolean
  readonly readDirectScalarCellNumber: (cellIndex: number) => number | undefined
  readonly directScalarCellNumericValue: (cellIndex: number) => number | undefined
  readonly directScalarCurrentResultMatchesCell: (formulaCellIndex: number, result: DirectScalarCurrentOperand) => boolean
  readonly lookupCurrent: Pick<
    OperationDirectLookupCurrentService,
    | 'tryDirectUniformLookupCurrentResult'
    | 'tryDirectUniformLookupCurrentResultFromNumeric'
    | 'canEvaluateDirectUniformLookupCurrentResultFromNumeric'
    | 'tryDirectExactLookupCurrentResult'
  >
  readonly scalarDeltaClosureLimit: number
}) {
  const tryDirectScalarNumericDelta = (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
  ): number | undefined => {
    return directScalarDeltaFromValues(directScalar, changedCellIndex, oldValue, newValue, args.readDirectScalarCellNumber)
  }

  const tryDirectScalarNumericDeltaFromNumbers = (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    oldChangedNumber: number,
    newChangedNumber: number,
  ): number | undefined => {
    return directScalarDeltaFromNumbers(directScalar, changedCellIndex, oldChangedNumber, newChangedNumber, args.readDirectScalarCellNumber)
  }

  const addDirectLookupCurrentResultIfChanged = (
    formulaCellIndex: number,
    result: DirectScalarCurrentOperand,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): boolean => {
    if (!args.directScalarCurrentResultMatchesCell(formulaCellIndex, result)) {
      postRecalcDirectFormulaIndices.addCurrentResult(formulaCellIndex, result)
      return true
    }
    return false
  }

  const canUseDirectFormulaPostRecalc = (formulaCellIndex: number): boolean => {
    const formula = args.state.formulas.get(formulaCellIndex)
    return (
      formula !== undefined &&
      (formula.directLookup !== undefined ||
        formula.directAggregate !== undefined ||
        formula.directScalar !== undefined ||
        formula.directCriteria !== undefined) &&
      args.hasNoCellDependents(formulaCellIndex)
    )
  }

  const markPostRecalcDirectLookupCurrentDependentsFromNumeric = (
    cellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): boolean => {
    const singleDependent = args.getSingleEntityDependent(makeCellEntity(cellIndex))
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      if (!canUseDirectFormulaPostRecalc(singleDependent)) {
        return false
      }
      const directLookupResult = args.lookupCurrent.tryDirectUniformLookupCurrentResultFromNumeric(
        singleDependent,
        exactLookupValue,
        approximateLookupValue,
      )
      if (directLookupResult === undefined) {
        return false
      }
      addDirectLookupCurrentResultIfChanged(singleDependent, directLookupResult, postRecalcDirectFormulaIndices)
      return true
    }

    const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
        return false
      }
      if (
        !args.lookupCurrent.canEvaluateDirectUniformLookupCurrentResultFromNumeric(
          formulaCellIndex,
          exactLookupValue,
          approximateLookupValue,
        )
      ) {
        return false
      }
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directLookupResult = args.lookupCurrent.tryDirectUniformLookupCurrentResultFromNumeric(
        formulaCellIndex,
        exactLookupValue,
        approximateLookupValue,
      )
      if (directLookupResult === undefined) {
        return false
      }
      addDirectLookupCurrentResultIfChanged(formulaCellIndex, directLookupResult, postRecalcDirectFormulaIndices)
    }
    return true
  }

  const markPostRecalcDirectFormulaDependents = (
    cellIndex: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    oldValue?: CellValue,
    newValue?: CellValue,
  ): boolean => {
    const singleDependent = args.getSingleEntityDependent(makeCellEntity(cellIndex))
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      if (!canUseDirectFormulaPostRecalc(singleDependent)) {
        return false
      }
      if (oldValue === undefined || newValue === undefined) {
        postRecalcDirectFormulaIndices.add(singleDependent)
        return true
      }
      const directScalar = args.state.formulas.get(singleDependent)?.directScalar
      if (directScalar !== undefined) {
        const delta = tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
        if (delta !== undefined) {
          postRecalcDirectFormulaIndices.addScalarDelta(singleDependent, delta)
          return true
        }
      }
      const directLookupResult = args.lookupCurrent.tryDirectUniformLookupCurrentResult(singleDependent)
      if (directLookupResult !== undefined) {
        if (!addDirectLookupCurrentResultIfChanged(singleDependent, directLookupResult, postRecalcDirectFormulaIndices)) {
          postRecalcDirectFormulaIndices.markDirectFormulaInputCovered(cellIndex)
        }
        return true
      }
      const directLookup = args.state.formulas.get(singleDependent)?.directLookup
      if (directLookup?.kind === 'exact' && directLookup.operandCellIndex === cellIndex) {
        const exactResult = args.lookupCurrent.tryDirectExactLookupCurrentResult(directLookup, newValue)
        if (exactResult !== undefined) {
          if (!addDirectLookupCurrentResultIfChanged(singleDependent, exactResult, postRecalcDirectFormulaIndices)) {
            postRecalcDirectFormulaIndices.markDirectFormulaInputCovered(cellIndex)
          }
          return true
        }
      }
      postRecalcDirectFormulaIndices.add(singleDependent)
      return true
    }
    const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
    for (let index = 0; index < dependents.length; index += 1) {
      if (!canUseDirectFormulaPostRecalc(dependents[index]!)) {
        return false
      }
    }
    if (oldValue !== undefined && newValue !== undefined && dependents.length > 16) {
      let canUseBulkScalarDeltas = true
      let commonDelta: number | undefined
      let allDeltasMatch = true
      for (let index = 0; index < dependents.length; index += 1) {
        const formulaCellIndex = dependents[index]!
        const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
        const delta = directScalar === undefined ? undefined : tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
        if (delta === undefined) {
          canUseBulkScalarDeltas = false
          break
        }
        if (commonDelta === undefined) {
          commonDelta = delta
        } else if (!Object.is(commonDelta, delta)) {
          allDeltasMatch = false
        }
      }
      if (canUseBulkScalarDeltas) {
        if (allDeltasMatch && commonDelta !== undefined) {
          postRecalcDirectFormulaIndices.appendConstantDelta(dependents, commonDelta, 'scalar')
        } else {
          const deltaCellIndices: number[] = []
          const deltas: number[] = []
          for (let index = 0; index < dependents.length; index += 1) {
            const formulaCellIndex = dependents[index]!
            const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
            const delta = directScalar === undefined ? undefined : tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
            if (delta === undefined) {
              canUseBulkScalarDeltas = false
              break
            }
            deltaCellIndices.push(formulaCellIndex)
            deltas.push(delta)
          }
          if (!canUseBulkScalarDeltas) {
            return false
          }
          postRecalcDirectFormulaIndices.appendDeltas(deltaCellIndices, deltas, 'scalar')
        }
        return true
      }
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
      if (oldValue === undefined || newValue === undefined) {
        postRecalcDirectFormulaIndices.add(formulaCellIndex)
        continue
      }
      if (directScalar !== undefined) {
        const delta = tryDirectScalarNumericDelta(directScalar, cellIndex, oldValue, newValue)
        if (delta !== undefined) {
          postRecalcDirectFormulaIndices.addScalarDelta(formulaCellIndex, delta)
          continue
        }
      }
      const directLookupResult = args.lookupCurrent.tryDirectUniformLookupCurrentResult(formulaCellIndex)
      if (directLookupResult !== undefined) {
        addDirectLookupCurrentResultIfChanged(formulaCellIndex, directLookupResult, postRecalcDirectFormulaIndices)
        continue
      }
      const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
      if (directLookup?.kind === 'exact' && directLookup.operandCellIndex === cellIndex) {
        const exactResult = args.lookupCurrent.tryDirectExactLookupCurrentResult(directLookup, newValue)
        if (exactResult !== undefined) {
          addDirectLookupCurrentResultIfChanged(formulaCellIndex, exactResult, postRecalcDirectFormulaIndices)
          continue
        }
      }
      postRecalcDirectFormulaIndices.add(formulaCellIndex)
    }
    return true
  }

  const markPostRecalcDirectScalarNumericDependents = (
    cellIndex: number,
    oldNumber: number,
    newNumber: number,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
    exactLookupValue?: number,
    approximateLookupValue?: number,
  ): boolean => {
    const singleDependent = args.getSingleEntityDependent(makeCellEntity(cellIndex))
    if (singleDependent === -1) {
      return true
    }
    if (singleDependent >= 0) {
      if (!canUseDirectFormulaPostRecalc(singleDependent)) {
        return false
      }
      const directScalar = args.state.formulas.get(singleDependent)?.directScalar
      const delta =
        directScalar === undefined ? undefined : tryDirectScalarNumericDeltaFromNumbers(directScalar, cellIndex, oldNumber, newNumber)
      if (delta === undefined) {
        const directLookupResult = args.lookupCurrent.tryDirectUniformLookupCurrentResultFromNumeric(
          singleDependent,
          exactLookupValue,
          approximateLookupValue,
        )
        if (directLookupResult === undefined) {
          return false
        }
        if (!addDirectLookupCurrentResultIfChanged(singleDependent, directLookupResult, postRecalcDirectFormulaIndices)) {
          postRecalcDirectFormulaIndices.markDirectFormulaInputCovered(cellIndex)
        }
        return true
      }
      postRecalcDirectFormulaIndices.addScalarDelta(singleDependent, delta)
      return true
    }

    const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
    if (dependents.length === 0) {
      return true
    }
    let commonDelta: number | undefined
    let allDeltasMatch = true
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
        return false
      }
      const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
      const delta =
        directScalar === undefined ? undefined : tryDirectScalarNumericDeltaFromNumbers(directScalar, cellIndex, oldNumber, newNumber)
      if (delta === undefined) {
        return false
      }
      if (commonDelta === undefined) {
        commonDelta = delta
      } else if (!Object.is(commonDelta, delta)) {
        allDeltasMatch = false
        break
      }
    }
    if (allDeltasMatch && commonDelta !== undefined) {
      postRecalcDirectFormulaIndices.appendConstantDelta(dependents, commonDelta, 'scalar')
      return true
    }
    const deltas: number[] = []
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (!canUseDirectFormulaPostRecalc(formulaCellIndex)) {
        return false
      }
      const directScalar = args.state.formulas.get(formulaCellIndex)?.directScalar
      const delta =
        directScalar === undefined ? undefined : tryDirectScalarNumericDeltaFromNumbers(directScalar, cellIndex, oldNumber, newNumber)
      if (delta === undefined) {
        return false
      }
      deltas[index] = delta
    }
    postRecalcDirectFormulaIndices.appendDeltas(dependents, deltas, 'scalar')
    return true
  }

  const tryMarkDirectScalarLinearDeltaClosure = (
    rootCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): boolean => {
    const oldRootNumber = directScalarValueNumber(oldValue)
    const newRootNumber = directScalarValueNumber(newValue)
    if (oldRootNumber === undefined || newRootNumber === undefined) {
      return false
    }
    let currentCellIndex = rootCellIndex
    let oldNumber = oldRootNumber
    let newNumber = newRootNumber
    let closureCount = 0
    const cellIndices: number[] = []
    let deltas: number[] | undefined
    let commonDelta: number | undefined
    let canUseValidatedTerminalWrites = true
    for (;;) {
      if (closureCount > args.scalarDeltaClosureLimit) {
        return false
      }
      const formulaCellIndex = args.getSingleEntityDependent(makeCellEntity(currentCellIndex))
      if (formulaCellIndex === -1) {
        break
      }
      if (formulaCellIndex < 0) {
        return false
      }
      if (formulaCellIndex === rootCellIndex) {
        return false
      }
      const formula = args.state.formulas.get(formulaCellIndex)
      if (
        !formula ||
        formula.directScalar === undefined ||
        formula.compiled.volatile ||
        formula.compiled.producesSpill ||
        ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
      ) {
        return false
      }
      const affine = singleInputAffineDirectScalar(formula.directScalar, currentCellIndex)
      let formulaOldNumber: number | undefined
      let formulaNewNumber: number | undefined
      let formulaDelta: number | undefined
      if (affine === null) {
        formulaOldNumber = args.directScalarCellNumericValue(formulaCellIndex)
        formulaDelta = tryDirectScalarNumericDeltaFromNumbers(formula.directScalar, currentCellIndex, oldNumber, newNumber)
      } else {
        if (args.state.workbook.cellStore.tags[formulaCellIndex] !== ValueTag.Number) {
          return false
        }
        formulaOldNumber = oldNumber * affine.scale + affine.offset
        formulaNewNumber = newNumber * affine.scale + affine.offset
        formulaDelta = formulaNewNumber - formulaOldNumber
      }
      if (formulaOldNumber === undefined) {
        return false
      }
      if (formulaDelta === undefined) {
        return false
      }
      if (canUseValidatedTerminalWrites && !args.canSkipDirectFormulaColumnVersion(formulaCellIndex)) {
        canUseValidatedTerminalWrites = false
      }
      if (commonDelta === undefined) {
        commonDelta = formulaDelta
      } else if (!Object.is(commonDelta, formulaDelta) && deltas === undefined) {
        deltas = []
        for (let index = 0; index < cellIndices.length; index += 1) {
          deltas[index] = commonDelta
        }
      }
      cellIndices.push(formulaCellIndex)
      if (deltas) {
        deltas.push(formulaDelta)
      }
      currentCellIndex = formulaCellIndex
      oldNumber = formulaOldNumber
      newNumber = formulaNewNumber ?? formulaOldNumber + formulaDelta
      closureCount += 1
    }
    if (cellIndices.length === 0) {
      return false
    }
    if (deltas) {
      postRecalcDirectFormulaIndices.appendDeltas(cellIndices, deltas, 'scalar')
    } else if (commonDelta !== undefined) {
      postRecalcDirectFormulaIndices.appendConstantDelta(cellIndices, commonDelta, 'scalar')
    }
    if (canUseValidatedTerminalWrites) {
      postRecalcDirectFormulaIndices.markScalarDeltaCellsValidated()
    }
    return true
  }

  const markDirectScalarDeltaClosure = (
    rootCellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  ): void => {
    const rootDependent = args.getSingleEntityDependent(makeCellEntity(rootCellIndex))
    if (rootDependent < 0 || postRecalcDirectFormulaIndices.hasDelta(rootDependent)) {
      return
    }
    if (args.getSingleEntityDependent(makeCellEntity(rootDependent)) === -1) {
      return
    }
    if (tryMarkDirectScalarLinearDeltaClosure(rootCellIndex, oldValue, newValue, postRecalcDirectFormulaIndices)) {
      return
    }
    const pending: Array<{ cellIndex: number; oldValue: CellValue; newValue: CellValue }> = [
      { cellIndex: rootCellIndex, oldValue, newValue },
    ]
    const closureDeltas = new Map<number, number>()
    const visited = new Set<number>([rootCellIndex])
    for (let cursor = 0; cursor < pending.length; cursor += 1) {
      if (closureDeltas.size > args.scalarDeltaClosureLimit) {
        return
      }
      const current = pending[cursor]!
      const dependents = args.getEntityDependents(makeCellEntity(current.cellIndex))
      for (let index = 0; index < dependents.length; index += 1) {
        const formulaCellIndex = dependents[index]!
        if (visited.has(formulaCellIndex)) {
          return
        }
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          formula.compiled.volatile ||
          formula.compiled.producesSpill ||
          ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
        ) {
          return
        }
        const formulaDelta = tryDirectScalarNumericDelta(formula.directScalar, current.cellIndex, current.oldValue, current.newValue)
        if (formulaDelta === undefined) {
          return
        }
        const formulaOldValue = args.state.workbook.cellStore.getValue(formulaCellIndex, (id) => args.state.strings.get(id))
        if (formulaOldValue.tag !== ValueTag.Number) {
          return
        }
        const accumulatedDelta = (closureDeltas.get(formulaCellIndex) ?? 0) + formulaDelta
        closureDeltas.set(formulaCellIndex, accumulatedDelta)
        visited.add(formulaCellIndex)
        pending.push({
          cellIndex: formulaCellIndex,
          oldValue: formulaOldValue,
          newValue: { tag: ValueTag.Number, value: formulaOldValue.value + accumulatedDelta },
        })
      }
    }
    closureDeltas.forEach((delta, formulaCellIndex) => {
      postRecalcDirectFormulaIndices.addScalarDelta(formulaCellIndex, delta)
    })
  }

  return {
    canUseDirectFormulaPostRecalc,
    markDirectScalarDeltaClosure,
    markPostRecalcDirectFormulaDependents,
    markPostRecalcDirectLookupCurrentDependentsFromNumeric,
    markPostRecalcDirectScalarNumericDependents,
    tryDirectScalarNumericDelta,
    tryDirectScalarNumericDeltaFromNumbers,
    tryMarkDirectScalarLinearDeltaClosure,
  } as const
}
