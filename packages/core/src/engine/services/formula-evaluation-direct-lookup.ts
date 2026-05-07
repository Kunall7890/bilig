import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineRuntimeState, PreparedApproximateVectorLookup, PreparedExactVectorLookup, RuntimeFormula } from '../runtime-state.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import type { SortedColumnSearchService } from './sorted-column-search-service.js'
import { directErrorResult, directNumberResult, sameExactNumericValue } from './formula-evaluation-helpers.js'

interface DirectVectorLookupContext {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly exactLookup: Pick<ExactColumnIndexService, 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly sortedLookup: Pick<SortedColumnSearchService, 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly readCellValueByIndex: (cellIndex: number | undefined) => CellValue
}

function refreshDirectExactLookup(
  context: DirectVectorLookupContext,
  directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact' }>,
): PreparedExactVectorLookup {
  const prepared = directLookup.prepared
  const lookupSheet = context.state.workbook.getSheet(prepared.sheetName)
  const currentColumnVersion = lookupSheet?.columnVersions[prepared.col] ?? 0
  const currentStructureVersion = lookupSheet?.structureVersion ?? 0
  if (currentColumnVersion === prepared.columnVersion && currentStructureVersion === prepared.structureVersion) {
    return prepared
  }
  const refreshed = context.exactLookup.prepareVectorLookup({
    sheetName: prepared.sheetName,
    rowStart: prepared.rowStart,
    rowEnd: prepared.rowEnd,
    col: prepared.col,
  })
  directLookup.prepared = refreshed
  return refreshed
}

function refreshDirectApproximateLookup(
  context: DirectVectorLookupContext,
  directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate' }>,
): PreparedApproximateVectorLookup {
  const prepared = directLookup.prepared
  const lookupSheet = context.state.workbook.getSheet(prepared.sheetName)
  const currentColumnVersion = lookupSheet?.columnVersions[prepared.col] ?? 0
  const currentStructureVersion = lookupSheet?.structureVersion ?? 0
  if (currentColumnVersion === prepared.columnVersion && currentStructureVersion === prepared.structureVersion) {
    return prepared
  }
  const refreshed = context.sortedLookup.prepareVectorLookup({
    sheetName: prepared.sheetName,
    rowStart: prepared.rowStart,
    rowEnd: prepared.rowEnd,
    col: prepared.col,
  })
  directLookup.prepared = refreshed
  return refreshed
}

function refreshDirectExactUniformLookup(
  context: DirectVectorLookupContext,
  formula: RuntimeFormula,
  directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact-uniform-numeric' }>,
):
  | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact-uniform-numeric' }>
  | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'exact' }> {
  const lookupSheet = context.state.workbook.getSheetById(directLookup.sheetId)
  const currentColumnVersion = lookupSheet?.columnVersions[directLookup.col] ?? 0
  const currentStructureVersion = lookupSheet?.structureVersion ?? 0
  if (
    currentStructureVersion === directLookup.structureVersion &&
    (currentColumnVersion === directLookup.columnVersion || currentColumnVersion === directLookup.tailPatch?.columnVersion)
  ) {
    return directLookup
  }
  const refreshed = context.exactLookup.prepareVectorLookup({
    sheetName: directLookup.sheetName,
    rowStart: directLookup.rowStart,
    rowEnd: directLookup.rowEnd,
    col: directLookup.col,
  })
  if (refreshed.comparableKind === 'numeric' && refreshed.uniformStart !== undefined && refreshed.uniformStep !== undefined) {
    directLookup.length = refreshed.length
    directLookup.columnVersion = refreshed.columnVersion
    directLookup.structureVersion = refreshed.structureVersion
    directLookup.sheetColumnVersions = refreshed.sheetColumnVersions
    directLookup.start = refreshed.uniformStart
    directLookup.step = refreshed.uniformStep
    delete directLookup.tailPatch
    return directLookup
  }
  const fallback = {
    kind: 'exact' as const,
    operandCellIndex: directLookup.operandCellIndex,
    prepared: refreshed,
    searchMode: directLookup.searchMode,
  }
  formula.directLookup = fallback
  return fallback
}

function refreshDirectApproximateUniformLookup(
  context: DirectVectorLookupContext,
  formula: RuntimeFormula,
  directLookup: Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate-uniform-numeric' }>,
):
  | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate-uniform-numeric' }>
  | Extract<NonNullable<RuntimeFormula['directLookup']>, { kind: 'approximate' }> {
  const lookupSheet = context.state.workbook.getSheetById(directLookup.sheetId)
  const currentColumnVersion = lookupSheet?.columnVersions[directLookup.col] ?? 0
  const currentStructureVersion = lookupSheet?.structureVersion ?? 0
  if (
    currentStructureVersion === directLookup.structureVersion &&
    (currentColumnVersion === directLookup.columnVersion || currentColumnVersion === directLookup.tailPatch?.columnVersion)
  ) {
    return directLookup
  }
  const refreshed = context.sortedLookup.prepareVectorLookup({
    sheetName: directLookup.sheetName,
    rowStart: directLookup.rowStart,
    rowEnd: directLookup.rowEnd,
    col: directLookup.col,
  })
  if (refreshed.comparableKind === 'numeric' && refreshed.uniformStart !== undefined && refreshed.uniformStep !== undefined) {
    directLookup.length = refreshed.length
    directLookup.columnVersion = refreshed.columnVersion
    directLookup.structureVersion = refreshed.structureVersion
    directLookup.sheetColumnVersions = refreshed.sheetColumnVersions
    directLookup.start = refreshed.uniformStart
    directLookup.step = refreshed.uniformStep
    delete directLookup.tailPatch
    return directLookup
  }
  const fallback = {
    kind: 'approximate' as const,
    operandCellIndex: directLookup.operandCellIndex,
    prepared: refreshed,
    matchMode: directLookup.matchMode,
  }
  formula.directLookup = fallback
  return fallback
}

export function tryEvaluateDirectVectorLookup(context: DirectVectorLookupContext, formula: RuntimeFormula): CellValue | undefined {
  const directLookup = formula.directLookup
  if (!directLookup) {
    return undefined
  }
  const cellStore = context.state.workbook.cellStore
  if (directLookup.kind === 'exact-uniform-numeric') {
    const refreshed = refreshDirectExactUniformLookup(context, formula, directLookup)
    if (refreshed.kind !== 'exact-uniform-numeric') {
      return tryEvaluateDirectVectorLookup(context, formula)
    }
    const tag = cellStore.tags[refreshed.operandCellIndex]
    if (tag === ValueTag.Error) {
      return undefined
    }
    if (tag !== ValueTag.Number) {
      return directErrorResult(ErrorCode.NA)
    }
    const numericValue = Object.is(cellStore.numbers[refreshed.operandCellIndex] ?? 0, -0)
      ? 0
      : (cellStore.numbers[refreshed.operandCellIndex] ?? 0)
    const tailPatch = refreshed.tailPatch
    if (tailPatch !== undefined) {
      if (sameExactNumericValue(numericValue, tailPatch.newNumeric)) {
        return directNumberResult(tailPatch.row - refreshed.rowStart + 1)
      }
      if (sameExactNumericValue(numericValue, tailPatch.oldNumeric)) {
        return directErrorResult(ErrorCode.NA)
      }
    }
    if (refreshed.step === 1) {
      if (!Number.isInteger(numericValue)) {
        return directErrorResult(ErrorCode.NA)
      }
      const position = numericValue - refreshed.start + 1
      return position >= 1 && position <= refreshed.length ? directNumberResult(position) : directErrorResult(ErrorCode.NA)
    }
    if (refreshed.step === -1) {
      if (!Number.isInteger(numericValue)) {
        return directErrorResult(ErrorCode.NA)
      }
      const position = refreshed.start - numericValue + 1
      return position >= 1 && position <= refreshed.length ? directNumberResult(position) : directErrorResult(ErrorCode.NA)
    }
    const relative = (numericValue - refreshed.start) / refreshed.step
    return Number.isInteger(relative) && relative >= 0 && relative < refreshed.length
      ? directNumberResult(relative + 1)
      : directErrorResult(ErrorCode.NA)
  }
  if (directLookup.kind === 'exact') {
    const prepared = refreshDirectExactLookup(context, directLookup)
    const cellIndex = directLookup.operandCellIndex
    const result = context.exactLookup.findPreparedVectorMatch({
      lookupValue: context.readCellValueByIndex(cellIndex),
      prepared,
      searchMode: directLookup.searchMode,
    })
    if (!result.handled) {
      return undefined
    }
    return result.position === undefined ? directErrorResult(ErrorCode.NA) : directNumberResult(result.position)
  }
  if (directLookup.kind === 'approximate-uniform-numeric') {
    const refreshed = refreshDirectApproximateUniformLookup(context, formula, directLookup)
    if (refreshed.kind !== 'approximate-uniform-numeric') {
      return tryEvaluateDirectVectorLookup(context, formula)
    }
    const tag = cellStore.tags[refreshed.operandCellIndex]
    let lookupValue = 0
    switch (tag) {
      case undefined:
      case ValueTag.Empty:
        lookupValue = 0
        break
      case ValueTag.Number:
        lookupValue = Object.is(cellStore.numbers[refreshed.operandCellIndex] ?? 0, -0)
          ? 0
          : (cellStore.numbers[refreshed.operandCellIndex] ?? 0)
        break
      case ValueTag.Boolean:
        lookupValue = (cellStore.numbers[refreshed.operandCellIndex] ?? 0) !== 0 ? 1 : 0
        break
      case ValueTag.Error:
      case ValueTag.String:
        return undefined
    }
    const lastValue = refreshed.start + refreshed.step * (refreshed.length - 1)
    const tailPatch = refreshed.tailPatch
    if (refreshed.matchMode === 1 && refreshed.step > 0) {
      if (lookupValue < refreshed.start) {
        return directErrorResult(ErrorCode.NA)
      }
      if (tailPatch !== undefined && tailPatch.row === refreshed.rowEnd && tailPatch.newNumeric > tailPatch.oldNumeric) {
        if (lookupValue >= tailPatch.newNumeric) {
          return directNumberResult(refreshed.length)
        }
        if (lookupValue >= tailPatch.oldNumeric) {
          return directNumberResult(refreshed.length - 1)
        }
      }
      if (lookupValue >= lastValue) {
        return directNumberResult(refreshed.length)
      }
      const position = Math.floor((lookupValue - refreshed.start) / refreshed.step) + 1
      return directNumberResult(Math.min(refreshed.length, Math.max(1, position)))
    }
    if (refreshed.matchMode === -1 && refreshed.step < 0) {
      if (lookupValue > refreshed.start) {
        return directErrorResult(ErrorCode.NA)
      }
      if (tailPatch !== undefined && tailPatch.row === refreshed.rowEnd && tailPatch.newNumeric < tailPatch.oldNumeric) {
        if (lookupValue <= tailPatch.newNumeric) {
          return directNumberResult(refreshed.length)
        }
        if (lookupValue <= tailPatch.oldNumeric) {
          return directNumberResult(refreshed.length - 1)
        }
      }
      if (lookupValue <= lastValue) {
        return directNumberResult(refreshed.length)
      }
      const position = Math.floor((refreshed.start - lookupValue) / -refreshed.step) + 1
      return directNumberResult(Math.min(refreshed.length, Math.max(1, position)))
    }
    return undefined
  }
  const prepared = refreshDirectApproximateLookup(context, directLookup)
  const cellIndex = directLookup.operandCellIndex
  const result = context.sortedLookup.findPreparedVectorMatch({
    lookupValue: context.readCellValueByIndex(cellIndex),
    prepared,
    matchMode: directLookup.matchMode,
  })
  if (!result.handled) {
    return undefined
  }
  return result.position === undefined ? directErrorResult(ErrorCode.NA) : directNumberResult(result.position)
}
