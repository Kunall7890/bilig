import type { CellValue } from '@bilig/protocol'
import { makeExactLookupColumnEntity, makeSortedLookupColumnEntity } from '../../entity-ids.js'
import type { RuntimeDirectLookupDescriptor } from '../runtime-state.js'
import { normalizeExactLookupKey, normalizeExactNumericValue, sameExactNumericValue } from './direct-lookup-helpers.js'

export interface OperationLookupDirtyWriteRequest {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly oldValue: CellValue
  readonly newValue: CellValue
  readonly oldStringId?: number
  readonly newStringId?: number
}

interface ExactLookupImpactEntry {
  readonly formulaCellIndex: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly operandKey: string | undefined
}

interface ExactLookupImpactCache {
  readonly entries: readonly ExactLookupImpactEntry[]
  readonly operandKeys: ReadonlySet<string>
}

export type ExactLookupImpactCaches = Map<string, ExactLookupImpactCache>

interface OperationLookupDirtyMarkerFormulaAccess {
  get(cellIndex: number): { readonly directLookup: RuntimeDirectLookupDescriptor | undefined } | undefined
}

interface OperationLookupDirtyMarkerWorkbookAccess {
  getSheet(sheetName: string): { readonly id: number } | undefined
}

export interface OperationLookupDirtyMarkerService {
  readonly markAffectedExactLookupDependents: (
    request: OperationLookupDirtyWriteRequest,
    formulaChangedCount: number,
    caches: ExactLookupImpactCaches,
  ) => number
  readonly markAffectedApproximateLookupDependents: (request: OperationLookupDirtyWriteRequest, formulaChangedCount: number) => number
  readonly noteExactLookupLiteralWriteWhenDirty: (
    request: OperationLookupDirtyWriteRequest,
    formulaChangedCount: number,
    caches: ExactLookupImpactCaches,
  ) => number
  readonly noteSortedLookupLiteralWriteWhenDirty: (request: OperationLookupDirtyWriteRequest, formulaChangedCount: number) => number
}

export function createOperationLookupDirtyMarkerService(args: {
  readonly state: {
    readonly workbook: OperationLookupDirtyMarkerWorkbookAccess
    readonly formulas: OperationLookupDirtyMarkerFormulaAccess
    readonly strings: { get(id: number): string }
  }
  readonly getEntityDependents: (entityId: number) => Uint32Array
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly readCellValueForLookup: (cellIndex: number | undefined) => { value: CellValue; stringId: number | undefined }
  readonly canSkipApproximateLookupDirtyMark: (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' | 'approximate-uniform-numeric' }>,
    request: OperationLookupDirtyWriteRequest,
  ) => boolean
  readonly noteExactLookupLiteralWrite: (request: OperationLookupDirtyWriteRequest) => void
  readonly noteSortedLookupLiteralWrite: (request: OperationLookupDirtyWriteRequest) => void
  readonly lookupImpactCacheKey: (sheetId: number, col: number) => string
}): OperationLookupDirtyMarkerService {
  const getExactLookupImpactCache = (sheetId: number, col: number, caches: ExactLookupImpactCaches): ExactLookupImpactCache => {
    const cacheKey = args.lookupImpactCacheKey(sheetId, col)
    const cached = caches.get(cacheKey)
    if (cached !== undefined) {
      return cached
    }
    const dependents = args.getEntityDependents(makeExactLookupColumnEntity(sheetId, col))
    const entries: ExactLookupImpactEntry[] = []
    const operandKeys = new Set<string>()
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
      /* v8 ignore next -- defensive guard for stale exact-lookup reverse edges. */
      if (directLookup?.kind !== 'exact' && directLookup?.kind !== 'exact-uniform-numeric') {
        continue
      }
      const rowStart = directLookup.kind === 'exact' ? directLookup.prepared.rowStart : directLookup.rowStart
      const rowEnd = directLookup.kind === 'exact' ? directLookup.prepared.rowEnd : directLookup.rowEnd
      const operand = args.readCellValueForLookup(directLookup.operandCellIndex)
      const operandKey = normalizeExactLookupKey(operand.value, (id) => args.state.strings.get(id), operand.stringId)
      if (operandKey !== undefined) {
        operandKeys.add(operandKey)
      }
      entries.push({
        formulaCellIndex,
        rowStart,
        rowEnd,
        operandKey,
      })
    }
    const cache = { entries, operandKeys }
    caches.set(cacheKey, cache)
    return cache
  }

  const markSingleNumericExactLookupImpact = (
    sheetId: number,
    request: OperationLookupDirtyWriteRequest,
    formulaChangedCount: number,
  ): number | undefined => {
    const formulaCellIndex = args.getSingleEntityDependent(makeExactLookupColumnEntity(sheetId, request.col))
    if (formulaCellIndex === -1) {
      return formulaChangedCount
    }
    if (formulaCellIndex < 0) {
      return undefined
    }
    const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
    if (directLookup?.kind !== 'exact' && directLookup?.kind !== 'exact-uniform-numeric') {
      return undefined
    }
    const rowStart = directLookup.kind === 'exact' ? directLookup.prepared.rowStart : directLookup.rowStart
    const rowEnd = directLookup.kind === 'exact' ? directLookup.prepared.rowEnd : directLookup.rowEnd
    if (request.row < rowStart || request.row > rowEnd) {
      return formulaChangedCount
    }
    const oldNumeric = normalizeExactNumericValue(request.oldValue)
    const newNumeric = normalizeExactNumericValue(request.newValue)
    if (oldNumeric === undefined || newNumeric === undefined) {
      return undefined
    }
    const operandNumeric = normalizeExactNumericValue(args.readCellValueForLookup(directLookup.operandCellIndex).value)
    if (operandNumeric === undefined) {
      return undefined
    }
    if (!sameExactNumericValue(oldNumeric, operandNumeric) && !sameExactNumericValue(newNumeric, operandNumeric)) {
      return formulaChangedCount
    }
    return args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
  }

  const markAffectedExactLookupDependents = (
    request: OperationLookupDirtyWriteRequest,
    formulaChangedCount: number,
    caches: ExactLookupImpactCaches,
  ): number => {
    const sheet = args.state.workbook.getSheet(request.sheetName)
    /* v8 ignore next -- defensive guard for stale lookup writes after sheet deletion. */
    if (!sheet) {
      return formulaChangedCount
    }
    const singleNumericImpact = markSingleNumericExactLookupImpact(sheet.id, request, formulaChangedCount)
    if (singleNumericImpact !== undefined) {
      return singleNumericImpact
    }
    const oldKey = normalizeExactLookupKey(request.oldValue, (id) => args.state.strings.get(id), request.oldStringId)
    const newKey = normalizeExactLookupKey(request.newValue, (id) => args.state.strings.get(id), request.newStringId)
    /* v8 ignore next -- error values cannot affect exact lookup matches. */
    if (oldKey === undefined && newKey === undefined) {
      return formulaChangedCount
    }
    const cache = getExactLookupImpactCache(sheet.id, request.col, caches)
    if (
      cache.entries.length === 0 ||
      ((oldKey === undefined || !cache.operandKeys.has(oldKey)) && (newKey === undefined || !cache.operandKeys.has(newKey)))
    ) {
      return formulaChangedCount
    }
    for (let index = 0; index < cache.entries.length; index += 1) {
      const entry = cache.entries[index]!
      /* v8 ignore next -- cached ranges are normally aligned with the lookup owner. */
      if (request.row < entry.rowStart || request.row > entry.rowEnd) {
        continue
      }
      /* v8 ignore next -- operand keys are prefiltered before scanning entries. */
      if (entry.operandKey === undefined || (entry.operandKey !== oldKey && entry.operandKey !== newKey)) {
        continue
      }
      formulaChangedCount = args.markFormulaChanged(entry.formulaCellIndex, formulaChangedCount)
    }
    return formulaChangedCount
  }

  const markAffectedApproximateLookupDependents = (request: OperationLookupDirtyWriteRequest, formulaChangedCount: number): number => {
    const sheet = args.state.workbook.getSheet(request.sheetName)
    if (!sheet) {
      return formulaChangedCount
    }
    const dependents = args.getEntityDependents(makeSortedLookupColumnEntity(sheet.id, request.col))
    if (dependents.length === 0) {
      return formulaChangedCount
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      const directLookup = args.state.formulas.get(formulaCellIndex)?.directLookup
      if (directLookup?.kind !== 'approximate' && directLookup?.kind !== 'approximate-uniform-numeric') {
        continue
      }
      const rowStart = directLookup.kind === 'approximate' ? directLookup.prepared.rowStart : directLookup.rowStart
      const rowEnd = directLookup.kind === 'approximate' ? directLookup.prepared.rowEnd : directLookup.rowEnd
      if (request.row < rowStart || request.row > rowEnd) {
        continue
      }
      if (args.canSkipApproximateLookupDirtyMark(directLookup, request)) {
        continue
      }
      formulaChangedCount = args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
    }
    return formulaChangedCount
  }

  const noteExactLookupLiteralWriteWhenDirty = (
    request: OperationLookupDirtyWriteRequest,
    formulaChangedCount: number,
    caches: ExactLookupImpactCaches,
  ): number => {
    const nextFormulaChangedCount = markAffectedExactLookupDependents(request, formulaChangedCount, caches)
    args.noteExactLookupLiteralWrite(request)
    return nextFormulaChangedCount
  }

  const noteSortedLookupLiteralWriteWhenDirty = (request: OperationLookupDirtyWriteRequest, formulaChangedCount: number): number => {
    const nextFormulaChangedCount = markAffectedApproximateLookupDependents(request, formulaChangedCount)
    args.noteSortedLookupLiteralWrite(request)
    return nextFormulaChangedCount
  }

  return {
    markAffectedExactLookupDependents,
    markAffectedApproximateLookupDependents,
    noteExactLookupLiteralWriteWhenDirty,
    noteSortedLookupLiteralWriteWhenDirty,
  }
}
