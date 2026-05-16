import { makeExactLookupColumnEntity } from '../../entity-ids.js'
import type { RuntimeDirectLookupDescriptor } from '../runtime-state.js'
import { directLookupRowBounds, sameExactNumericValue } from './direct-lookup-helpers.js'

export type ExactLookupBatchSkipPlan =
  | { readonly kind: 'all' }
  | {
      readonly kind: 'single-numeric-dependent'
      readonly operandNumeric: number
      readonly rowEnd: number
      readonly rowStart: number
    }
  | {
      readonly kind: 'single-uniform-numeric-dependent'
      readonly operandNumeric: number
      readonly rowEnd: number
      readonly rowStart: number
      readonly start: number
      readonly step: number
    }
  | { readonly kind: 'fallback' }

export interface ExactLookupBatchSkipPlanInput {
  readonly col: number
  readonly formulas: {
    readonly get: (cellIndex: number) => { readonly directLookup: RuntimeDirectLookupDescriptor | undefined } | undefined
  }
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly readNumericValue: (cellIndex: number) => number | undefined
  readonly sheetId: number
}

export const prepareExactLookupBatchSkipPlan = (input: ExactLookupBatchSkipPlanInput): ExactLookupBatchSkipPlan => {
  const singleDependent = input.getSingleEntityDependent(makeExactLookupColumnEntity(input.sheetId, input.col))
  if (singleDependent === -1) {
    return { kind: 'all' }
  }
  if (singleDependent < 0) {
    return { kind: 'fallback' }
  }
  const directLookup = input.formulas.get(singleDependent)?.directLookup
  if (directLookup?.kind !== 'exact' && directLookup?.kind !== 'exact-uniform-numeric') {
    return { kind: 'fallback' }
  }
  if (directLookup.kind === 'exact-uniform-numeric' && directLookup.tailPatch !== undefined) {
    return { kind: 'fallback' }
  }
  const operandNumeric = input.readNumericValue(directLookup.operandCellIndex)
  if (operandNumeric === undefined) {
    return { kind: 'fallback' }
  }
  const { rowStart, rowEnd } = directLookupRowBounds(directLookup)
  if (directLookup.kind === 'exact-uniform-numeric') {
    return {
      kind: 'single-uniform-numeric-dependent',
      operandNumeric,
      rowEnd,
      rowStart,
      start: directLookup.start,
      step: directLookup.step,
    }
  }
  return {
    kind: 'single-numeric-dependent',
    operandNumeric,
    rowEnd,
    rowStart,
  }
}

export const exactLookupBatchOldNumeric = (plan: ExactLookupBatchSkipPlan, row: number): number | undefined => {
  if (plan.kind !== 'single-uniform-numeric-dependent' || row < plan.rowStart || row > plan.rowEnd) {
    return undefined
  }
  return plan.start + plan.step * (row - plan.rowStart)
}

export const exactLookupBatchWriteHandled = (
  plan: ExactLookupBatchSkipPlan,
  row: number,
  oldNumeric: number,
  newNumeric: number,
): boolean => {
  switch (plan.kind) {
    case 'all':
      return true
    case 'single-numeric-dependent':
    case 'single-uniform-numeric-dependent':
      return (
        row < plan.rowStart ||
        row > plan.rowEnd ||
        (!sameExactNumericValue(oldNumeric, plan.operandNumeric) && !sameExactNumericValue(newNumeric, plan.operandNumeric))
      )
    case 'fallback':
      return false
  }
}
