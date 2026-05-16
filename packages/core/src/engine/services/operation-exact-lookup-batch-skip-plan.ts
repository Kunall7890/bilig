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
  | {
      readonly kind: 'multi-numeric-dependents'
      readonly dependents: readonly ExactLookupBatchNumericDependent[]
    }
  | { readonly kind: 'fallback' }

export interface ExactLookupBatchSkipPlanInput {
  readonly col: number
  readonly formulas: {
    readonly get: (cellIndex: number) => { readonly directLookup: RuntimeDirectLookupDescriptor | undefined } | undefined
  }
  readonly getEntityDependents: (entityId: number) => ArrayLike<number>
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly readNumericValue: (cellIndex: number) => number | undefined
  readonly sheetId: number
}

interface ExactLookupBatchNumericDependent {
  readonly operandNumeric: number
  readonly rowEnd: number
  readonly rowStart: number
}

const prepareNumericDependent = (
  input: ExactLookupBatchSkipPlanInput,
  formulaCellIndex: number,
): ExactLookupBatchNumericDependent | ExactLookupBatchSkipPlan => {
  const directLookup = input.formulas.get(formulaCellIndex)?.directLookup
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
  return {
    operandNumeric,
    rowEnd,
    rowStart,
  }
}

export const prepareExactLookupBatchSkipPlan = (input: ExactLookupBatchSkipPlanInput): ExactLookupBatchSkipPlan => {
  const lookupEntity = makeExactLookupColumnEntity(input.sheetId, input.col)
  const singleDependent = input.getSingleEntityDependent(lookupEntity)
  if (singleDependent === -1) {
    return { kind: 'all' }
  }
  if (singleDependent >= 0) {
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
    if (directLookup.kind !== 'exact-uniform-numeric') {
      return {
        kind: 'single-numeric-dependent',
        operandNumeric,
        rowEnd,
        rowStart,
      }
    }
    return {
      kind: 'single-uniform-numeric-dependent',
      operandNumeric,
      rowEnd,
      rowStart,
      start: directLookup.start,
      step: directLookup.step,
    }
  }

  const dependents = input.getEntityDependents(lookupEntity)
  if (dependents.length === 0) {
    return { kind: 'all' }
  }
  const prepared: ExactLookupBatchNumericDependent[] = []
  for (let index = 0; index < dependents.length; index += 1) {
    const dependent = prepareNumericDependent(input, dependents[index]!)
    if ('kind' in dependent) {
      return dependent
    }
    prepared.push(dependent)
  }
  return {
    kind: 'multi-numeric-dependents',
    dependents: prepared,
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
    case 'multi-numeric-dependents':
      for (let index = 0; index < plan.dependents.length; index += 1) {
        const dependent = plan.dependents[index]!
        if (
          row >= dependent.rowStart &&
          row <= dependent.rowEnd &&
          (sameExactNumericValue(oldNumeric, dependent.operandNumeric) || sameExactNumericValue(newNumeric, dependent.operandNumeric))
        ) {
          return false
        }
      }
      return true
    case 'fallback':
      return false
  }
}
