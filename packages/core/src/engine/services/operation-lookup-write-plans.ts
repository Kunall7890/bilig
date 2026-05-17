import {
  applyLookupNumericColumnWriteTailPatch,
  type LookupNumericColumnWritePlan,
  type OperationLookupPlanner,
} from './operation-lookup-planner.js'

type LookupPlanner = Pick<OperationLookupPlanner, 'planExactLookupNumericColumnWrite' | 'planApproximateLookupNumericColumnWrite'>

export interface OperationLookupNumericWritePlanRequest {
  readonly isRestore: boolean
  readonly hasAggregateDependents: boolean
  readonly hasExactLookupDependents: boolean
  readonly hasSortedLookupDependents: boolean
  readonly sheetId: number
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly oldExactLookupNumber: number | undefined
  readonly newExactLookupNumber: number | undefined
  readonly oldApproximateLookupNumber: number | undefined
  readonly newApproximateLookupNumber: number | undefined
  readonly planner: LookupPlanner
}

export interface OperationLookupNumericWritePlans {
  readonly exact: LookupNumericColumnWritePlan | undefined
  readonly sorted: LookupNumericColumnWritePlan | undefined
  readonly exactHandled: boolean
  readonly sortedHandled: boolean
  readonly needsLookupValueRead: boolean
  readonly handledLookupDependents: boolean
}

export function planOperationLookupNumericWrites(request: OperationLookupNumericWritePlanRequest): OperationLookupNumericWritePlans {
  const exact =
    !request.isRestore &&
    request.hasExactLookupDependents &&
    !request.hasAggregateDependents &&
    request.oldExactLookupNumber !== undefined &&
    request.newExactLookupNumber !== undefined
      ? request.planner.planExactLookupNumericColumnWrite(
          request.sheetId,
          request.col,
          request.row,
          request.oldExactLookupNumber,
          request.newExactLookupNumber,
        )
      : undefined
  const sorted =
    !request.isRestore &&
    request.hasSortedLookupDependents &&
    !request.hasAggregateDependents &&
    request.oldApproximateLookupNumber !== undefined &&
    request.newApproximateLookupNumber !== undefined
      ? request.planner.planApproximateLookupNumericColumnWrite(
          request.sheetId,
          request.sheetName,
          request.col,
          request.row,
          request.oldApproximateLookupNumber,
          request.newApproximateLookupNumber,
        )
      : undefined
  const exactHandled = exact?.handled === true
  const sortedHandled = sorted?.handled === true
  return {
    exact,
    sorted,
    exactHandled,
    sortedHandled,
    needsLookupValueRead:
      request.hasAggregateDependents ||
      (request.hasExactLookupDependents && !exactHandled) ||
      (request.hasSortedLookupDependents && !sortedHandled),
    handledLookupDependents: (request.hasExactLookupDependents && exactHandled) || (request.hasSortedLookupDependents && sortedHandled),
  }
}

export function applyOperationLookupNumericWriteTailPatches(
  plans: Pick<OperationLookupNumericWritePlans, 'exact' | 'sorted'>,
  request: {
    readonly row: number
    readonly oldExactLookupNumber: number | undefined
    readonly newExactLookupNumber: number | undefined
    readonly oldApproximateLookupNumber: number | undefined
    readonly newApproximateLookupNumber: number | undefined
    readonly columnVersionAfterWrite: number
  },
): { readonly exact: boolean; readonly sorted: boolean } {
  let exact = false
  let sorted = false
  if (request.oldExactLookupNumber !== undefined && request.newExactLookupNumber !== undefined) {
    exact = applyLookupNumericColumnWriteTailPatch(plans.exact, {
      row: request.row,
      oldNumeric: request.oldExactLookupNumber,
      newNumeric: request.newExactLookupNumber,
      columnVersion: request.columnVersionAfterWrite,
    })
  }
  if (request.oldApproximateLookupNumber !== undefined && request.newApproximateLookupNumber !== undefined) {
    sorted = applyLookupNumericColumnWriteTailPatch(plans.sorted, {
      row: request.row,
      oldNumeric: request.oldApproximateLookupNumber,
      newNumeric: request.newApproximateLookupNumber,
      columnVersion: request.columnVersionAfterWrite,
    })
  }
  return { exact, sorted }
}
