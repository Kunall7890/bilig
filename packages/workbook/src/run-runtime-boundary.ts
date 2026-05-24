import type { WorkbookActionPlan } from './model.js'
import type { WorkbookRunAdapter, WorkbookRunOptions } from './run.js'
import type { WorkbookRunApplySummary, WorkbookRunError } from './result.js'

export type NormalizedWorkbookRunOptions = {
  readonly requireApplyProof: boolean
  readonly requirePlanId: boolean
  readonly requireResolvedRefs: boolean
}

export type WorkbookRunOptionsNormalization =
  | {
      readonly status: 'valid'
      readonly options: NormalizedWorkbookRunOptions
    }
  | {
      readonly status: 'invalid'
      readonly error: WorkbookRunError
    }

class WorkbookRunOptionsError extends Error {
  readonly path: string

  constructor(path: string, message: string) {
    super(message)
    this.name = 'WorkbookRunOptionsError'
    this.path = path
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function runOptionsErrorPath(error: unknown): string {
  return error instanceof WorkbookRunOptionsError ? error.path : 'options'
}

function runError(code: WorkbookRunError['code'], message: string): WorkbookRunError {
  return {
    code,
    message,
  }
}

function runErrorAt(code: WorkbookRunError['code'], message: string, path: string, issueCode = code): WorkbookRunError {
  return {
    code,
    message,
    path,
    issueCode,
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function ownDataDescriptor(value: object, key: string): PropertyDescriptor | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && 'value' in descriptor ? descriptor : undefined
}

function ownAdapterMethod<Refs>(adapter: WorkbookRunAdapter<Refs>, method: keyof WorkbookRunAdapter): unknown {
  if (!isRecord(adapter)) {
    return undefined
  }
  const descriptor = ownDataDescriptor(adapter, method)
  return descriptor?.value
}

function isApplyMethod<Refs>(value: unknown): value is NonNullable<WorkbookRunAdapter<Refs>['apply']> {
  return typeof value === 'function'
}

function isReadMethod<Refs>(value: unknown): value is NonNullable<WorkbookRunAdapter<Refs>['read']> {
  return typeof value === 'function'
}

function isVerifyChecksMethod<Refs>(value: unknown): value is NonNullable<WorkbookRunAdapter<Refs>['verifyChecks']> {
  return typeof value === 'function'
}

export function adapterApplyMethod<Refs>(adapter: WorkbookRunAdapter<Refs>): NonNullable<WorkbookRunAdapter<Refs>['apply']> | undefined {
  const value = ownAdapterMethod(adapter, 'apply')
  return isApplyMethod<Refs>(value) ? value : undefined
}

export function adapterReadMethod<Refs>(adapter: WorkbookRunAdapter<Refs>): NonNullable<WorkbookRunAdapter<Refs>['read']> | undefined {
  const value = ownAdapterMethod(adapter, 'read')
  return isReadMethod<Refs>(value) ? value : undefined
}

export function adapterVerifyChecksMethod<Refs>(
  adapter: WorkbookRunAdapter<Refs>,
): NonNullable<WorkbookRunAdapter<Refs>['verifyChecks']> | undefined {
  const value = ownAdapterMethod(adapter, 'verifyChecks')
  return isVerifyChecksMethod<Refs>(value) ? value : undefined
}

function normalizeRunOption(source: object, key: keyof WorkbookRunOptions): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(source, key)
  const path = `options.${key}`
  if (descriptor === undefined) {
    return false
  }
  if (!('value' in descriptor)) {
    throw new WorkbookRunOptionsError(path, `Workbook run option ${key} must be a data property`)
  }
  if (typeof descriptor.value !== 'boolean') {
    throw new WorkbookRunOptionsError(path, `Workbook run option ${key} must be a boolean`)
  }
  return descriptor.value
}

export function normalizeRunOptions(options: unknown): WorkbookRunOptionsNormalization {
  if (options === undefined) {
    return {
      status: 'valid',
      options: {
        requireApplyProof: false,
        requirePlanId: false,
        requireResolvedRefs: false,
      },
    }
  }
  if (!isRecord(options) || Array.isArray(options)) {
    return {
      status: 'invalid',
      error: runErrorAt('invalid_run_options', 'Workbook run options must be an object', 'options'),
    }
  }
  try {
    const strict = normalizeRunOption(options, 'strict')
    const requireApplyProof = normalizeRunOption(options, 'requireApplyProof')
    const requirePlanId = normalizeRunOption(options, 'requirePlanId')
    return {
      status: 'valid',
      options: {
        requireApplyProof: strict || requireApplyProof,
        requirePlanId: strict || requirePlanId,
        requireResolvedRefs: strict,
      },
    }
  } catch (error) {
    return {
      status: 'invalid',
      error: runErrorAt('invalid_run_options', errorMessage(error), runOptionsErrorPath(error), 'invalid_run_options'),
    }
  }
}

function resolvedRefProofNeedsData(command: WorkbookActionPlan['commands'][number]): boolean {
  return command.target !== undefined || (command.kind === 'writeFormula' && command.inputs.length > 0)
}

function hasResolvedRefProofData(value: unknown): boolean {
  return typeof value === 'object' && value !== null && !Array.isArray(value) && Object.keys(value).length > 0
}

export function applyProofErrors(
  plan: WorkbookActionPlan,
  apply: WorkbookRunApplySummary,
  options: NormalizedWorkbookRunOptions,
): readonly WorkbookRunError[] {
  if (apply.matched === false) {
    return [runError('apply_mismatch', 'Adapter applied ops do not match its preview ops')]
  }
  if (options.requireApplyProof && apply.matched === null) {
    return [runError('apply_not_verified', 'Adapter did not return both previewOps and appliedOps')]
  }
  if (options.requireApplyProof && plan.commands.length > 0 && apply.commandReceipts === undefined) {
    return [runError('apply_not_verified', 'Adapter did not bind planned commands to materialized ops')]
  }
  if (options.requireApplyProof && apply.commandReceipts !== undefined) {
    const unprovedReceipt = apply.commandReceipts.find((receipt) => receipt.appliedOps.length === 0)
    if (unprovedReceipt !== undefined) {
      return [
        runError('apply_not_verified', `Adapter did not bind command ${String(unprovedReceipt.commandIndex)} to concrete applied ops`),
      ]
    }
  }
  if (options.requireResolvedRefs && apply.commandReceipts !== undefined) {
    const missingResolvedRefs = apply.commandReceipts.find((receipt) => {
      const command = plan.commands[receipt.commandIndex]
      return command !== undefined && resolvedRefProofNeedsData(command) && !hasResolvedRefProofData(receipt.resolvedRefs)
    })
    if (missingResolvedRefs !== undefined) {
      return [
        runError('apply_not_verified', `Adapter did not return resolved ref proof for command ${String(missingResolvedRefs.commandIndex)}`),
      ]
    }
  }
  if (options.requirePlanId && apply.planId === undefined) {
    return [runError('plan_not_verified', 'Adapter did not bind apply proof to a plan id')]
  }
  return []
}
