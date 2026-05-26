import { toWorkbookRefData, type WorkbookRef } from './find.js'
import { normalizeWorkbookActionInput } from './input.js'
import type { WorkbookActionPlan } from './model.js'
import type { WorkbookRunAdapter, WorkbookRunOptions } from './run.js'
import type {
  WorkbookCheckResult,
  WorkbookCommandResolvedRefs,
  WorkbookResolvedRefValue,
  WorkbookRunApplyCommandReceipt,
  WorkbookRunApplySummary,
  WorkbookRunError,
} from './result.js'

export type NormalizedWorkbookRunOptions = {
  readonly requireApplyProof: boolean
  readonly requirePlanId: boolean
  readonly requireResolvedRefs: boolean
  readonly requireChecks: boolean
  readonly requireCheckProof: boolean
  readonly requireRevision: boolean
  readonly requireNoUnverified: boolean
  readonly expectedBaseRevision?: number
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

const RUN_OPTION_KEYS: ReadonlySet<string> = new Set([
  'strict',
  'requireApplyProof',
  'requirePlanId',
  'requireResolvedRefs',
  'requireChecks',
  'requireCheckProof',
  'requireRevision',
  'requireNoUnverified',
  'expectedBaseRevision',
])

function keyName(key: string | symbol): string {
  return typeof key === 'symbol' ? key.toString() : key
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
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
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

function normalizeOptionalRevisionOption(source: object): number | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(source, 'expectedBaseRevision')
  if (descriptor === undefined) {
    return undefined
  }
  if (!('value' in descriptor)) {
    throw new WorkbookRunOptionsError('options.expectedBaseRevision', 'Workbook run option expectedBaseRevision must be a data property')
  }
  if (typeof descriptor.value !== 'number' || !Number.isSafeInteger(descriptor.value) || descriptor.value < 0) {
    throw new WorkbookRunOptionsError(
      'options.expectedBaseRevision',
      'Workbook run option expectedBaseRevision must be a non-negative safe integer',
    )
  }
  return descriptor.value
}

function assertKnownRunOptions(source: object): void {
  for (const key of Reflect.ownKeys(source)) {
    const name = keyName(key)
    if (typeof key === 'symbol' || !RUN_OPTION_KEYS.has(key)) {
      throw new WorkbookRunOptionsError(`options.${name}`, `Workbook run option ${name} is unknown`)
    }
  }
}

export function normalizeRunOptions(options: unknown): WorkbookRunOptionsNormalization {
  if (options === undefined) {
    return {
      status: 'valid',
      options: {
        requireApplyProof: false,
        requirePlanId: false,
        requireResolvedRefs: false,
        requireChecks: false,
        requireCheckProof: false,
        requireRevision: false,
        requireNoUnverified: false,
      },
    }
  }
  if (!isRecord(options)) {
    return {
      status: 'invalid',
      error: runErrorAt('invalid_run_options', 'Workbook run options must be a plain object', 'options'),
    }
  }
  try {
    assertKnownRunOptions(options)
    const strict = normalizeRunOption(options, 'strict')
    const requireApplyProof = normalizeRunOption(options, 'requireApplyProof')
    const requirePlanId = normalizeRunOption(options, 'requirePlanId')
    const requireResolvedRefs = normalizeRunOption(options, 'requireResolvedRefs')
    const requireChecks = normalizeRunOption(options, 'requireChecks')
    const requireCheckProof = normalizeRunOption(options, 'requireCheckProof')
    const requireRevision = normalizeRunOption(options, 'requireRevision')
    const requireNoUnverified = normalizeRunOption(options, 'requireNoUnverified')
    const expectedBaseRevision = normalizeOptionalRevisionOption(options)
    return {
      status: 'valid',
      options: {
        requireApplyProof: strict || requireApplyProof,
        requirePlanId: strict || requirePlanId,
        requireResolvedRefs: strict || requireResolvedRefs,
        requireChecks: strict || requireChecks,
        requireCheckProof: strict || requireCheckProof,
        requireRevision: strict || requireRevision,
        requireNoUnverified: strict || requireNoUnverified,
        ...(expectedBaseRevision !== undefined ? { expectedBaseRevision } : {}),
      },
    }
  } catch (error) {
    return {
      status: 'invalid',
      error: runErrorAt('invalid_run_options', errorMessage(error), runOptionsErrorPath(error), 'invalid_run_options'),
    }
  }
}

function isRangeResolvedRefValue(value: WorkbookResolvedRefValue | undefined): boolean {
  if (value === undefined) {
    return false
  }
  const refs = Array.isArray(value) ? value : [value]
  return refs.every((ref) => ref.kind === 'range')
}

function singleResolvedRefValueMatches(value: WorkbookResolvedRefValue | undefined, expected: unknown): boolean {
  if (value === undefined || Array.isArray(value)) {
    return false
  }
  return JSON.stringify(normalizeWorkbookActionInput(value)) === JSON.stringify(normalizeWorkbookActionInput(expected))
}

function resolvedInputsForCommand(command: WorkbookActionPlan['commands'][number]): readonly WorkbookRef[] {
  return command.kind === 'writeFormula' ? command.inputs : []
}

function commandNeedsResolvedRefs(command: WorkbookActionPlan['commands'][number]): boolean {
  return command.target !== undefined || resolvedInputsForCommand(command).length > 0
}

function resolvedRefsHaveExpectedShape(
  command: WorkbookActionPlan['commands'][number],
  resolvedRefs: WorkbookCommandResolvedRefs,
): boolean {
  if (command.target !== undefined) {
    if (!isRangeResolvedRefValue(resolvedRefs.target)) {
      return false
    }
    if (command.target.kind === 'range' && !singleResolvedRefValueMatches(resolvedRefs.target, toWorkbookRefData(command.target))) {
      return false
    }
  }

  const inputs = resolvedInputsForCommand(command)
  if (inputs.length === 0) {
    return true
  }
  if (resolvedRefs.inputs === undefined || resolvedRefs.inputs.length !== inputs.length) {
    return false
  }
  return resolvedRefs.inputs.every((resolvedInput, index) => {
    if (!isRangeResolvedRefValue(resolvedInput)) {
      return false
    }
    const plannedInput = inputs[index]
    return plannedInput?.kind === 'range' ? singleResolvedRefValueMatches(resolvedInput, toWorkbookRefData(plannedInput)) : true
  })
}

function resolvedRefProofError(plan: WorkbookActionPlan, receipt: WorkbookRunApplyCommandReceipt): WorkbookRunError | undefined {
  const command = plan.commands[receipt.commandIndex]
  if (command === undefined) {
    return undefined
  }
  const needsTarget = command.target !== undefined
  const needsInputs = resolvedInputsForCommand(command).length > 0
  if (!needsTarget && !needsInputs) {
    return undefined
  }
  if (receipt.resolvedRefs === undefined) {
    return runError('apply_not_verified', `Adapter did not return resolved ref proof for command ${String(receipt.commandIndex)}`)
  }
  if (!resolvedRefsHaveExpectedShape(command, receipt.resolvedRefs)) {
    return runError(
      'apply_not_verified',
      `Adapter resolved ref proof for command ${String(receipt.commandIndex)} must bind planned refs to concrete ranges`,
    )
  }
  return undefined
}

export function preApplyProofErrors(
  plan: WorkbookActionPlan,
  needsApply: boolean,
  options: NormalizedWorkbookRunOptions,
): readonly WorkbookRunError[] {
  if (needsApply && options.requireChecks && plan.checks.length === 0) {
    return [runError('check_not_verified', 'Strict workbook runs require at least one check before applying mutating plans')]
  }
  return []
}

export function applyProofErrors(
  plan: WorkbookActionPlan,
  apply: WorkbookRunApplySummary,
  options: NormalizedWorkbookRunOptions,
): readonly WorkbookRunError[] {
  if (apply.matched === false) {
    return [runError('apply_mismatch', 'Adapter applied ops do not match its preview ops')]
  }
  if ((options.requireApplyProof || options.requireNoUnverified) && apply.matched === null) {
    return [runError('apply_not_verified', 'Adapter did not return both previewOps and appliedOps')]
  }
  if ((options.requireApplyProof || options.requireNoUnverified) && plan.commands.length > 0 && apply.commandReceipts === undefined) {
    return [runError('apply_not_verified', 'Adapter did not bind planned commands to materialized ops')]
  }
  if (options.requireResolvedRefs && apply.commandReceipts === undefined && plan.commands.some(commandNeedsResolvedRefs)) {
    return [runError('apply_not_verified', 'Adapter did not bind planned commands to resolved ref proof')]
  }
  if (options.requireApplyProof && apply.commandReceipts !== undefined) {
    const unprovedReceipt = apply.commandReceipts.find((receipt) => receipt.appliedOps.length === 0 && receipt.noop === undefined)
    if (unprovedReceipt !== undefined) {
      return [
        runError('apply_not_verified', `Adapter did not bind command ${String(unprovedReceipt.commandIndex)} to concrete applied ops`),
      ]
    }
  }
  if (options.requireResolvedRefs && apply.commandReceipts !== undefined) {
    const resolvedRefsError = apply.commandReceipts
      .map((receipt) => resolvedRefProofError(plan, receipt))
      .find((error) => error !== undefined)
    if (resolvedRefsError !== undefined) {
      return [resolvedRefsError]
    }
  }
  if (options.requirePlanId && apply.planId === undefined) {
    return [runError('plan_not_verified', 'Adapter did not bind apply proof to a plan id')]
  }
  if (options.expectedBaseRevision !== undefined && apply.baseRevision === undefined) {
    return [
      runError('plan_not_verified', `Adapter did not bind apply proof to expected base revision ${String(options.expectedBaseRevision)}`),
    ]
  }
  if (options.expectedBaseRevision !== undefined && apply.baseRevision !== options.expectedBaseRevision) {
    return [
      runError(
        'plan_not_verified',
        `Adapter apply proof base revision ${String(apply.baseRevision)} did not match expected base revision ${String(options.expectedBaseRevision)}`,
      ),
    ]
  }
  if (options.requireRevision && (apply.baseRevision === undefined || apply.revision === undefined)) {
    return [runError('plan_not_verified', 'Adapter did not bind apply proof to workbook revisions')]
  }
  return []
}

export function checkProofErrors(
  checks: readonly WorkbookCheckResult[],
  options: NormalizedWorkbookRunOptions,
): readonly WorkbookRunError[] {
  if (!options.requireCheckProof) {
    return []
  }
  const checkWithoutProof = checks.find((check) => check.status === 'passed' && check.proof === undefined)
  if (checkWithoutProof === undefined) {
    return []
  }
  return [
    runError(
      'check_not_verified',
      `${checkWithoutProof.target?.label ?? checkWithoutProof.kind} passed check ${checkWithoutProof.kind} without proof`,
    ),
  ]
}
