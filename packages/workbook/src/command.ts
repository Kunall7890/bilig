import { describePlan, type WorkbookActionPlanDescription } from './describe.js'
import { getOwnActionInput, hasOwnActionInput, type WorkbookActionInput } from './input.js'
import {
  planWorkbookAction,
  type WorkbookActionMap,
  type WorkbookActionPlan,
  type WorkbookActionPlanResult,
  type WorkbookModel,
} from './model.js'
import { describeRuntimeRequirements, type WorkbookRuntimeRequirements } from './requirements.js'
import { verifyPlan, type WorkbookPlanIssue, type WorkbookPlanVerification } from './verify.js'

export type WorkbookRevision = string | number

export interface WorkbookCommandBundleOptions {
  readonly idempotencyKey?: string
  readonly baseRevision?: WorkbookRevision
}

export interface WorkbookCommandBundle<Refs = unknown> {
  readonly schemaVersion: 1
  readonly commandId: string
  readonly idempotencyKey?: string
  readonly baseRevision?: WorkbookRevision
  readonly modelName: string
  readonly actionName: string
  readonly input?: WorkbookActionInput
  readonly plan: WorkbookActionPlan<Refs>
  readonly verification: WorkbookPlanVerification
  readonly requirements: WorkbookRuntimeRequirements
}

export interface WorkbookCommandBundleDescription {
  readonly schemaVersion: 1
  readonly commandId: string
  readonly idempotencyKey?: string
  readonly baseRevision?: WorkbookRevision
  readonly modelName: string
  readonly actionName: string
  readonly input?: WorkbookActionInput
  readonly plan: WorkbookActionPlanDescription
  readonly verification: WorkbookPlanVerification
  readonly requirements: WorkbookRuntimeRequirements
}

export type WorkbookCommandBundleResult<Refs = unknown> =
  | {
      readonly status: 'planned'
      readonly command: WorkbookCommandBundle<Refs>
    }
  | Extract<WorkbookActionPlanResult<Refs>, { readonly status: 'failed' }>

export type WorkbookCommandBundleIssueCode =
  | 'invalid_schema_version'
  | 'invalid_command_id'
  | 'invalid_idempotency_key'
  | 'invalid_base_revision'
  | 'model_name_mismatch'
  | 'action_name_mismatch'
  | 'input_mismatch'
  | 'requirements_mismatch'
  | 'verification_mismatch'
  | 'plan_invalid'

export const workbookCommandBundleIssueCodes = Object.freeze([
  'invalid_schema_version',
  'invalid_command_id',
  'invalid_idempotency_key',
  'invalid_base_revision',
  'model_name_mismatch',
  'action_name_mismatch',
  'input_mismatch',
  'requirements_mismatch',
  'verification_mismatch',
  'plan_invalid',
] satisfies readonly WorkbookCommandBundleIssueCode[])

export interface WorkbookCommandBundleIssue {
  readonly code: WorkbookCommandBundleIssueCode
  readonly message: string
  readonly path: string
  readonly planIssue?: WorkbookPlanIssue
}

export interface WorkbookCommandBundleVerification {
  readonly status: 'valid' | 'invalid'
  readonly commandId: string
  readonly modelName: string
  readonly actionName: string
  readonly issues: readonly WorkbookCommandBundleIssue[]
}

export function isWorkbookCommandBundleIssueCode(value: unknown): value is WorkbookCommandBundleIssueCode {
  return typeof value === 'string' && workbookCommandBundleIssueCodes.some((code) => code === value)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function normalizeOptionalText(value: string | undefined, label: string): string | undefined {
  if (value === undefined) {
    return undefined
  }
  const text = value.trim()
  if (text === '') {
    throw new Error(`${label} cannot be empty`)
  }
  return text
}

function isWorkbookRevision(value: unknown): value is WorkbookRevision {
  if (typeof value === 'string') {
    return value.trim() !== ''
  }
  return typeof value === 'number' && Number.isFinite(value)
}

function normalizeOptionalRevision(value: WorkbookRevision | undefined): WorkbookRevision | undefined {
  if (value === undefined) {
    return undefined
  }
  if (!isWorkbookRevision(value)) {
    throw new Error('Workbook command baseRevision must be a non-empty string or finite number')
  }
  return typeof value === 'string' ? value.trim() : value
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalValue)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entry]) => entry !== undefined)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalValue(entry)]),
    )
  }
  return value
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

function hashText(value: string): string {
  let hash = 0xcbf29ce484222325n
  const prime = 0x100000001b3n
  for (let index = 0; index < value.length; index += 1) {
    hash ^= BigInt(value.charCodeAt(index))
    hash = BigInt.asUintN(64, hash * prime)
  }
  return hash.toString(16).padStart(16, '0')
}

function planInputProperty(plan: WorkbookActionPlan): { readonly input: WorkbookActionInput } | {} {
  return hasOwnActionInput(plan) ? { input: getOwnActionInput(plan) } : {}
}

function bundleIdentityPayload<Refs>(
  plan: WorkbookActionPlan<Refs>,
  options: {
    readonly baseRevision: WorkbookRevision | undefined
    readonly idempotencyKey: string | undefined
  },
): object {
  const verification = verifyPlan(plan)
  return {
    schemaVersion: 1,
    modelName: plan.modelName,
    actionName: plan.actionName,
    ...planInputProperty(plan),
    ...(options.baseRevision !== undefined ? { baseRevision: options.baseRevision } : {}),
    ...(options.idempotencyKey !== undefined ? { idempotencyKey: options.idempotencyKey } : {}),
    plan: describePlan(plan),
    verification,
    requirements: describeRuntimeRequirements(plan),
  }
}

function commandIdFor<Refs>(
  plan: WorkbookActionPlan<Refs>,
  options: {
    readonly baseRevision: WorkbookRevision | undefined
    readonly idempotencyKey: string | undefined
  },
): string {
  return `cmd_${hashText(canonicalJson(bundleIdentityPayload(plan, options)))}`
}

function deepFreeze<T>(value: T, seen: WeakSet<object> = new WeakSet()): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Reflect.ownKeys(value).forEach((key) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    if (descriptor !== undefined && 'value' in descriptor) {
      deepFreeze(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}

export function buildWorkbookCommandBundle<Refs>(
  plan: WorkbookActionPlan<Refs>,
  options: WorkbookCommandBundleOptions = {},
): WorkbookCommandBundle<Refs> {
  const baseRevision = normalizeOptionalRevision(options.baseRevision)
  const idempotencyKey = normalizeOptionalText(options.idempotencyKey, 'Workbook command idempotencyKey')
  const verification = verifyPlan(plan)
  const requirements = describeRuntimeRequirements(plan)
  return deepFreeze({
    schemaVersion: 1,
    commandId: commandIdFor(plan, { baseRevision, idempotencyKey }),
    ...(idempotencyKey !== undefined ? { idempotencyKey } : {}),
    ...(baseRevision !== undefined ? { baseRevision } : {}),
    modelName: plan.modelName,
    actionName: plan.actionName,
    ...planInputProperty(plan),
    plan,
    verification,
    requirements,
  })
}

export function planWorkbookCommand<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
  actionName: string,
  input?: WorkbookActionInput,
  options: WorkbookCommandBundleOptions = {},
): WorkbookCommandBundleResult<Refs> {
  const result = planWorkbookAction(model, actionName, input)
  if (result.status === 'failed') {
    return result
  }
  return {
    status: 'planned',
    command: buildWorkbookCommandBundle(result.plan, options),
  }
}

function issue(input: {
  readonly code: WorkbookCommandBundleIssueCode
  readonly message: string
  readonly path: string
  readonly planIssue?: WorkbookPlanIssue
}): WorkbookCommandBundleIssue {
  return {
    code: input.code,
    message: input.message,
    path: input.path,
    ...(input.planIssue !== undefined ? { planIssue: input.planIssue } : {}),
  }
}

function bundleInput(bundle: WorkbookCommandBundle): WorkbookActionInput | undefined {
  return Object.prototype.hasOwnProperty.call(bundle, 'input') ? bundle.input : undefined
}

export function verifyWorkbookCommandBundle<Refs>(bundle: WorkbookCommandBundle<Refs>): WorkbookCommandBundleVerification {
  const issues: WorkbookCommandBundleIssue[] = []
  const modelName = isRecord(bundle) && typeof bundle.modelName === 'string' ? bundle.modelName : ''
  const actionName = isRecord(bundle) && typeof bundle.actionName === 'string' ? bundle.actionName : ''
  const commandId = isRecord(bundle) && typeof bundle.commandId === 'string' ? bundle.commandId : ''

  if (!isRecord(bundle)) {
    return {
      status: 'invalid',
      commandId,
      modelName,
      actionName,
      issues: [issue({ code: 'invalid_schema_version', path: 'command', message: 'Workbook command bundle must be an object' })],
    }
  }

  if (bundle.schemaVersion !== 1) {
    issues.push(
      issue({
        code: 'invalid_schema_version',
        path: 'schemaVersion',
        message: 'Workbook command bundle schemaVersion must be 1',
      }),
    )
  }

  if (typeof bundle.commandId !== 'string' || bundle.commandId.trim() === '') {
    issues.push(
      issue({
        code: 'invalid_command_id',
        path: 'commandId',
        message: 'Workbook command bundle commandId must be a non-empty string',
      }),
    )
  }

  if (bundle.idempotencyKey !== undefined && (typeof bundle.idempotencyKey !== 'string' || bundle.idempotencyKey.trim() === '')) {
    issues.push(
      issue({
        code: 'invalid_idempotency_key',
        path: 'idempotencyKey',
        message: 'Workbook command bundle idempotencyKey must be a non-empty string when provided',
      }),
    )
  }

  if (bundle.baseRevision !== undefined && !isWorkbookRevision(bundle.baseRevision)) {
    issues.push(
      issue({
        code: 'invalid_base_revision',
        path: 'baseRevision',
        message: 'Workbook command bundle baseRevision must be a non-empty string or finite number when provided',
      }),
    )
  }

  if (bundle.modelName !== bundle.plan.modelName) {
    issues.push(
      issue({
        code: 'model_name_mismatch',
        path: 'modelName',
        message: `Workbook command bundle modelName ${String(bundle.modelName)} does not match plan ${bundle.plan.modelName}`,
      }),
    )
  }

  if (bundle.actionName !== bundle.plan.actionName) {
    issues.push(
      issue({
        code: 'action_name_mismatch',
        path: 'actionName',
        message: `Workbook command bundle actionName ${String(bundle.actionName)} does not match plan ${bundle.plan.actionName}`,
      }),
    )
  }

  const expectedInput = hasOwnActionInput(bundle.plan) ? getOwnActionInput(bundle.plan) : undefined
  if (canonicalJson(bundleInput(bundle)) !== canonicalJson(expectedInput)) {
    issues.push(
      issue({
        code: 'input_mismatch',
        path: 'input',
        message: 'Workbook command bundle input does not match plan input',
      }),
    )
  }

  const expectedRequirements = describeRuntimeRequirements(bundle.plan)
  if (canonicalJson(bundle.requirements) !== canonicalJson(expectedRequirements)) {
    issues.push(
      issue({
        code: 'requirements_mismatch',
        path: 'requirements',
        message: 'Workbook command bundle requirements do not match the plan',
      }),
    )
  }

  const expectedVerification = verifyPlan(bundle.plan)
  if (canonicalJson(bundle.verification) !== canonicalJson(expectedVerification)) {
    issues.push(
      issue({
        code: 'verification_mismatch',
        path: 'verification',
        message: 'Workbook command bundle verification does not match the plan',
      }),
    )
  }

  expectedVerification.issues.forEach((planIssue) => {
    issues.push(
      issue({
        code: 'plan_invalid',
        path: `verification.${planIssue.path}`,
        message: planIssue.message,
        planIssue,
      }),
    )
  })

  if (isWorkbookRevision(bundle.baseRevision) || bundle.baseRevision === undefined) {
    const expectedCommandId = commandIdFor(bundle.plan, {
      baseRevision: bundle.baseRevision,
      idempotencyKey:
        typeof bundle.idempotencyKey === 'string' && bundle.idempotencyKey.trim() !== '' ? bundle.idempotencyKey.trim() : undefined,
    })
    if (bundle.commandId !== expectedCommandId) {
      issues.push(
        issue({
          code: 'invalid_command_id',
          path: 'commandId',
          message: `Workbook command bundle commandId ${String(bundle.commandId)} does not match ${expectedCommandId}`,
        }),
      )
    }
  }

  return {
    status: issues.length === 0 ? 'valid' : 'invalid',
    commandId,
    modelName,
    actionName,
    issues,
  }
}

export function describeCommandBundle<Refs>(bundle: WorkbookCommandBundle<Refs>): WorkbookCommandBundleDescription {
  return {
    schemaVersion: 1,
    commandId: bundle.commandId,
    ...(bundle.idempotencyKey !== undefined ? { idempotencyKey: bundle.idempotencyKey } : {}),
    ...(bundle.baseRevision !== undefined ? { baseRevision: bundle.baseRevision } : {}),
    modelName: bundle.modelName,
    actionName: bundle.actionName,
    ...planInputProperty(bundle.plan),
    plan: describePlan(bundle.plan),
    verification: bundle.verification,
    requirements: bundle.requirements,
  }
}
