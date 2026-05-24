import { isLiteralInput } from '@bilig/protocol'
import {
  describePlan,
  type WorkbookActionCommandDescription,
  type WorkbookActionPlanDescription,
  type WorkbookChangeSummaryDescription,
  type WorkbookCheckExpectationDescription,
  type WorkbookCheckResultDescription,
  type WorkbookFormulaLabelDescription,
  type WorkbookRefDescription,
} from './describe.js'
import { hydrateWorkbookRef, isWorkbookRefData, type WorkbookRef } from './find.js'
import type { WorkbookFormulaLabel } from './formula.js'
import { isWorkbookOp } from './guards.js'
import { WorkbookActionInputError, isWorkbookActionInput, normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import type { WorkbookActionCommand, WorkbookActionPlan } from './model.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookChangeSummary, WorkbookCheckExpectation, WorkbookCheckResult, WorkbookCheckStatus } from './result.js'

export type WorkbookPlanData = WorkbookActionPlanDescription
export type WorkbookPlanId = string

export interface WorkbookPlanDataRefs {
  readonly refsUsed: readonly WorkbookRef[]
}

export type WorkbookExecutablePlan<Refs = unknown> = WorkbookActionPlan<Refs> | WorkbookPlanData
export type WorkbookPlanDataIssueCode = 'invalid_plan_data'

export interface WorkbookPlanDataIssue {
  readonly code: WorkbookPlanDataIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookPlanDataCheckResult =
  | {
      readonly status: 'valid'
      readonly plan: WorkbookPlanData
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookPlanDataIssue[]
    }

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    const descriptors = Object.getOwnPropertyDescriptors(value)
    return Array.from({ length: value.length }, (_entry, index) => {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined) {
        return undefined
      }
      if (!('value' in descriptor)) {
        throw new Error('Workbook plan data arrays must contain only data properties')
      }
      return canonicalValue(descriptor.value)
    })
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(Object.getOwnPropertyDescriptors(value))
        .filter(([, descriptor]) => descriptor.enumerable)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, descriptor]) => {
          if (!('value' in descriptor)) {
            throw new Error('Workbook plan data objects must contain only data properties')
          }
          return [key, canonicalValue(descriptor.value)]
        }),
    )
  }
  return value
}

function fnv1a64(input: string, seed: bigint): string {
  let hash = seed
  for (let index = 0; index < input.length; index += 1) {
    hash ^= BigInt(input.charCodeAt(index))
    hash = BigInt.asUintN(64, hash * 0x100000001b3n)
  }
  return hash.toString(16).padStart(16, '0')
}

function checkedPlanData(value: unknown): WorkbookPlanData {
  const check = checkPlanData(value)
  if (check.status === 'invalid') {
    const [firstIssue] = check.issues
    throw new Error(firstIssue === undefined ? 'Workbook plan data is invalid' : `Workbook plan data is invalid: ${firstIssue.message}`)
  }
  return check.plan
}

export function workbookPlanId<Refs>(plan: WorkbookExecutablePlan<Refs>): WorkbookPlanId {
  const data = isHydratedPlan<Refs>(plan) ? describePlan(plan) : checkedPlanData(plan)
  const json = canonicalJson(data)
  return `bilig-plan-v1:${fnv1a64(json, 0xcbf29ce484222325n)}${fnv1a64(json, 0x84222325cbf29cen)}`
}

function ownValue(value: object, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function hasOwnValue(value: object, key: string): boolean {
  return Object.getOwnPropertyDescriptor(value, key) !== undefined
}

function hasString(value: Record<string, unknown>, key: string): boolean {
  return typeof ownValue(value, key) === 'string'
}

function hasOptionalString(value: Record<string, unknown>, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor === undefined || typeof descriptor.value === 'string'
}

function arrayEvery<T>(value: unknown, predicate: (entry: unknown) => entry is T): value is readonly T[]
function arrayEvery(value: unknown, predicate: (entry: unknown) => boolean): boolean
function arrayEvery(value: unknown, predicate: (entry: unknown) => boolean): boolean {
  if (!Array.isArray(value)) {
    return false
  }

  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor) || !predicate(descriptor.value)) {
      return false
    }
  }

  return true
}

function mapArrayData<T, Result>(
  value: readonly T[],
  guard: (entry: unknown) => entry is T,
  mapper: (entry: T) => Result,
): readonly Result[] {
  const mapped: Result[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor) || !guard(descriptor.value)) {
      throw new Error('Workbook plan data arrays must contain only data properties')
    }
    mapped.push(mapper(descriptor.value))
  }
  return mapped
}

function isWorkbookRefDescription(value: unknown): value is WorkbookRefDescription {
  return isWorkbookRefData(value)
}

function isFormulaLabelData(value: unknown): value is WorkbookFormulaLabelDescription {
  return isRecord(value) && hasString(value, 'name') && isWorkbookRefDescription(ownValue(value, 'ref'))
}

function isFormulaLabelDataArray(value: unknown): value is readonly WorkbookFormulaLabelDescription[] {
  return arrayEvery(value, isFormulaLabelData)
}

function isRefDataArray(value: unknown): value is readonly WorkbookRefDescription[] {
  return arrayEvery(value, isWorkbookRefDescription)
}

function isWorkbookActionCommandData(value: unknown): value is WorkbookActionCommandDescription {
  if (!isRecord(value) || !hasString(value, 'kind')) {
    return false
  }
  switch (ownValue(value, 'kind')) {
    case 'writeFormula':
      return (
        isWorkbookRefDescription(ownValue(value, 'target')) &&
        typeof ownValue(value, 'formula') === 'string' &&
        isRefDataArray(ownValue(value, 'inputs')) &&
        isFormulaLabelDataArray(ownValue(value, 'labels'))
      )
    case 'writeValue':
      return isWorkbookRefDescription(ownValue(value, 'target')) && isLiteralInput(ownValue(value, 'value'))
    case 'format':
      return (
        isWorkbookRefDescription(ownValue(value, 'target')) &&
        (!hasOwnValue(value, 'style') || isRecord(ownValue(value, 'style'))) &&
        (!hasOwnValue(value, 'numberFormat') ||
          typeof ownValue(value, 'numberFormat') === 'string' ||
          ownValue(value, 'numberFormat') === null)
      )
    case 'clear':
      return isWorkbookRefDescription(ownValue(value, 'target'))
    case 'op':
      return (
        isWorkbookOp(ownValue(value, 'op')) &&
        (!hasOwnValue(value, 'target') || isWorkbookRefDescription(ownValue(value, 'target'))) &&
        hasOptionalString(value, 'message')
      )
    default:
      return false
  }
}

function isChangeData(value: unknown): value is WorkbookChangeSummaryDescription {
  return (
    isRecord(value) &&
    hasString(value, 'kind') &&
    hasString(value, 'message') &&
    (!hasOwnValue(value, 'target') || isWorkbookRefDescription(ownValue(value, 'target')))
  )
}

function isCheckStatus(value: unknown): value is WorkbookCheckStatus {
  return value === 'planned' || value === 'passed' || value === 'failed'
}

function isCheckExpectationData(value: unknown): value is WorkbookCheckExpectationDescription {
  if (!isRecord(value) || !hasString(value, 'kind')) {
    return false
  }
  switch (ownValue(value, 'kind')) {
    case 'valueEquals':
      return isLiteralInput(ownValue(value, 'value'))
    case 'formulaEquals':
      return (
        typeof ownValue(value, 'formula') === 'string' &&
        isRefDataArray(ownValue(value, 'inputs')) &&
        isFormulaLabelDataArray(ownValue(value, 'labels'))
      )
    default:
      return false
  }
}

function isCheckData(value: unknown): value is WorkbookCheckResultDescription {
  return (
    isRecord(value) &&
    isCheckStatus(ownValue(value, 'status')) &&
    hasString(value, 'kind') &&
    hasString(value, 'message') &&
    (!hasOwnValue(value, 'target') || isWorkbookRefDescription(ownValue(value, 'target'))) &&
    (!hasOwnValue(value, 'refs') || isRefDataArray(ownValue(value, 'refs'))) &&
    (!hasOwnValue(value, 'expectation') || isCheckExpectationData(ownValue(value, 'expectation'))) &&
    (!hasOwnValue(value, 'proof') || isWorkbookActionInput(ownValue(value, 'proof')))
  )
}

export function isPlanData(value: unknown): value is WorkbookPlanData {
  if (!isRecord(value)) {
    return false
  }
  const input = ownValue(value, 'input')
  const refsUsed = ownValue(value, 'refsUsed')
  const commands = ownValue(value, 'commands')
  const ops = ownValue(value, 'ops')
  const changed = ownValue(value, 'changed')
  const checks = ownValue(value, 'checks')

  return (
    hasString(value, 'modelName') &&
    hasString(value, 'actionName') &&
    (!hasOwnValue(value, 'input') || isWorkbookActionInput(input)) &&
    isRefDataArray(refsUsed) &&
    arrayEvery(commands, isWorkbookActionCommandData) &&
    arrayEvery(ops, isWorkbookOp) &&
    arrayEvery(changed, isChangeData) &&
    arrayEvery(checks, isCheckData)
  )
}

function planDataIssue(path: string, message: string): WorkbookPlanDataIssue {
  return Object.freeze({
    code: 'invalid_plan_data',
    path,
    message,
  })
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function inputPath(basePath: string, error: unknown): string {
  if (!(error instanceof WorkbookActionInputError)) {
    return basePath
  }
  if (error.path === 'input') {
    return basePath
  }
  if (error.path.startsWith('input.')) {
    return `${basePath}${error.path.slice('input'.length)}`
  }
  if (error.path.startsWith('input[')) {
    return `${basePath}${error.path.slice('input'.length)}`
  }
  return basePath
}

function pushRequiredStringIssue(issues: WorkbookPlanDataIssue[], value: Record<string, unknown>, key: string): void {
  if (typeof ownValue(value, key) !== 'string') {
    issues.push(planDataIssue(key, `Workbook plan data ${key} must be a string`))
  }
}

function pushOptionalInputIssue(issues: WorkbookPlanDataIssue[], value: Record<string, unknown>): void {
  const input = ownValue(value, 'input')
  if (!hasOwnValue(value, 'input')) {
    return
  }
  try {
    normalizeWorkbookActionInput(input)
  } catch (error) {
    issues.push(planDataIssue(inputPath('input', error), `Workbook plan data input must be JSON-safe: ${errorMessage(error)}`))
  }
}

function pushArrayIssues<T>(
  issues: WorkbookPlanDataIssue[],
  value: Record<string, unknown>,
  key: string,
  guard: (entry: unknown) => entry is T,
  label: string,
): void {
  const array = ownValue(value, key)
  if (!Array.isArray(array)) {
    issues.push(planDataIssue(key, `Workbook plan data ${key} must be an array`))
    return
  }
  for (let index = 0; index < array.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(array, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      issues.push(planDataIssue(`${key}[${index}]`, `Workbook plan data ${label} at ${key}[${index}] is invalid`))
      continue
    }
    const entry = descriptor.value
    if (!guard(entry)) {
      issues.push(planDataIssue(`${key}[${index}]`, `Workbook plan data ${label} at ${key}[${index}] is invalid`))
    }
  }
}

function pushCheckIssues(issues: WorkbookPlanDataIssue[], entry: unknown, index: number): void {
  const path = `checks[${index}]`
  if (isCheckData(entry)) {
    return
  }
  if (!isRecord(entry)) {
    issues.push(planDataIssue(path, `Workbook plan data check at ${path} is invalid`))
    return
  }
  if (
    !isCheckStatus(ownValue(entry, 'status')) ||
    !hasString(entry, 'kind') ||
    !hasString(entry, 'message') ||
    (hasOwnValue(entry, 'target') && !isWorkbookRefDescription(ownValue(entry, 'target'))) ||
    (hasOwnValue(entry, 'refs') && !isRefDataArray(ownValue(entry, 'refs'))) ||
    (hasOwnValue(entry, 'expectation') && !isCheckExpectationData(ownValue(entry, 'expectation')))
  ) {
    issues.push(planDataIssue(path, `Workbook plan data check at ${path} is invalid`))
    return
  }
  try {
    if (hasOwnValue(entry, 'proof')) {
      normalizeWorkbookActionInput(ownValue(entry, 'proof'))
    }
  } catch (error) {
    issues.push(
      planDataIssue(inputPath(`${path}.proof`, error), `Workbook plan data check proof must be JSON-safe: ${errorMessage(error)}`),
    )
    return
  }
  issues.push(planDataIssue(path, `Workbook plan data check at ${path} is invalid`))
}

function pushCheckArrayIssues(issues: WorkbookPlanDataIssue[], value: Record<string, unknown>): void {
  const checks = ownValue(value, 'checks')
  if (!Array.isArray(checks)) {
    issues.push(planDataIssue('checks', 'Workbook plan data checks must be an array'))
    return
  }
  for (let index = 0; index < checks.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(checks, String(index))
    const entry = descriptor !== undefined && descriptor.enumerable && 'value' in descriptor ? descriptor.value : undefined
    pushCheckIssues(issues, entry, index)
  }
}

export function checkPlanData(value: unknown): WorkbookPlanDataCheckResult {
  if (!isRecord(value)) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze([planDataIssue('plan', 'Workbook plan data must be an object')]),
    })
  }

  const issues: WorkbookPlanDataIssue[] = []
  pushRequiredStringIssue(issues, value, 'modelName')
  pushRequiredStringIssue(issues, value, 'actionName')
  pushOptionalInputIssue(issues, value)
  pushArrayIssues(issues, value, 'refsUsed', isWorkbookRefDescription, 'ref')
  pushArrayIssues(issues, value, 'commands', isWorkbookActionCommandData, 'command')
  pushArrayIssues(issues, value, 'ops', isWorkbookOp, 'op')
  pushArrayIssues(issues, value, 'changed', isChangeData, 'change')
  pushCheckArrayIssues(issues, value)

  if (issues.length > 0) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze(issues),
    })
  }
  if (!isPlanData(value)) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze([planDataIssue('plan', 'Workbook plan data is invalid')]),
    })
  }

  return Object.freeze({
    status: 'valid',
    plan: value,
    issues: Object.freeze([] as const),
  })
}

function hydrateRef(ref: WorkbookRefDescription): WorkbookRef {
  return hydrateWorkbookRef(ref)
}

function hydrateFormulaLabel(label: WorkbookFormulaLabelDescription): WorkbookFormulaLabel {
  return Object.freeze({
    name: label.name,
    ref: hydrateRef(label.ref),
  })
}

function hydrateCommand(command: WorkbookActionCommandDescription): WorkbookActionCommand {
  switch (command.kind) {
    case 'writeFormula':
      return Object.freeze({
        kind: 'writeFormula',
        target: hydrateRef(command.target),
        formula: command.formula,
        inputs: Object.freeze(mapArrayData(command.inputs, isWorkbookRefDescription, hydrateRef)),
        labels: Object.freeze(mapArrayData(command.labels, isFormulaLabelData, hydrateFormulaLabel)),
      })
    case 'writeValue':
      return Object.freeze({
        kind: 'writeValue',
        target: hydrateRef(command.target),
        value: command.value,
      })
    case 'format':
      return Object.freeze({
        kind: 'format',
        target: hydrateRef(command.target),
        ...(command.style !== undefined ? { style: Object.freeze(structuredClone(command.style)) } : {}),
        ...(command.numberFormat !== undefined ? { numberFormat: command.numberFormat } : {}),
      })
    case 'clear':
      return Object.freeze({
        kind: 'clear',
        target: hydrateRef(command.target),
      })
    case 'op':
      return Object.freeze({
        kind: 'op',
        op: Object.freeze(structuredClone(command.op)),
        ...(command.target !== undefined ? { target: hydrateRef(command.target) } : {}),
        ...(command.message !== undefined ? { message: command.message } : {}),
      })
  }
}

function hydrateChange(change: WorkbookChangeSummaryDescription): WorkbookChangeSummary {
  return Object.freeze({
    kind: change.kind,
    ...(change.target !== undefined ? { target: hydrateRef(change.target) } : {}),
    message: change.message,
  })
}

function hydrateExpectation(expectation: WorkbookCheckExpectationDescription): WorkbookCheckExpectation {
  switch (expectation.kind) {
    case 'valueEquals':
      return Object.freeze({
        kind: 'valueEquals',
        value: expectation.value,
      })
    case 'formulaEquals':
      return Object.freeze({
        kind: 'formulaEquals',
        formula: expectation.formula,
        inputs: Object.freeze(mapArrayData(expectation.inputs, isWorkbookRefDescription, hydrateRef)),
        labels: Object.freeze(mapArrayData(expectation.labels, isFormulaLabelData, hydrateFormulaLabel)),
      })
  }
}

function hydrateCheck(check: WorkbookCheckResultDescription): WorkbookCheckResult {
  const proof: WorkbookActionInput | undefined = check.proof === undefined ? undefined : normalizeWorkbookActionInput(check.proof)
  return Object.freeze({
    status: check.status,
    kind: check.kind,
    ...(check.target !== undefined ? { target: hydrateRef(check.target) } : {}),
    ...(check.refs !== undefined ? { refs: Object.freeze(mapArrayData(check.refs, isWorkbookRefDescription, hydrateRef)) } : {}),
    message: check.message,
    ...(check.expectation !== undefined ? { expectation: hydrateExpectation(check.expectation) } : {}),
    ...(proof !== undefined ? { proof } : {}),
  })
}

function cloneOp(op: WorkbookOp): WorkbookOp {
  return Object.freeze(structuredClone(op))
}

export function toPlanData<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookPlanData {
  return describePlan(plan)
}

export function hydratePlanData(data: unknown): WorkbookActionPlan<WorkbookPlanDataRefs> {
  const check = checkPlanData(data)
  if (check.status === 'invalid') {
    const [firstIssue] = check.issues
    throw new Error(firstIssue === undefined ? 'Workbook plan data is invalid' : `Workbook plan data is invalid: ${firstIssue.message}`)
  }
  const plan = check.plan
  const refsUsed = Object.freeze(mapArrayData(plan.refsUsed, isWorkbookRefDescription, hydrateRef))
  const input = plan.input === undefined ? undefined : normalizeWorkbookActionInput(plan.input)
  return Object.freeze({
    modelName: plan.modelName,
    actionName: plan.actionName,
    ...(input !== undefined ? { input } : {}),
    refs: Object.freeze({ refsUsed }),
    refsUsed,
    commands: Object.freeze(mapArrayData(plan.commands, isWorkbookActionCommandData, hydrateCommand)),
    ops: Object.freeze(mapArrayData(plan.ops, isWorkbookOp, cloneOp)),
    changed: Object.freeze(mapArrayData(plan.changed, isChangeData, hydrateChange)),
    checks: Object.freeze(mapArrayData(plan.checks, isCheckData, hydrateCheck)),
  })
}

export function isHydratedPlan<Refs>(value: unknown): value is WorkbookActionPlan<Refs> {
  return isRecord(value) && isRecord(ownValue(value, 'refs'))
}

export function executablePlan<Refs>(
  plan: WorkbookExecutablePlan<Refs>,
): WorkbookActionPlan<Refs> | WorkbookActionPlan<WorkbookPlanDataRefs> {
  return isHydratedPlan<Refs>(plan) ? plan : hydratePlanData(plan)
}
