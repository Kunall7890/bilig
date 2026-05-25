import { describeRef, type WorkbookRefDescription } from './describe.js'
import { isWorkbookRefData, toWorkbookRefData, type WorkbookRef } from './find.js'
import type { WorkbookActionCommand, WorkbookActionPlan } from './model.js'
import type { WorkbookOp } from './ops.js'
import { hydratePlanData, isHydratedPlan, type WorkbookExecutablePlan } from './plan-data.js'
import type { WorkbookCheckResult } from './result.js'

type WorkbookConcreteCommandOp = Extract<WorkbookOp, { kind: 'setCellFormula' | 'setCellValue' | 'setCellFormat' | 'clearCell' }>

export type WorkbookRuntimeRequirementKind = 'apply' | 'read' | 'verify'

export type WorkbookRuntimeCapability = 'writeFormula' | 'writeValue' | 'format' | 'clear' | 'applyOp' | 'read' | 'verifyCheck'

export const workbookRuntimeRequirementKinds = Object.freeze([
  'apply',
  'read',
  'verify',
] as const satisfies readonly WorkbookRuntimeRequirementKind[])
const WORKBOOK_RUNTIME_REQUIREMENT_KIND_SET = new Set<string>(workbookRuntimeRequirementKinds)

export const workbookRuntimeCapabilities = Object.freeze([
  'writeFormula',
  'writeValue',
  'format',
  'clear',
  'applyOp',
  'read',
  'verifyCheck',
] as const satisfies readonly WorkbookRuntimeCapability[])
const WORKBOOK_RUNTIME_CAPABILITY_SET = new Set<string>(workbookRuntimeCapabilities)

export interface WorkbookRuntimeRequirement {
  readonly kind: WorkbookRuntimeRequirementKind
  readonly capability: WorkbookRuntimeCapability
  readonly message: string
  readonly commandIndex?: number
  readonly checkIndex?: number
  readonly opIndex?: number
  readonly opKind?: string
  readonly checkKind?: string
  readonly target?: WorkbookRefDescription
  readonly refs?: readonly WorkbookRefDescription[]
}

export interface WorkbookRuntimeRequirements {
  readonly modelName: string
  readonly actionName: string
  readonly requirements: readonly WorkbookRuntimeRequirement[]
}

export type WorkbookRuntimeRequirementsIssueCode = 'invalid_runtime_requirements'

export interface WorkbookRuntimeRequirementsIssue {
  readonly code: WorkbookRuntimeRequirementsIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookRuntimeRequirementsCheckResult =
  | {
      readonly status: 'valid'
      readonly requirements: WorkbookRuntimeRequirements
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookRuntimeRequirementsIssue[]
    }

export type WorkbookRuntimeAdapterIssueCode = 'invalid_requirements' | 'missing_apply' | 'missing_read' | 'missing_check_verifier'
export type WorkbookRuntimeAdapterMethod = 'apply' | 'read' | 'verifyChecks'

export interface WorkbookRuntimeAdapterIssue {
  readonly code: WorkbookRuntimeAdapterIssueCode
  readonly capability?: WorkbookRuntimeCapability
  readonly method?: WorkbookRuntimeAdapterMethod
  readonly requirementIndexes: readonly number[]
  readonly path?: string
  readonly message: string
}

export type WorkbookRuntimeAdapterCheckResult =
  | {
      readonly status: 'valid'
      readonly modelName: string
      readonly actionName: string
      readonly requiredCapabilities: readonly WorkbookRuntimeCapability[]
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly modelName: string
      readonly actionName: string
      readonly requiredCapabilities: readonly WorkbookRuntimeCapability[]
      readonly issues: readonly WorkbookRuntimeAdapterIssue[]
    }

export interface WorkbookRuntimeAdapterCandidate {
  readonly apply?: unknown
  readonly read?: unknown
  readonly verifyChecks?: unknown
}

export function isWorkbookRuntimeRequirementKind(value: unknown): value is WorkbookRuntimeRequirementKind {
  return typeof value === 'string' && WORKBOOK_RUNTIME_REQUIREMENT_KIND_SET.has(value)
}

export function isWorkbookRuntimeCapability(value: unknown): value is WorkbookRuntimeCapability {
  return typeof value === 'string' && WORKBOOK_RUNTIME_CAPABILITY_SET.has(value)
}

function describedRef(ref: WorkbookRef | undefined): { readonly target: WorkbookRefDescription } | {} {
  return ref === undefined ? {} : { target: describeRef(ref) }
}

function describedRefs(refs: readonly WorkbookRef[] | undefined): { readonly refs: readonly WorkbookRefDescription[] } | {} {
  return refs === undefined || refs.length === 0 ? {} : { refs: refs.map(describeRef) }
}

function requiredStringValue(value: object, key: string): string {
  const entry = ownValue(value, key)
  if (typeof entry !== 'string') {
    throw new Error(`Workbook runtime requirements ${key} must be a string`)
  }
  return entry
}

function optionalIndexValue(value: object, key: 'commandIndex' | 'checkIndex' | 'opIndex'): number | undefined {
  if (!hasOwnValue(value, key)) {
    return undefined
  }
  const entry = ownValue(value, key)
  if (typeof entry !== 'number' || !Number.isInteger(entry) || entry < 0) {
    throw new Error(`Workbook runtime requirement ${key} must be a non-negative integer`)
  }
  return entry
}

function optionalStringValue(value: object, key: 'opKind' | 'checkKind'): string | undefined {
  if (!hasOwnValue(value, key)) {
    return undefined
  }
  const entry = ownValue(value, key)
  if (typeof entry !== 'string') {
    throw new Error(`Workbook runtime requirement ${key} must be a string`)
  }
  return entry
}

function normalizedRefDescription(ref: unknown): WorkbookRefDescription {
  if (!isWorkbookRefData(ref)) {
    throw new Error('Workbook runtime requirement ref must be workbook ref data')
  }
  return toWorkbookRefData(ref)
}

function normalizeRuntimeRequirement(requirement: WorkbookRuntimeRequirement): WorkbookRuntimeRequirement {
  const kind = ownValue(requirement, 'kind')
  const capability = ownValue(requirement, 'capability')
  if (!isWorkbookRuntimeRequirementKind(kind) || !isWorkbookRuntimeCapability(capability)) {
    throw new Error('Workbook runtime requirement is invalid')
  }

  const commandIndex = optionalIndexValue(requirement, 'commandIndex')
  const checkIndex = optionalIndexValue(requirement, 'checkIndex')
  const opIndex = optionalIndexValue(requirement, 'opIndex')
  const opKind = optionalStringValue(requirement, 'opKind')
  const checkKind = optionalStringValue(requirement, 'checkKind')
  const target = ownValue(requirement, 'target')
  const refs = ownValue(requirement, 'refs')
  const normalizedRefs = refs !== undefined && Array.isArray(refs) ? Object.freeze(refs.map(normalizedRefDescription)) : undefined

  return Object.freeze({
    kind,
    capability,
    ...(commandIndex !== undefined ? { commandIndex } : {}),
    ...(checkIndex !== undefined ? { checkIndex } : {}),
    ...(opIndex !== undefined ? { opIndex } : {}),
    ...(opKind !== undefined ? { opKind } : {}),
    ...(checkKind !== undefined ? { checkKind } : {}),
    ...(target !== undefined ? { target: normalizedRefDescription(target) } : {}),
    ...(normalizedRefs !== undefined ? { refs: normalizedRefs } : {}),
    message: requiredStringValue(requirement, 'message'),
  })
}

function normalizeRuntimeRequirements(requirements: WorkbookRuntimeRequirements): WorkbookRuntimeRequirements {
  const entries = ownValue(requirements, 'requirements')
  if (!arrayEveryData(entries, isRuntimeRequirement)) {
    throw new Error('Workbook runtime requirements are invalid')
  }
  return Object.freeze({
    modelName: requiredStringValue(requirements, 'modelName'),
    actionName: requiredStringValue(requirements, 'actionName'),
    requirements: Object.freeze(entries.map(normalizeRuntimeRequirement)),
  })
}

function commandCapability(command: WorkbookActionCommand): WorkbookRuntimeCapability {
  switch (command.kind) {
    case 'writeFormula':
      return 'writeFormula'
    case 'writeValue':
      return 'writeValue'
    case 'format':
      return 'format'
    case 'clear':
      return 'clear'
    case 'op':
      return 'applyOp'
  }
}

function commandRefs(command: WorkbookActionCommand): readonly WorkbookRef[] | undefined {
  return command.kind === 'writeFormula' ? command.inputs : undefined
}

function commandMessage(command: WorkbookActionCommand): string {
  switch (command.kind) {
    case 'writeFormula':
      return `Apply formula write to ${command.target.label}`
    case 'writeValue':
      return `Apply value write to ${command.target.label}`
    case 'format':
      return `Apply format to ${command.target.label}`
    case 'clear':
      return `Apply clear to ${command.target.label}`
    case 'op':
      return command.target === undefined
        ? `Apply workbook op ${command.op.kind}`
        : `Apply workbook op ${command.op.kind} to ${command.target.label}`
  }
}

function commandRequirement(command: WorkbookActionCommand, commandIndex: number): WorkbookRuntimeRequirement {
  return {
    kind: 'apply',
    capability: commandCapability(command),
    commandIndex,
    ...(command.kind === 'op' ? { opKind: command.op.kind } : {}),
    ...describedRef(command.target),
    ...describedRefs(commandRefs(command)),
    message: commandMessage(command),
  }
}

function concreteSingleCell(target: WorkbookRef): { readonly sheetName: string; readonly address: string } | null {
  if (target.kind !== 'range') {
    return null
  }
  const range = target.range
  return range.startAddress === range.endAddress ? { sheetName: range.sheetName, address: range.startAddress } : null
}

function commandConcreteOp(command: WorkbookActionCommand): WorkbookConcreteCommandOp | null {
  if (command.kind === 'op') {
    return command.op.kind === 'setCellFormula' ||
      command.op.kind === 'setCellValue' ||
      command.op.kind === 'setCellFormat' ||
      command.op.kind === 'clearCell'
      ? command.op
      : null
  }

  const target = concreteSingleCell(command.target)
  if (target === null) {
    return null
  }

  switch (command.kind) {
    case 'writeFormula':
      return {
        kind: 'setCellFormula',
        sheetName: target.sheetName,
        address: target.address,
        formula: command.formula,
      }
    case 'writeValue':
      return {
        kind: 'setCellValue',
        sheetName: target.sheetName,
        address: target.address,
        value: command.value,
      }
    case 'format':
      if (command.numberFormat === undefined) {
        return null
      }
      return {
        kind: 'setCellFormat',
        sheetName: target.sheetName,
        address: target.address,
        format: command.numberFormat,
      }
    case 'clear':
      return {
        kind: 'clearCell',
        sheetName: target.sheetName,
        address: target.address,
      }
  }
}

function readRequirement(check: WorkbookCheckResult, checkIndex: number): WorkbookRuntimeRequirement | null {
  if (check.expectation === undefined) {
    return null
  }
  return {
    kind: 'read',
    capability: 'read',
    checkIndex,
    checkKind: check.kind,
    ...describedRef(check.target),
    ...describedRefs(check.expectation.kind === 'formulaEquals' ? check.expectation.inputs : undefined),
    message: `Read ${check.target?.label ?? check.kind} for ${check.expectation.kind}`,
  }
}

function verifyRequirement(check: WorkbookCheckResult, checkIndex: number): WorkbookRuntimeRequirement | null {
  if (check.expectation !== undefined) {
    return null
  }
  return {
    kind: 'verify',
    capability: 'verifyCheck',
    checkIndex,
    checkKind: check.kind,
    ...describedRef(check.target),
    ...describedRefs(check.refs),
    message: `Verify ${check.kind} for ${check.target?.label ?? check.kind}`,
  }
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalValue)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalValue(entry)]),
    )
  }
  return value
}

function opKey(op: WorkbookOp): string {
  return JSON.stringify(canonicalValue(op))
}

function commandCoveredOpKeys(commands: readonly WorkbookActionCommand[]): ReadonlySet<string> {
  const keys = new Set<string>()
  commands.forEach((command) => {
    if (command.kind === 'op') {
      keys.add(opKey(command.op))
      return
    }
    const concreteOp = commandConcreteOp(command)
    if (concreteOp !== null) {
      keys.add(opKey(concreteOp))
    }
  })
  return keys
}

function describeLiveRuntimeRequirements<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRuntimeRequirements {
  const requirements: WorkbookRuntimeRequirement[] = plan.commands.map(commandRequirement)
  const commandCoveredOps = commandCoveredOpKeys(plan.commands)

  plan.ops.forEach((op, opIndex) => {
    if (commandCoveredOps.has(opKey(op))) {
      return
    }
    requirements.push({
      kind: 'apply',
      capability: 'applyOp',
      opIndex,
      opKind: op.kind,
      message: `Apply workbook op ${op.kind}`,
    })
  })

  plan.checks.forEach((check, checkIndex) => {
    const read = readRequirement(check, checkIndex)
    if (read !== null) {
      requirements.push(read)
    }
  })

  plan.checks.forEach((check, checkIndex) => {
    const verify = verifyRequirement(check, checkIndex)
    if (verify !== null) {
      requirements.push(verify)
    }
  })

  return normalizeRuntimeRequirements({
    modelName: plan.modelName,
    actionName: plan.actionName,
    requirements,
  })
}

export function describeRuntimeRequirements<Refs>(plan: WorkbookExecutablePlan<Refs>): WorkbookRuntimeRequirements {
  if (isHydratedPlan(plan)) {
    return describeLiveRuntimeRequirements(plan)
  }
  return describeLiveRuntimeRequirements(hydratePlanData(plan))
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
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

function arrayEveryData<T>(value: unknown, predicate: (entry: unknown) => entry is T): value is readonly T[]
function arrayEveryData(value: unknown, predicate: (entry: unknown) => boolean): boolean
function arrayEveryData(value: unknown, predicate: (entry: unknown) => boolean): boolean {
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

function hasOptionalIndex(value: Record<string, unknown>, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  const entry = descriptor?.value
  return descriptor === undefined || (Number.isInteger(entry) && typeof entry === 'number' && entry >= 0)
}

function isWorkbookRefDescription(value: unknown): value is WorkbookRefDescription {
  return isWorkbookRefData(value)
}

function isRefDescriptionArray(value: unknown): value is readonly WorkbookRefDescription[] {
  return arrayEveryData(value, isWorkbookRefDescription)
}

function capabilityMatchesKind(kind: WorkbookRuntimeRequirementKind, capability: WorkbookRuntimeCapability): boolean {
  switch (kind) {
    case 'apply':
      return (
        capability === 'writeFormula' ||
        capability === 'writeValue' ||
        capability === 'format' ||
        capability === 'clear' ||
        capability === 'applyOp'
      )
    case 'read':
      return capability === 'read'
    case 'verify':
      return capability === 'verifyCheck'
  }
}

function isRuntimeRequirement(value: unknown): value is WorkbookRuntimeRequirement {
  if (!isRecord(value)) {
    return false
  }
  const kind = ownValue(value, 'kind')
  const capability = ownValue(value, 'capability')
  if (!isWorkbookRuntimeRequirementKind(kind) || !isWorkbookRuntimeCapability(capability)) {
    return false
  }
  return (
    capabilityMatchesKind(kind, capability) &&
    hasString(value, 'message') &&
    hasOptionalIndex(value, 'commandIndex') &&
    hasOptionalIndex(value, 'checkIndex') &&
    hasOptionalIndex(value, 'opIndex') &&
    hasOptionalString(value, 'opKind') &&
    hasOptionalString(value, 'checkKind') &&
    (!hasOwnValue(value, 'target') || isWorkbookRefDescription(ownValue(value, 'target'))) &&
    (!hasOwnValue(value, 'refs') || isRefDescriptionArray(ownValue(value, 'refs')))
  )
}

function isRuntimeRequirements(value: unknown): value is WorkbookRuntimeRequirements {
  if (!isRecord(value)) {
    return false
  }
  const requirements = ownValue(value, 'requirements')

  return hasString(value, 'modelName') && hasString(value, 'actionName') && arrayEveryData(requirements, isRuntimeRequirement)
}

function isRuntimeRequirementsCandidate<Refs>(
  value: WorkbookExecutablePlan<Refs> | WorkbookRuntimeRequirements,
): value is WorkbookRuntimeRequirements {
  return isRecord(value) && Object.getOwnPropertyDescriptor(value, 'requirements') !== undefined
}

function runtimeRequirementsIssue(path: string, message: string): WorkbookRuntimeRequirementsIssue {
  return Object.freeze({
    code: 'invalid_runtime_requirements',
    path,
    message,
  })
}

function pushRequiredStringIssue(issues: WorkbookRuntimeRequirementsIssue[], value: Record<string, unknown>, key: string): void {
  if (typeof ownValue(value, key) !== 'string') {
    issues.push(runtimeRequirementsIssue(key, `Workbook runtime requirements ${key} must be a string`))
  }
}

function pushOptionalIndexIssue(
  issues: WorkbookRuntimeRequirementsIssue[],
  value: Record<string, unknown>,
  key: 'commandIndex' | 'checkIndex' | 'opIndex',
  path: string,
): void {
  if (!hasOptionalIndex(value, key)) {
    issues.push(runtimeRequirementsIssue(`${path}.${key}`, `Workbook runtime requirement ${key} must be a non-negative integer`))
  }
}

function pushOptionalStringIssue(
  issues: WorkbookRuntimeRequirementsIssue[],
  value: Record<string, unknown>,
  key: 'opKind' | 'checkKind',
  path: string,
): void {
  if (!hasOptionalString(value, key)) {
    issues.push(runtimeRequirementsIssue(`${path}.${key}`, `Workbook runtime requirement ${key} must be a string`))
  }
}

function pushRequirementIssues(issues: WorkbookRuntimeRequirementsIssue[], value: unknown, index: number): void {
  const path = `requirements[${index}]`
  if (!isRecord(value)) {
    issues.push(runtimeRequirementsIssue(path, `Workbook runtime requirement at ${path} must be an object`))
    return
  }

  const kind = ownValue(value, 'kind')
  const capability = ownValue(value, 'capability')
  if (!isWorkbookRuntimeRequirementKind(kind)) {
    issues.push(runtimeRequirementsIssue(`${path}.kind`, 'Workbook runtime requirement kind is invalid'))
  }
  if (!isWorkbookRuntimeCapability(capability)) {
    issues.push(runtimeRequirementsIssue(`${path}.capability`, 'Workbook runtime requirement capability is invalid'))
  }
  if (isWorkbookRuntimeRequirementKind(kind) && isWorkbookRuntimeCapability(capability) && !capabilityMatchesKind(kind, capability)) {
    issues.push(
      runtimeRequirementsIssue(`${path}.capability`, `Workbook runtime requirement capability ${capability} does not match kind ${kind}`),
    )
  }
  if (!hasString(value, 'message')) {
    issues.push(runtimeRequirementsIssue(`${path}.message`, 'Workbook runtime requirement message must be a string'))
  }
  pushOptionalIndexIssue(issues, value, 'commandIndex', path)
  pushOptionalIndexIssue(issues, value, 'checkIndex', path)
  pushOptionalIndexIssue(issues, value, 'opIndex', path)
  pushOptionalStringIssue(issues, value, 'opKind', path)
  pushOptionalStringIssue(issues, value, 'checkKind', path)
  if (hasOwnValue(value, 'target') && !isWorkbookRefDescription(ownValue(value, 'target'))) {
    issues.push(runtimeRequirementsIssue(`${path}.target`, 'Workbook runtime requirement target must be workbook ref data'))
  }
  const refs = ownValue(value, 'refs')
  if (hasOwnValue(value, 'refs')) {
    if (!Array.isArray(refs)) {
      issues.push(runtimeRequirementsIssue(`${path}.refs`, 'Workbook runtime requirement refs must be an array'))
      return
    }
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const descriptor = Object.getOwnPropertyDescriptor(refs, String(refIndex))
      const ref = descriptor !== undefined && descriptor.enumerable && 'value' in descriptor ? descriptor.value : undefined
      if (!isWorkbookRefDescription(ref)) {
        issues.push(runtimeRequirementsIssue(`${path}.refs[${refIndex}]`, 'Workbook runtime requirement ref must be workbook ref data'))
      }
    }
  }
}

export function checkRuntimeRequirements(value: unknown): WorkbookRuntimeRequirementsCheckResult {
  if (!isRecord(value)) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze([runtimeRequirementsIssue('requirements', 'Workbook runtime requirements must be an object')]),
    })
  }

  const issues: WorkbookRuntimeRequirementsIssue[] = []
  pushRequiredStringIssue(issues, value, 'modelName')
  pushRequiredStringIssue(issues, value, 'actionName')
  const requirements = ownValue(value, 'requirements')
  if (!Array.isArray(requirements)) {
    issues.push(runtimeRequirementsIssue('requirements', 'Workbook runtime requirements requirements must be an array'))
  } else {
    for (let index = 0; index < requirements.length; index += 1) {
      const descriptor = Object.getOwnPropertyDescriptor(requirements, String(index))
      const requirement = descriptor !== undefined && descriptor.enumerable && 'value' in descriptor ? descriptor.value : undefined
      pushRequirementIssues(issues, requirement, index)
    }
  }

  if (issues.length > 0) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze(issues),
    })
  }
  if (!isRuntimeRequirements(value)) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze([runtimeRequirementsIssue('requirements', 'Workbook runtime requirements are invalid')]),
    })
  }

  return Object.freeze({
    status: 'valid',
    requirements: normalizeRuntimeRequirements(value),
    issues: Object.freeze([] as const),
  })
}

type RuntimeRequirementsResult =
  | {
      readonly status: 'valid'
      readonly requirements: WorkbookRuntimeRequirements
    }
  | {
      readonly status: 'invalid'
      readonly modelName: string
      readonly actionName: string
      readonly issues: readonly WorkbookRuntimeAdapterIssue[]
    }

function safeString(value: unknown, key: string, fallback: string): string {
  if (!isRecord(value)) {
    return fallback
  }
  const entry = ownValue(value, key)
  return typeof entry === 'string' && entry.trim() !== '' ? entry : fallback
}

function invalidRequirementsAdapterIssue(path: string, message: string): WorkbookRuntimeAdapterIssue {
  return Object.freeze({
    code: 'invalid_requirements',
    path,
    requirementIndexes: Object.freeze([] as const),
    message,
  })
}

function invalidRequirementsResult<Refs>(
  input: WorkbookExecutablePlan<Refs> | WorkbookRuntimeRequirements,
  issues: readonly WorkbookRuntimeAdapterIssue[],
): RuntimeRequirementsResult {
  return Object.freeze({
    status: 'invalid',
    modelName: safeString(input, 'modelName', 'unknown-model'),
    actionName: safeString(input, 'actionName', 'unknown-action'),
    issues: Object.freeze([...issues]),
  })
}

function requirementsFor<Refs>(input: WorkbookExecutablePlan<Refs> | WorkbookRuntimeRequirements): RuntimeRequirementsResult {
  if (isRuntimeRequirementsCandidate(input)) {
    const check = checkRuntimeRequirements(input)
    if (check.status === 'invalid') {
      return invalidRequirementsResult(
        input,
        check.issues.map((issue) => invalidRequirementsAdapterIssue(issue.path, issue.message)),
      )
    }
    return Object.freeze({
      status: 'valid',
      requirements: check.requirements,
    })
  }
  try {
    return Object.freeze({
      status: 'valid',
      requirements: describeRuntimeRequirements(input),
    })
  } catch (error) {
    return invalidRequirementsResult(input, [
      invalidRequirementsAdapterIssue(
        'plan',
        error instanceof Error ? error.message : `Workbook runtime requirements are invalid: ${String(error)}`,
      ),
    ])
  }
}

function pushCapability(capabilities: WorkbookRuntimeCapability[], capability: WorkbookRuntimeCapability): void {
  if (!capabilities.includes(capability)) {
    capabilities.push(capability)
  }
}

function requirementIndexesFor(
  requirements: readonly WorkbookRuntimeRequirement[],
  capability: WorkbookRuntimeCapability,
): readonly number[] {
  return requirements.flatMap((requirement, index) => (requirement.capability === capability ? [index] : []))
}

function hasMethod(adapter: WorkbookRuntimeAdapterCandidate, method: WorkbookRuntimeAdapterMethod): boolean {
  if (!isRecord(adapter)) {
    return false
  }
  const descriptor = Object.getOwnPropertyDescriptor(adapter, method)
  return descriptor !== undefined && 'value' in descriptor && typeof descriptor.value === 'function'
}

function adapterIssue(
  requirements: readonly WorkbookRuntimeRequirement[],
  code: WorkbookRuntimeAdapterIssueCode,
  capability: WorkbookRuntimeCapability,
  method: WorkbookRuntimeAdapterMethod,
): WorkbookRuntimeAdapterIssue {
  return Object.freeze({
    code,
    capability,
    method,
    requirementIndexes: Object.freeze(requirementIndexesFor(requirements, capability)),
    message: `Adapter is missing ${method} for ${capability}`,
  })
}

export function checkRuntimeAdapter<Refs>(
  planOrRequirements: WorkbookExecutablePlan<Refs> | WorkbookRuntimeRequirements,
  adapter: WorkbookRuntimeAdapterCandidate,
): WorkbookRuntimeAdapterCheckResult {
  const requirementsResult = requirementsFor(planOrRequirements)
  if (requirementsResult.status === 'invalid') {
    return Object.freeze({
      status: 'invalid',
      modelName: requirementsResult.modelName,
      actionName: requirementsResult.actionName,
      requiredCapabilities: Object.freeze([] as const),
      issues: requirementsResult.issues,
    })
  }

  const requirements = requirementsResult.requirements
  const requiredCapabilities: WorkbookRuntimeCapability[] = []
  requirements.requirements.forEach((requirement) => {
    pushCapability(requiredCapabilities, requirement.capability)
  })

  const issues: WorkbookRuntimeAdapterIssue[] = []
  const needsApply = requiredCapabilities.some(
    (capability) =>
      capability === 'writeFormula' ||
      capability === 'writeValue' ||
      capability === 'format' ||
      capability === 'clear' ||
      capability === 'applyOp',
  )
  if (needsApply && !hasMethod(adapter, 'apply')) {
    const firstApplyCapability = requiredCapabilities.find(
      (capability) =>
        capability === 'writeFormula' ||
        capability === 'writeValue' ||
        capability === 'format' ||
        capability === 'clear' ||
        capability === 'applyOp',
    )
    if (firstApplyCapability !== undefined) {
      issues.push(adapterIssue(requirements.requirements, 'missing_apply', firstApplyCapability, 'apply'))
    }
  }
  if (requiredCapabilities.includes('read') && !hasMethod(adapter, 'read')) {
    issues.push(adapterIssue(requirements.requirements, 'missing_read', 'read', 'read'))
  }
  if (requiredCapabilities.includes('verifyCheck') && !hasMethod(adapter, 'verifyChecks')) {
    issues.push(adapterIssue(requirements.requirements, 'missing_check_verifier', 'verifyCheck', 'verifyChecks'))
  }

  const shared = {
    modelName: requirements.modelName,
    actionName: requirements.actionName,
    requiredCapabilities: Object.freeze(requiredCapabilities),
  }
  if (issues.length > 0) {
    return Object.freeze({
      status: 'invalid',
      ...shared,
      issues: Object.freeze(issues),
    })
  }
  return Object.freeze({
    status: 'valid',
    ...shared,
    issues: Object.freeze([] as const),
  })
}
