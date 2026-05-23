import type { CellStylePatch, LiteralInput } from '@bilig/protocol'
import { formula, type WorkbookFormulaLabel, type WorkbookFormulaOperand } from './formula.js'
import { collectWorkbookRefs, createWorkbookFindApi, type WorkbookFindApi, type WorkbookRef } from './find.js'
import { createWorkbookCheckApi, type WorkbookCheckApi } from './check.js'
import {
  checkInput,
  WorkbookActionInputError,
  normalizeOptionalWorkbookActionInput,
  normalizeWorkbookActionInput,
  normalizeWorkbookActionInputDescription,
  type WorkbookActionInput,
  type WorkbookActionInputIssue,
  type WorkbookActionInputDescription,
} from './input.js'
import { isWorkbookOp } from './guards.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookChangeSummary, WorkbookCheckResult, WorkbookRunError, WorkbookRunErrorCode } from './result.js'

export type WorkbookActionCommand =
  | {
      readonly kind: 'writeFormula'
      readonly target: WorkbookRef
      readonly formula: string
      readonly inputs: readonly WorkbookRef[]
      readonly labels: readonly WorkbookFormulaLabel[]
    }
  | {
      readonly kind: 'writeValue'
      readonly target: WorkbookRef
      readonly value: LiteralInput
    }
  | {
      readonly kind: 'format'
      readonly target: WorkbookRef
      readonly style?: CellStylePatch
      readonly numberFormat?: string | null
    }
  | {
      readonly kind: 'clear'
      readonly target: WorkbookRef
    }
  | {
      readonly kind: 'op'
      readonly op: WorkbookOp
      readonly target?: WorkbookRef
      readonly message?: string
    }

export interface WorkbookAddOpOptions {
  readonly target?: WorkbookRef
  readonly message?: string
}

export interface WorkbookFindWorkbook extends WorkbookFindApi {}

export interface WorkbookCheckWorkbook extends WorkbookFindApi {
  readonly check: WorkbookCheckApi
}

export interface WorkbookActionWorkbook extends WorkbookCheckWorkbook {
  readonly writeFormula: (target: WorkbookRef, value: WorkbookFormulaOperand) => void
  readonly writeValue: (target: WorkbookRef, value: LiteralInput) => void
  readonly format: (target: WorkbookRef, options: { readonly style?: CellStylePatch; readonly numberFormat?: string | null }) => void
  readonly clear: (target: WorkbookRef) => void
  readonly addOp: (op: WorkbookOp, options?: WorkbookAddOpOptions) => void
}

export interface WorkbookModelWorkbook extends WorkbookActionWorkbook {}

export interface WorkbookActionContext<Refs> {
  readonly refs: Refs
  readonly workbook: WorkbookActionWorkbook
  readonly input?: WorkbookActionInput
}

export interface WorkbookCheckContext<Refs> {
  readonly refs: Refs
  readonly workbook: WorkbookCheckWorkbook
  readonly input?: WorkbookActionInput
}

export type WorkbookAction<Refs> = (context: WorkbookActionContext<Refs>) => void

export interface WorkbookActionConfig<Refs> {
  readonly description?: string
  readonly input?: WorkbookActionInputDescription
  readonly run: WorkbookAction<Refs>
}

export type WorkbookActionDefinition<Refs> = WorkbookAction<Refs> | WorkbookActionConfig<Refs>

export type WorkbookActionMap<Refs> = Record<string, WorkbookActionDefinition<Refs>>

export interface WorkbookModelConfig<Refs, Actions extends WorkbookActionMap<Refs>> {
  readonly name: string
  readonly description?: string
  readonly find: (workbook: WorkbookFindWorkbook) => Refs
  readonly checks?: (context: WorkbookCheckContext<Refs>) => readonly WorkbookCheckResult[]
  readonly actions: Actions
}

export interface WorkbookModel<
  Refs = unknown,
  Actions extends WorkbookActionMap<Refs> = WorkbookActionMap<Refs>,
> extends WorkbookModelConfig<Refs, Actions> {}

export interface WorkbookActionPlan<Refs = unknown> {
  readonly modelName: string
  readonly actionName: string
  readonly input?: WorkbookActionInput
  readonly refs: Refs
  readonly refsUsed: readonly WorkbookRef[]
  readonly commands: readonly WorkbookActionCommand[]
  readonly ops: readonly WorkbookOp[]
  readonly changed: readonly WorkbookChangeSummary[]
  readonly checks: readonly WorkbookCheckResult[]
}

export interface WorkbookActionInspection {
  readonly name: string
  readonly description?: string
  readonly input?: WorkbookActionInputDescription
}

export interface WorkbookModelInspection {
  readonly name: string
  readonly description?: string
  readonly actions: readonly string[]
  readonly actionDetails: readonly WorkbookActionInspection[]
  readonly hasChecks: boolean
}

export type WorkbookActionPlanResult<Refs = unknown> =
  | {
      readonly status: 'planned'
      readonly plan: WorkbookActionPlan<Refs>
    }
  | {
      readonly status: 'failed'
      readonly modelName: string
      readonly actionName: string
      readonly input?: WorkbookActionInput
      readonly errors: readonly WorkbookRunError[]
      readonly checks: readonly WorkbookCheckResult[]
    }

export function defineModel<Refs, Actions extends WorkbookActionMap<Refs>>(
  config: WorkbookModelConfig<Refs, Actions>,
): WorkbookModel<Refs, Actions>
export function defineModel<Refs>(config: WorkbookModelConfig<Refs, WorkbookActionMap<Refs>>): WorkbookModel<Refs> {
  const name = normalizeRequiredName(config.name, 'Workbook model name')
  const description = normalizeOptionalDescription(config.description, `Workbook model ${name} description`)
  const actionNames = Object.keys(config.actions)
  if (actionNames.length === 0) {
    throw new Error(`Workbook model ${name} must define at least one action`)
  }
  actionNames.forEach((actionName) => {
    normalizeRequiredName(actionName, `Workbook model ${name} action name`)
  })
  const actions: WorkbookActionMap<Refs> = {}
  actionNames.forEach((actionName) => {
    actions[actionName] = normalizeActionDefinition(name, actionName, config.actions[actionName])
  })
  Object.freeze(actions)
  return Object.freeze({
    name,
    ...(description !== undefined ? { description } : {}),
    find: config.find,
    ...(config.checks !== undefined ? { checks: config.checks } : {}),
    actions,
  })
}

export function inspectModel<Refs, Actions extends WorkbookActionMap<Refs>>(model: WorkbookModel<Refs, Actions>): WorkbookModelInspection {
  const actions = Object.keys(model.actions).toSorted()
  const description = normalizeOptionalDescription(model.description, `Workbook model ${model.name} description`)
  return {
    name: model.name,
    ...(description !== undefined ? { description } : {}),
    actions,
    actionDetails: actions.map((actionName) => inspectAction(actionName, model.actions[actionName])),
    hasChecks: model.checks !== undefined,
  }
}

function normalizeRequiredName(value: string, label: string): string {
  const name = value.trim()
  if (name === '') {
    throw new Error(`${label} cannot be empty`)
  }
  if (name !== value) {
    throw new Error(`${label} must not have leading or trailing whitespace`)
  }
  return name
}

function normalizeOptionalDescription(value: unknown, label: string): string | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'string') {
    throw new Error(`${label} must be a string`)
  }
  const description = value.trim()
  if (description === '') {
    throw new Error(`${label} cannot be empty`)
  }
  return description
}

function isActionConfig<Refs>(definition: WorkbookActionDefinition<Refs> | undefined): definition is WorkbookActionConfig<Refs> {
  return typeof definition === 'object' && definition !== null
}

function ownPropertyValue(value: object, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function ownActionRun<Refs>(definition: object): WorkbookAction<Refs> | undefined {
  const run = ownPropertyValue(definition, 'run')
  if (typeof run !== 'function') {
    return undefined
  }
  return (context) => {
    Reflect.apply(run, undefined, [context])
  }
}

function normalizeActionDefinition<Refs>(
  modelName: string,
  actionName: string,
  definition: WorkbookActionDefinition<Refs> | undefined,
): WorkbookActionDefinition<Refs> {
  if (typeof definition === 'function') {
    return definition
  }
  if (!isActionConfig(definition)) {
    throw new Error(`Workbook model ${modelName} action ${actionName} must be a function or action object with run`)
  }
  const run = ownActionRun<Refs>(definition)
  if (run === undefined) {
    throw new Error(`Workbook model ${modelName} action ${actionName} must be a function or action object with run`)
  }
  const description = normalizeOptionalDescription(
    ownPropertyValue(definition, 'description'),
    `Workbook model ${modelName} action ${actionName} description`,
  )
  const inputValue = ownPropertyValue(definition, 'input')
  const input = inputValue === undefined ? undefined : normalizeWorkbookActionInputDescription(inputValue)
  return Object.freeze({
    ...(description !== undefined ? { description } : {}),
    ...(input !== undefined ? { input } : {}),
    run,
  })
}

function actionRunner<Refs>(definition: WorkbookActionDefinition<Refs>): WorkbookAction<Refs> {
  return typeof definition === 'function' ? definition : definition.run
}

function actionInputDescription<Refs>(definition: WorkbookActionDefinition<Refs>): WorkbookActionInputDescription | undefined {
  return typeof definition === 'function' ? undefined : definition.input
}

function inspectAction<Refs>(name: string, definition: WorkbookActionDefinition<Refs> | undefined): WorkbookActionInspection {
  if (definition === undefined || typeof definition === 'function') {
    return { name }
  }
  const description = normalizeOptionalDescription(ownPropertyValue(definition, 'description'), `Workbook action ${name} description`)
  const inputValue = ownPropertyValue(definition, 'input')
  const input = inputValue === undefined ? undefined : normalizeWorkbookActionInputDescription(inputValue)
  return {
    name,
    ...(description !== undefined ? { description } : {}),
    ...(input !== undefined ? { input } : {}),
  }
}

function concreteSingleCell(target: WorkbookRef): { sheetName: string; address: string } | null {
  if (target.kind !== 'range') {
    return null
  }
  const range = target.range
  return range.startAddress === range.endAddress ? { sheetName: range.sheetName, address: range.startAddress } : null
}

function commandMessage(command: WorkbookActionCommand): string {
  switch (command.kind) {
    case 'writeFormula':
      return `Write formula to ${command.target.label}`
    case 'writeValue':
      return `Write value to ${command.target.label}`
    case 'format':
      return `Format ${command.target.label}`
    case 'clear':
      return `Clear ${command.target.label}`
    case 'op':
      return (
        command.message ??
        (command.target === undefined
          ? `Add workbook op ${command.op.kind}`
          : `Add workbook op ${command.op.kind} for ${command.target.label}`)
      )
  }
}

function commandTarget(command: WorkbookActionCommand): WorkbookRef | undefined {
  return command.target
}

function cloneWorkbookOp(op: WorkbookOp): WorkbookOp {
  return structuredClone(op)
}

function freezeWorkbookOp(op: WorkbookOp): WorkbookOp {
  return Object.freeze(cloneWorkbookOp(op))
}

function freezeActionCommand(command: WorkbookActionCommand): WorkbookActionCommand {
  switch (command.kind) {
    case 'writeFormula':
      return Object.freeze({
        kind: 'writeFormula',
        target: command.target,
        formula: command.formula,
        inputs: Object.freeze([...command.inputs]),
        labels: Object.freeze(command.labels.map((label) => Object.freeze({ ...label }))),
      })
    case 'writeValue':
      return Object.freeze({
        kind: 'writeValue',
        target: command.target,
        value: command.value,
      })
    case 'format':
      return Object.freeze({
        kind: 'format',
        target: command.target,
        ...(command.style !== undefined ? { style: Object.freeze(structuredClone(command.style)) } : {}),
        ...(command.numberFormat !== undefined ? { numberFormat: command.numberFormat } : {}),
      })
    case 'clear':
      return Object.freeze({
        kind: 'clear',
        target: command.target,
      })
    case 'op':
      return Object.freeze({
        kind: 'op',
        op: freezeWorkbookOp(command.op),
        ...(command.target !== undefined ? { target: command.target } : {}),
        ...(command.message !== undefined ? { message: command.message } : {}),
      })
  }
}

function freezeCheckExpectation(check: WorkbookCheckResult): WorkbookCheckResult['expectation'] {
  if (check.expectation === undefined) {
    return undefined
  }
  if (check.expectation.kind === 'formulaEquals') {
    return Object.freeze({
      ...check.expectation,
      inputs: Object.freeze([...check.expectation.inputs]),
      labels: Object.freeze(check.expectation.labels.map((label) => Object.freeze({ ...label }))),
    })
  }
  return Object.freeze({ ...check.expectation })
}

function freezeCheckResult(check: WorkbookCheckResult): WorkbookCheckResult {
  const expectation = freezeCheckExpectation(check)
  const proof = check.proof === undefined ? undefined : normalizeWorkbookActionInput(check.proof)
  return Object.freeze({
    status: check.status,
    kind: check.kind,
    ...(check.target !== undefined ? { target: check.target } : {}),
    ...(check.refs !== undefined ? { refs: Object.freeze([...check.refs]) } : {}),
    message: check.message,
    ...(expectation !== undefined ? { expectation } : {}),
    ...(proof !== undefined ? { proof } : {}),
  })
}

function freezeChangeSummary(change: WorkbookChangeSummary): WorkbookChangeSummary {
  return Object.freeze({
    kind: change.kind,
    ...(change.target !== undefined ? { target: change.target } : {}),
    message: change.message,
  })
}

function isFreezableRefContainer(value: object): boolean {
  const prototype = Object.getPrototypeOf(value)
  return Array.isArray(value) || prototype === Object.prototype || prototype === null
}

function freezeRefs<Refs>(refs: Refs, seen = new WeakSet<object>()): Refs {
  if (typeof refs !== 'object' || refs === null) {
    return refs
  }
  if (seen.has(refs)) {
    return refs
  }
  seen.add(refs)

  Object.values(Object.getOwnPropertyDescriptors(refs)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeRefs(descriptor.value, seen)
    }
  })

  return isFreezableRefContainer(refs) ? (Object.freeze(refs) as Refs) : refs
}

function createCheckWorkbook(input: { readonly checks: WorkbookCheckResult[] }): WorkbookCheckWorkbook {
  return Object.freeze({
    ...createWorkbookFindApi(),
    check: createWorkbookCheckApi((entry) => input.checks.push(entry)),
  })
}

function createActionWorkbook(input: {
  readonly commands: WorkbookActionCommand[]
  readonly ops: WorkbookOp[]
  readonly checks: WorkbookCheckResult[]
}): WorkbookActionWorkbook {
  const checkWorkbook = createCheckWorkbook({ checks: input.checks })

  function pushCommand(command: WorkbookActionCommand): void {
    if (command.kind === 'op') {
      const commandOp = cloneWorkbookOp(command.op)
      const planOp = cloneWorkbookOp(command.op)
      input.commands.push({
        kind: 'op',
        op: commandOp,
        ...(command.target !== undefined ? { target: command.target } : {}),
        ...(command.message !== undefined ? { message: command.message } : {}),
      })
      input.ops.push(planOp)
      return
    }

    input.commands.push(command)

    const target = concreteSingleCell(command.target)
    if (target === null) {
      return
    }
    switch (command.kind) {
      case 'writeFormula':
        input.ops.push({
          kind: 'setCellFormula',
          sheetName: target.sheetName,
          address: target.address,
          formula: command.formula,
        })
        return
      case 'writeValue':
        input.ops.push({
          kind: 'setCellValue',
          sheetName: target.sheetName,
          address: target.address,
          value: command.value,
        })
        return
      case 'clear':
        input.ops.push({
          kind: 'clearCell',
          sheetName: target.sheetName,
          address: target.address,
        })
        return
      case 'format':
        if (command.numberFormat !== undefined) {
          input.ops.push({
            kind: 'setCellFormat',
            sheetName: target.sheetName,
            address: target.address,
            format: command.numberFormat,
          })
        }
        return
    }
  }

  const workbook: WorkbookActionWorkbook = {
    ...checkWorkbook,
    writeFormula(target: WorkbookRef, value: WorkbookFormulaOperand) {
      pushCommand({
        kind: 'writeFormula',
        target,
        formula: formula.source(value),
        inputs: formula.inputs(value),
        labels: formula.labels(value),
      })
    },
    writeValue(target: WorkbookRef, value: LiteralInput) {
      pushCommand({
        kind: 'writeValue',
        target,
        value,
      })
    },
    format(target: WorkbookRef, options: { readonly style?: CellStylePatch; readonly numberFormat?: string | null }) {
      pushCommand({
        kind: 'format',
        target,
        ...(options.style !== undefined ? { style: options.style } : {}),
        ...(options.numberFormat !== undefined ? { numberFormat: options.numberFormat } : {}),
      })
    },
    clear(target: WorkbookRef) {
      pushCommand({
        kind: 'clear',
        target,
      })
    },
    addOp(op: WorkbookOp, options: WorkbookAddOpOptions = {}) {
      if (!isWorkbookOp(op)) {
        throw new Error('Workbook op is not a valid WorkbookOp')
      }
      pushCommand({
        kind: 'op',
        op,
        ...(options.target !== undefined ? { target: options.target } : {}),
        ...(options.message !== undefined ? { message: options.message } : {}),
      })
    },
  }
  return Object.freeze(workbook)
}

function pushReturnedChecks(target: WorkbookCheckResult[], returned: readonly WorkbookCheckResult[] | undefined): void {
  returned?.forEach((entry) => {
    if (!target.includes(entry)) {
      target.push(entry)
    }
  })
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function inputProperty(input: WorkbookActionInput | undefined): { readonly input: WorkbookActionInput } | {} {
  return input === undefined ? {} : { input }
}

function failedPlan<Refs>(
  modelName: string,
  actionName: string,
  code: WorkbookRunErrorCode,
  message: string,
  checks: readonly WorkbookCheckResult[] = [],
  input?: WorkbookActionInput,
): WorkbookActionPlanResult<Refs> {
  return {
    status: 'failed',
    modelName,
    actionName,
    ...inputProperty(input),
    checks,
    errors: [{ code, message }],
  }
}

function actionNotFound(modelName: string, actionName: string): WorkbookRunError {
  return {
    code: 'action_not_found',
    message: `Workbook model ${modelName} does not define action ${actionName}`,
  }
}

function actionInputError(issue: WorkbookActionInputIssue): WorkbookRunError {
  return {
    code: 'invalid_action_input',
    message: issue.message,
    path: issue.path,
    issueCode: issue.code,
  }
}

function invalidActionInputError(error: unknown): WorkbookRunError {
  return {
    code: 'invalid_action_input',
    message: errorMessage(error),
    path: error instanceof WorkbookActionInputError ? error.path : 'input',
    issueCode: 'invalid_action_input',
  }
}

function createActionPlan<Refs>(
  modelName: string,
  actionName: string,
  input: WorkbookActionInput | undefined,
  refs: Refs,
  commands: readonly WorkbookActionCommand[],
  ops: readonly WorkbookOp[],
  checks: readonly WorkbookCheckResult[],
): WorkbookActionPlan<Refs> {
  const plannedCommands = Object.freeze(commands.map(freezeActionCommand))
  const plannedOps = Object.freeze(ops.map(freezeWorkbookOp))
  const plannedChecks = Object.freeze(checks.map(freezeCheckResult))
  const plannedRefs = freezeRefs(refs)
  const plannedChanged = Object.freeze(
    plannedCommands.map((command) => {
      const target = commandTarget(command)
      return freezeChangeSummary({
        kind: command.kind,
        ...(target !== undefined ? { target } : {}),
        message: commandMessage(command),
      })
    }),
  )

  return Object.freeze({
    modelName,
    actionName,
    ...inputProperty(input),
    refs: plannedRefs,
    refsUsed: Object.freeze([...collectWorkbookRefs(plannedRefs)]),
    commands: plannedCommands,
    ops: plannedOps,
    changed: plannedChanged,
    checks: plannedChecks,
  })
}

export function planWorkbookAction<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
  actionName: string,
  input?: WorkbookActionInput,
): WorkbookActionPlanResult<Refs> {
  let normalizedInput: WorkbookActionInput | undefined
  try {
    normalizedInput = normalizeOptionalWorkbookActionInput(input)
  } catch (error) {
    return {
      status: 'failed',
      modelName: model.name,
      actionName,
      checks: [],
      errors: [invalidActionInputError(error)],
    }
  }

  const actionDefinition = model.actions[actionName]
  if (actionDefinition === undefined) {
    return {
      status: 'failed',
      modelName: model.name,
      actionName,
      ...inputProperty(normalizedInput),
      checks: [],
      errors: [actionNotFound(model.name, actionName)],
    }
  }

  const inputDescription = actionInputDescription(actionDefinition)
  if (inputDescription !== undefined) {
    const inputCheck = checkInput(inputDescription, normalizedInput)
    if (inputCheck.status === 'invalid') {
      return {
        status: 'failed',
        modelName: model.name,
        actionName,
        ...inputProperty(normalizedInput),
        checks: [],
        errors: inputCheck.issues.map(actionInputError),
      }
    }
    normalizedInput = inputCheck.input
  }

  const action = actionRunner(actionDefinition)

  const commands: WorkbookActionCommand[] = []
  const ops: WorkbookOp[] = []
  const checks: WorkbookCheckResult[] = []
  const findWorkbook = Object.freeze(createWorkbookFindApi())

  let refs: Refs
  try {
    refs = model.find(findWorkbook)
  } catch (error) {
    return failedPlan<Refs>(model.name, actionName, 'find_failed', errorMessage(error), checks, normalizedInput)
  }

  const checkContext: WorkbookCheckContext<Refs> = { refs, workbook: createCheckWorkbook({ checks }), ...inputProperty(normalizedInput) }
  try {
    pushReturnedChecks(checks, model.checks?.(checkContext))
  } catch (error) {
    return failedPlan<Refs>(model.name, actionName, 'checks_failed', errorMessage(error), checks, normalizedInput)
  }

  const actionContext: WorkbookActionContext<Refs> = {
    refs,
    workbook: createActionWorkbook({ commands, ops, checks }),
    ...inputProperty(normalizedInput),
  }
  try {
    action(actionContext)
  } catch (error) {
    return failedPlan<Refs>(model.name, actionName, 'action_failed', errorMessage(error), checks, normalizedInput)
  }

  return {
    status: 'planned',
    plan: createActionPlan(model.name, actionName, normalizedInput, refs, commands, ops, checks),
  }
}

export function buildWorkbookActionPlan<Refs, Actions extends WorkbookActionMap<Refs>, ActionName extends keyof Actions & string>(
  model: WorkbookModel<Refs, Actions>,
  actionName: ActionName,
  input?: WorkbookActionInput,
): WorkbookActionPlan<Refs> {
  const result = planWorkbookAction(model, actionName, input)
  if (result.status === 'failed') {
    const [error] = result.errors
    throw new Error(error?.message ?? `Workbook model ${model.name} failed to plan action ${actionName}`)
  }
  return result.plan
}
