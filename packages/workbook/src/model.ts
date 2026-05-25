import { isLiteralInput, type CellStylePatch, type LiteralInput } from '@bilig/protocol'
import { formula, type WorkbookFormulaLabel, type WorkbookFormulaOperand } from './formula.js'
import { collectWorkbookRefs, createWorkbookFindApi, isWorkbookRef, type WorkbookFindApi, type WorkbookRef } from './find.js'
import { createWorkbookCheckApi, type WorkbookCheckApi } from './check.js'
import {
  checkInput,
  normalizeOptionalWorkbookActionInput,
  normalizeWorkbookActionInput,
  normalizeWorkbookActionInputDescription,
  type WorkbookActionInput,
  type WorkbookActionInputDescription,
} from './input.js'
import { isObject, isObjectRecord, optionalDataProperty, requiredDataProperty, type OptionalDataValue } from './data-properties.js'
import {
  errorMessage,
  failedInvalidActionNamePlan,
  failedActionInputIssuesPlan,
  failedActionInputPlan,
  failedActionNotFoundPlan,
  failedInvalidModelPlan,
  failedPlan,
  freezeModelInspection,
  inputProperty,
  plannedActionPlanResult,
} from './model-plan-result.js'
import {
  normalizeWorkbookActionFormatOptions,
  normalizeWorkbookActionLiteralInput,
  normalizeWorkbookActionOp,
  normalizeWorkbookActionTarget,
  normalizeWorkbookAddOpOptions,
} from './model-action-validation.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookChangeSummary, WorkbookCheckExpectation, WorkbookCheckResult, WorkbookRunError } from './result.js'

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
  if (!isObjectRecord(config)) {
    throw new Error('Workbook model config must be an object')
  }
  const name = normalizeRequiredName(requiredModelConfigValue(config, 'name'), 'Workbook model name')
  const descriptionValue = optionalModelConfigValue(config, 'description')
  const description = normalizeOptionalDescription(
    descriptionValue.status === 'present' ? descriptionValue.value : undefined,
    `Workbook model ${name} description`,
  )
  const find = requiredModelConfigValue(config, 'find')
  if (!isFindFunction<Refs>(find)) {
    throw new Error(`Workbook model ${name} find must be a function`)
  }
  const checksValue = optionalModelConfigValue(config, 'checks')
  let checks: ((context: WorkbookCheckContext<Refs>) => readonly WorkbookCheckResult[]) | undefined
  if (checksValue.status === 'present') {
    if (!isChecksFunction<Refs>(checksValue.value)) {
      throw new Error(`Workbook model ${name} checks must be a function`)
    }
    checks = checksValue.value
  }
  const actionMap = requiredModelConfigValue(config, 'actions')
  if (!isObject(actionMap) || Array.isArray(actionMap)) {
    throw new Error(`Workbook model ${name} actions must be an object`)
  }
  const actionNames = Object.keys(actionMap)
  if (actionNames.length === 0) {
    throw new Error(`Workbook model ${name} must define at least one action`)
  }
  actionNames.forEach((actionName) => {
    normalizeRequiredName(actionName, `Workbook model ${name} action name`)
  })
  const actions: WorkbookActionMap<Refs> = {}
  Object.setPrototypeOf(actions, null)
  actionNames.forEach((actionName) => {
    actions[actionName] = normalizeActionDefinition(
      name,
      actionName,
      requiredDataProperty(actionMap, actionName, `Workbook model ${name} action ${actionName}`),
    )
  })
  Object.freeze(actions)
  return Object.freeze({
    name,
    ...(description !== undefined ? { description } : {}),
    find: (workbook: WorkbookFindWorkbook) => Reflect.apply(find, undefined, [workbook]),
    ...(checks !== undefined
      ? {
          checks: (context: WorkbookCheckContext<Refs>) => Reflect.apply(checks, undefined, [context]),
        }
      : {}),
    actions,
  })
}

export function inspectModel(model: unknown): WorkbookModelInspection {
  if (!isObjectRecord(model)) {
    throw new Error('Workbook model must be an object')
  }
  const name = normalizeRequiredName(requiredDataProperty(model, 'name', 'Workbook model name'), 'Workbook model name')
  const actionMap = requiredDataProperty(model, 'actions', `Workbook model ${name} actions`)
  if (!isObject(actionMap) || Array.isArray(actionMap)) {
    throw new Error(`Workbook model ${name} actions must be an object`)
  }
  const actions = Object.keys(actionMap).toSorted()
  const descriptionValue = optionalDataProperty(model, 'description', `Workbook model ${name} description`)
  const checksValue = optionalDataProperty(model, 'checks', `Workbook model ${name} checks`)
  const description = normalizeOptionalDescription(
    descriptionValue.status === 'present' ? descriptionValue.value : undefined,
    `Workbook model ${name} description`,
  )
  return freezeModelInspection({
    name,
    ...(description !== undefined ? { description } : {}),
    actions,
    actionDetails: actions.map((actionName) =>
      inspectAction(actionName, requiredDataProperty(actionMap, actionName, `Workbook model ${name} action ${actionName}`)),
    ),
    hasChecks: checksValue.status === 'present' && checksValue.value !== undefined,
  })
}

function normalizeRequiredName(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`${label} must be a string`)
  }
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

function isActionConfig<Refs>(definition: unknown): definition is WorkbookActionConfig<Refs> {
  return typeof definition === 'object' && definition !== null && !Array.isArray(definition)
}

function isFindFunction<Refs>(value: unknown): value is (workbook: WorkbookFindWorkbook) => Refs {
  return typeof value === 'function'
}

function isChecksFunction<Refs>(value: unknown): value is (context: WorkbookCheckContext<Refs>) => readonly WorkbookCheckResult[] {
  return typeof value === 'function'
}

function isWorkbookActionFunction<Refs>(value: unknown): value is WorkbookAction<Refs> {
  return typeof value === 'function'
}

function optionalDataValue(value: object, key: string, path: string): OptionalDataValue {
  return optionalDataProperty(value, key, `Workbook check at ${path}`)
}

function requiredDataValue(value: object, key: string, path: string): unknown {
  const property = optionalDataValue(value, key, path)
  if (property.status === 'missing') {
    throw new Error(`Workbook check at ${path} is missing`)
  }
  return property.value
}

function dataArrayEntries(value: unknown, path: string): readonly (readonly [unknown, string])[] {
  if (!Array.isArray(value)) {
    throw new Error(`Workbook check at ${path} must be an array`)
  }

  const entries: (readonly [unknown, string])[] = []
  for (let index = 0; index < value.length; index += 1) {
    const entryPath = `${path}[${index}]`
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      throw new Error(`Workbook check at ${entryPath} must be a data property`)
    }
    entries.push([descriptor.value, entryPath])
  }
  return entries
}

function checkRef(value: unknown, path: string): WorkbookRef {
  if (!isWorkbookRef(value)) {
    throw new Error(`Workbook check at ${path} must be a workbook ref`)
  }
  return value
}

function checkRefArray(value: unknown, path: string): readonly WorkbookRef[] {
  return Object.freeze(dataArrayEntries(value, path).map(([entry, entryPath]) => checkRef(entry, entryPath)))
}

function checkFormulaLabel(value: unknown, path: string): WorkbookFormulaLabel {
  if (!isObjectRecord(value)) {
    throw new Error(`Workbook check at ${path} must be a formula label`)
  }
  const name = requiredDataValue(value, 'name', `${path}.name`)
  if (typeof name !== 'string') {
    throw new Error(`Workbook check at ${path}.name must be a string`)
  }
  const ref = checkRef(requiredDataValue(value, 'ref', `${path}.ref`), `${path}.ref`)
  return Object.freeze({
    name,
    ref,
  })
}

function checkFormulaLabels(value: unknown, path: string): readonly WorkbookFormulaLabel[] {
  return Object.freeze(dataArrayEntries(value, path).map(([entry, entryPath]) => checkFormulaLabel(entry, entryPath)))
}

function cloneCheckExpectation(value: unknown, path: string): WorkbookCheckExpectation {
  if (!isObjectRecord(value)) {
    throw new Error(`Workbook check at ${path} must be an expectation`)
  }
  const kind = requiredDataValue(value, 'kind', `${path}.kind`)
  if (kind === 'valueEquals') {
    const expected = requiredDataValue(value, 'value', `${path}.value`)
    if (!isLiteralInput(expected)) {
      throw new Error(`Workbook check at ${path}.value must be a finite JSON literal`)
    }
    return Object.freeze({
      kind: 'valueEquals',
      value: expected,
    })
  }
  if (kind === 'formulaEquals') {
    const expectedFormula = requiredDataValue(value, 'formula', `${path}.formula`)
    if (typeof expectedFormula !== 'string') {
      throw new Error(`Workbook check at ${path}.formula must be a string`)
    }
    return Object.freeze({
      kind: 'formulaEquals',
      formula: expectedFormula,
      inputs: checkRefArray(requiredDataValue(value, 'inputs', `${path}.inputs`), `${path}.inputs`),
      labels: checkFormulaLabels(requiredDataValue(value, 'labels', `${path}.labels`), `${path}.labels`),
    })
  }
  throw new Error(`Workbook check at ${path}.kind is invalid`)
}

function cloneCheckResult(check: unknown, path: string): WorkbookCheckResult {
  if (!isObjectRecord(check)) {
    throw new Error(`Workbook check at ${path} must be an object`)
  }

  const status = requiredDataValue(check, 'status', `${path}.status`)
  if (status !== 'planned' && status !== 'passed' && status !== 'failed') {
    throw new Error(`Workbook check at ${path}.status is invalid`)
  }

  const kind = requiredDataValue(check, 'kind', `${path}.kind`)
  if (typeof kind !== 'string') {
    throw new Error(`Workbook check at ${path}.kind must be a string`)
  }

  const message = requiredDataValue(check, 'message', `${path}.message`)
  if (typeof message !== 'string') {
    throw new Error(`Workbook check at ${path}.message must be a string`)
  }

  const target = optionalDataValue(check, 'target', `${path}.target`)
  const refs = optionalDataValue(check, 'refs', `${path}.refs`)
  const expectation = optionalDataValue(check, 'expectation', `${path}.expectation`)
  const proof = optionalDataValue(check, 'proof', `${path}.proof`)
  const proofValue = proof.status === 'present' ? normalizeWorkbookActionInput(proof.value) : undefined

  return Object.freeze({
    status,
    kind,
    ...(target.status === 'present' ? { target: checkRef(target.value, `${path}.target`) } : {}),
    ...(refs.status === 'present' ? { refs: checkRefArray(refs.value, `${path}.refs`) } : {}),
    message,
    ...(expectation.status === 'present' ? { expectation: cloneCheckExpectation(expectation.value, `${path}.expectation`) } : {}),
    ...(proofValue !== undefined ? { proof: proofValue } : {}),
  })
}

function optionalModelConfigValue(value: object, key: string): OptionalDataValue {
  return optionalDataProperty(value, key, `Workbook model config ${key}`)
}

function requiredModelConfigValue(value: object, key: string): unknown {
  return requiredDataProperty(value, key, `Workbook model config ${key}`)
}

function actionConfigValue(modelName: string, actionName: string, definition: object, key: string): OptionalDataValue {
  return optionalDataProperty(definition, key, `Workbook model ${modelName} action ${actionName} ${key}`)
}

function ownActionRun<Refs>(modelName: string, actionName: string, definition: object): WorkbookAction<Refs> | undefined {
  const runProperty = actionConfigValue(modelName, actionName, definition, 'run')
  if (runProperty.status === 'missing') {
    return undefined
  }
  const run = runProperty.value
  if (typeof run !== 'function') {
    return undefined
  }
  return (context) => {
    Reflect.apply(run, undefined, [context])
  }
}

function normalizeActionDefinition<Refs>(modelName: string, actionName: string, definition: unknown): WorkbookActionDefinition<Refs> {
  if (isWorkbookActionFunction<Refs>(definition)) {
    return definition
  }
  if (!isActionConfig(definition)) {
    throw new Error(`Workbook model ${modelName} action ${actionName} must be a function or action object with run`)
  }
  const run = ownActionRun<Refs>(modelName, actionName, definition)
  if (run === undefined) {
    throw new Error(`Workbook model ${modelName} action ${actionName} must be a function or action object with run`)
  }
  const descriptionValue = actionConfigValue(modelName, actionName, definition, 'description')
  const description = normalizeOptionalDescription(
    descriptionValue.status === 'present' ? descriptionValue.value : undefined,
    `Workbook model ${modelName} action ${actionName} description`,
  )
  const inputValue = actionConfigValue(modelName, actionName, definition, 'input')
  const input = inputValue.status === 'present' ? normalizeWorkbookActionInputDescription(inputValue.value) : undefined
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

function inspectAction(name: string, definition: unknown): WorkbookActionInspection {
  if (definition === undefined || typeof definition === 'function') {
    return { name }
  }
  if (!isActionConfig(definition)) {
    throw new Error(`Workbook action ${name} must be a function or action object`)
  }
  const descriptionValue = optionalDataProperty(definition, 'description', `Workbook action ${name} description`)
  const description = normalizeOptionalDescription(
    descriptionValue.status === 'present' ? descriptionValue.value : undefined,
    `Workbook action ${name} description`,
  )
  const inputValue = optionalDataProperty(definition, 'input', `Workbook action ${name} input`)
  const input = inputValue.status === 'present' ? normalizeWorkbookActionInputDescription(inputValue.value) : undefined
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
        op: normalizeWorkbookActionOp(command.op),
        ...(command.target !== undefined ? { target: command.target } : {}),
        ...(command.message !== undefined ? { message: command.message } : {}),
      })
  }
}

function freezeCheckResult(check: WorkbookCheckResult): WorkbookCheckResult {
  return cloneCheckResult(check, 'check')
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
      const commandOp = normalizeWorkbookActionOp(command.op)
      const planOp = normalizeWorkbookActionOp(command.op)
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
      const normalizedTarget = normalizeWorkbookActionTarget('writeFormula', target)
      pushCommand({
        kind: 'writeFormula',
        target: normalizedTarget,
        formula: formula.source(value),
        inputs: formula.inputs(value),
        labels: formula.labels(value),
      })
    },
    writeValue(target: WorkbookRef, value: LiteralInput) {
      const normalizedTarget = normalizeWorkbookActionTarget('writeValue', target)
      const normalizedValue = normalizeWorkbookActionLiteralInput('writeValue', value)
      pushCommand({
        kind: 'writeValue',
        target: normalizedTarget,
        value: normalizedValue,
      })
    },
    format(target: WorkbookRef, options: { readonly style?: CellStylePatch; readonly numberFormat?: string | null }) {
      const normalizedTarget = normalizeWorkbookActionTarget('format', target)
      const normalizedOptions = normalizeWorkbookActionFormatOptions(options)
      pushCommand({
        kind: 'format',
        target: normalizedTarget,
        ...(normalizedOptions.style !== undefined ? { style: normalizedOptions.style } : {}),
        ...(normalizedOptions.numberFormat !== undefined ? { numberFormat: normalizedOptions.numberFormat } : {}),
      })
    },
    clear(target: WorkbookRef) {
      const normalizedTarget = normalizeWorkbookActionTarget('clear', target)
      pushCommand({
        kind: 'clear',
        target: normalizedTarget,
      })
    },
    addOp(op: WorkbookOp, options: WorkbookAddOpOptions = {}) {
      const normalizedOp = normalizeWorkbookActionOp(op)
      const normalizedOptions = normalizeWorkbookAddOpOptions(options)
      pushCommand({
        kind: 'op',
        op: normalizedOp,
        ...(normalizedOptions.target !== undefined ? { target: normalizedOptions.target } : {}),
        ...(normalizedOptions.message !== undefined ? { message: normalizedOptions.message } : {}),
      })
    },
  }
  return Object.freeze(workbook)
}

function pushReturnedChecks(target: WorkbookCheckResult[], returned: readonly WorkbookCheckResult[] | undefined): void {
  if (returned === undefined) {
    return
  }
  dataArrayEntries(returned, 'checks').forEach(([entry, path]) => {
    if (!target.some((existing) => existing === entry)) {
      target.push(cloneCheckResult(entry, path))
    }
  })
}

interface WorkbookPlanningModelData<Refs> {
  readonly name: string
  readonly find: (workbook: WorkbookFindWorkbook) => Refs
  readonly checks?: (context: WorkbookCheckContext<Refs>) => readonly WorkbookCheckResult[]
  readonly actions: object
}

function readPlanningModelData<Refs>(model: unknown): WorkbookPlanningModelData<Refs> {
  if (!isObjectRecord(model)) {
    throw new Error('Workbook model must be an object')
  }
  const name = normalizeRequiredName(requiredDataProperty(model, 'name', 'Workbook model name'), 'Workbook model name')
  const find = requiredDataProperty(model, 'find', `Workbook model ${name} find`)
  if (!isFindFunction<Refs>(find)) {
    throw new Error(`Workbook model ${name} find must be a function`)
  }
  const checksValue = optionalDataProperty(model, 'checks', `Workbook model ${name} checks`)
  let checks: ((context: WorkbookCheckContext<Refs>) => readonly WorkbookCheckResult[]) | undefined
  if (checksValue.status === 'present' && checksValue.value !== undefined) {
    if (!isChecksFunction<Refs>(checksValue.value)) {
      throw new Error(`Workbook model ${name} checks must be a function`)
    }
    checks = checksValue.value
  }
  const actions = requiredDataProperty(model, 'actions', `Workbook model ${name} actions`)
  if (!isObject(actions) || Array.isArray(actions)) {
    throw new Error(`Workbook model ${name} actions must be an object`)
  }
  return {
    name,
    find,
    ...(checks !== undefined ? { checks } : {}),
    actions,
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
  const plannedOps = Object.freeze(ops.map(normalizeWorkbookActionOp))
  const plannedChecks = Object.freeze(checks.map(freezeCheckResult))
  const plannedRefs = freezeRefs(refs)
  const plannedChanged = Object.freeze(
    plannedCommands.map((command) => {
      const target = command.target
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
): WorkbookActionPlanResult<Refs>
export function planWorkbookAction(model: unknown, actionName: string, input?: WorkbookActionInput): WorkbookActionPlanResult
export function planWorkbookAction(model: unknown, actionName: unknown, input?: WorkbookActionInput): WorkbookActionPlanResult
export function planWorkbookAction<Refs>(model: unknown, actionName: unknown, input?: WorkbookActionInput): WorkbookActionPlanResult<Refs> {
  let plannedActionName: string
  try {
    plannedActionName = normalizeRequiredName(actionName, 'Workbook action name')
  } catch (error) {
    return failedInvalidActionNamePlan<Refs>(error)
  }

  let modelData: WorkbookPlanningModelData<Refs>
  try {
    modelData = readPlanningModelData(model)
  } catch (error) {
    return failedInvalidModelPlan<Refs>(plannedActionName, undefined, error)
  }

  let normalizedInput: WorkbookActionInput | undefined
  try {
    normalizedInput = normalizeOptionalWorkbookActionInput(input)
  } catch (error) {
    return failedActionInputPlan(modelData.name, plannedActionName, error)
  }

  let actionValue: unknown
  try {
    actionValue = Object.hasOwn(modelData.actions, plannedActionName)
      ? requiredDataProperty(modelData.actions, plannedActionName, `Workbook model ${modelData.name} action ${plannedActionName}`)
      : undefined
  } catch (error) {
    return failedInvalidModelPlan<Refs>(plannedActionName, normalizedInput, error, modelData.name)
  }
  if (actionValue === undefined) {
    return failedActionNotFoundPlan(modelData.name, plannedActionName, normalizedInput)
  }

  let actionDefinition: WorkbookActionDefinition<Refs>
  try {
    actionDefinition = normalizeActionDefinition(modelData.name, plannedActionName, actionValue)
  } catch (error) {
    return failedInvalidModelPlan<Refs>(plannedActionName, normalizedInput, error, modelData.name)
  }

  const inputDescription = actionInputDescription(actionDefinition)
  if (inputDescription !== undefined) {
    const inputCheck = checkInput(inputDescription, normalizedInput)
    if (inputCheck.status === 'invalid') {
      return failedActionInputIssuesPlan(modelData.name, plannedActionName, inputCheck.issues, normalizedInput)
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
    refs = Reflect.apply(modelData.find, undefined, [findWorkbook])
  } catch (error) {
    return failedPlan<Refs>(modelData.name, plannedActionName, 'find_failed', errorMessage(error), checks, normalizedInput)
  }

  const checkContext: WorkbookCheckContext<Refs> = { refs, workbook: createCheckWorkbook({ checks }), ...inputProperty(normalizedInput) }
  try {
    pushReturnedChecks(checks, modelData.checks === undefined ? undefined : Reflect.apply(modelData.checks, undefined, [checkContext]))
  } catch (error) {
    return failedPlan<Refs>(modelData.name, plannedActionName, 'checks_failed', errorMessage(error), checks, normalizedInput)
  }

  const actionContext: WorkbookActionContext<Refs> = {
    refs,
    workbook: createActionWorkbook({ commands, ops, checks }),
    ...inputProperty(normalizedInput),
  }
  try {
    action(actionContext)
  } catch (error) {
    return failedPlan<Refs>(modelData.name, plannedActionName, 'action_failed', errorMessage(error), checks, normalizedInput)
  }

  return plannedActionPlanResult(createActionPlan(modelData.name, plannedActionName, normalizedInput, refs, commands, ops, checks))
}

export function buildWorkbookActionPlan<Refs, Actions extends WorkbookActionMap<Refs>, ActionName extends keyof Actions & string>(
  model: WorkbookModel<Refs, Actions>,
  actionName: ActionName,
  input?: WorkbookActionInput,
): WorkbookActionPlan<Refs> {
  const result = planWorkbookAction(model, actionName, input)
  if (result.status === 'failed') {
    const [error] = result.errors
    throw new Error(error?.message ?? `Workbook model ${result.modelName} failed to plan action ${actionName}`)
  }
  return result.plan
}
