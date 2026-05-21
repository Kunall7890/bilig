import type { CellStylePatch, LiteralInput } from '@bilig/protocol'
import { formula, type WorkbookFormulaOperand } from './formula.js'
import { collectWorkbookRefs, createWorkbookFindApi, type WorkbookFindApi, type WorkbookRef } from './find.js'
import { createWorkbookCheckApi, type WorkbookCheckApi } from './check.js'
import { normalizeOptionalWorkbookActionInput, type WorkbookActionInput } from './input.js'
import { isWorkbookOp } from './guards.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookChangeSummary, WorkbookCheckResult, WorkbookRunError } from './result.js'

export type WorkbookActionCommand =
  | {
      readonly kind: 'writeFormula'
      readonly target: WorkbookRef
      readonly formula: string
      readonly inputs: readonly WorkbookRef[]
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

export interface WorkbookModelWorkbook extends WorkbookFindApi {
  readonly check: WorkbookCheckApi
  readonly writeFormula: (target: WorkbookRef, value: WorkbookFormulaOperand) => void
  readonly writeValue: (target: WorkbookRef, value: LiteralInput) => void
  readonly format: (target: WorkbookRef, options: { readonly style?: CellStylePatch; readonly numberFormat?: string | null }) => void
  readonly clear: (target: WorkbookRef) => void
  readonly addOp: (op: WorkbookOp, options?: WorkbookAddOpOptions) => void
}

export interface WorkbookActionContext<Refs> {
  readonly refs: Refs
  readonly workbook: WorkbookModelWorkbook
  readonly input?: WorkbookActionInput
}

export type WorkbookAction<Refs> = (context: WorkbookActionContext<Refs>) => void

export type WorkbookActionMap<Refs> = Record<string, WorkbookAction<Refs>>

export interface WorkbookModelConfig<Refs, Actions extends WorkbookActionMap<Refs>> {
  readonly name: string
  readonly find: (workbook: WorkbookModelWorkbook) => Refs
  readonly checks?: (context: WorkbookActionContext<Refs>) => readonly WorkbookCheckResult[]
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

export interface WorkbookModelInspection {
  readonly name: string
  readonly actions: readonly string[]
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
): WorkbookModel<Refs, Actions> {
  if (config.name.trim() === '') {
    throw new Error('Workbook model name cannot be empty')
  }
  const actionNames = Object.keys(config.actions)
  if (actionNames.length === 0) {
    throw new Error(`Workbook model ${config.name} must define at least one action`)
  }
  const emptyActionName = actionNames.find((name) => name.trim() === '')
  if (emptyActionName !== undefined) {
    throw new Error(`Workbook model ${config.name} has an empty action name`)
  }
  return config
}

export function inspectModel<Refs, Actions extends WorkbookActionMap<Refs>>(model: WorkbookModel<Refs, Actions>): WorkbookModelInspection {
  return {
    name: model.name,
    actions: Object.keys(model.actions).toSorted(),
    hasChecks: model.checks !== undefined,
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

function createModelWorkbook(input: {
  readonly commands: WorkbookActionCommand[]
  readonly ops: WorkbookOp[]
  readonly checks: WorkbookCheckResult[]
}): WorkbookModelWorkbook {
  const find = createWorkbookFindApi()
  const check = createWorkbookCheckApi((entry) => input.checks.push(entry))

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

  return {
    ...find,
    check,
    writeFormula(target, value) {
      pushCommand({
        kind: 'writeFormula',
        target,
        formula: formula.source(value),
        inputs: formula.inputs(value),
      })
    },
    writeValue(target, value) {
      pushCommand({
        kind: 'writeValue',
        target,
        value,
      })
    },
    format(target, options) {
      pushCommand({
        kind: 'format',
        target,
        ...(options.style !== undefined ? { style: options.style } : {}),
        ...(options.numberFormat !== undefined ? { numberFormat: options.numberFormat } : {}),
      })
    },
    clear(target) {
      pushCommand({
        kind: 'clear',
        target,
      })
    },
    addOp(op, options = {}) {
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
  code: string,
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

function createActionPlan<Refs>(
  modelName: string,
  actionName: string,
  input: WorkbookActionInput | undefined,
  refs: Refs,
  commands: readonly WorkbookActionCommand[],
  ops: readonly WorkbookOp[],
  checks: readonly WorkbookCheckResult[],
): WorkbookActionPlan<Refs> {
  return {
    modelName,
    actionName,
    ...inputProperty(input),
    refs,
    refsUsed: collectWorkbookRefs(refs),
    commands,
    ops,
    changed: commands.map((command) => {
      const target = commandTarget(command)
      return {
        kind: command.kind,
        ...(target !== undefined ? { target } : {}),
        message: commandMessage(command),
      }
    }),
    checks,
  }
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
    return failedPlan<Refs>(model.name, actionName, 'invalid_action_input', errorMessage(error))
  }

  const action: WorkbookAction<Refs> | undefined = model.actions[actionName]
  if (action === undefined) {
    return {
      status: 'failed',
      modelName: model.name,
      actionName,
      ...inputProperty(normalizedInput),
      checks: [],
      errors: [actionNotFound(model.name, actionName)],
    }
  }

  const commands: WorkbookActionCommand[] = []
  const ops: WorkbookOp[] = []
  const checks: WorkbookCheckResult[] = []
  const workbook = createModelWorkbook({ commands, ops, checks })

  let refs: Refs
  try {
    refs = model.find(workbook)
  } catch (error) {
    return failedPlan<Refs>(model.name, actionName, 'find_failed', errorMessage(error), checks, normalizedInput)
  }

  const context: WorkbookActionContext<Refs> = { refs, workbook, ...inputProperty(normalizedInput) }
  try {
    pushReturnedChecks(checks, model.checks?.(context))
  } catch (error) {
    return failedPlan<Refs>(model.name, actionName, 'checks_failed', errorMessage(error), checks, normalizedInput)
  }

  try {
    action(context)
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
