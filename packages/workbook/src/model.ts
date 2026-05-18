import type { CellStylePatch, LiteralInput } from '@bilig/protocol'
import { formula, type WorkbookFormulaOperand } from './formula.js'
import { collectWorkbookRefs, createWorkbookFindApi, type WorkbookFindApi, type WorkbookRef } from './find.js'
import { createWorkbookCheckApi, type WorkbookCheckApi } from './check.js'
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
      readonly numberFormat?: string
    }
  | {
      readonly kind: 'clear'
      readonly target: WorkbookRef
    }

export interface WorkbookModelWorkbook extends WorkbookFindApi {
  readonly check: WorkbookCheckApi
  readonly writeFormula: (target: WorkbookRef, value: WorkbookFormulaOperand) => void
  readonly writeValue: (target: WorkbookRef, value: LiteralInput) => void
  readonly format: (target: WorkbookRef, options: { readonly style?: CellStylePatch; readonly numberFormat?: string }) => void
  readonly clear: (target: WorkbookRef) => void
}

export interface WorkbookActionContext<Refs> {
  readonly refs: Refs
  readonly workbook: WorkbookModelWorkbook
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
  }
}

function createModelWorkbook(input: {
  readonly commands: WorkbookActionCommand[]
  readonly ops: WorkbookOp[]
  readonly checks: WorkbookCheckResult[]
}): WorkbookModelWorkbook {
  const find = createWorkbookFindApi()
  const check = createWorkbookCheckApi((entry) => input.checks.push(entry))

  function pushCommand(command: WorkbookActionCommand): void {
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

function failedPlan<Refs>(code: string, message: string, checks: readonly WorkbookCheckResult[] = []): WorkbookActionPlanResult<Refs> {
  return {
    status: 'failed',
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
  refs: Refs,
  commands: readonly WorkbookActionCommand[],
  ops: readonly WorkbookOp[],
  checks: readonly WorkbookCheckResult[],
): WorkbookActionPlan<Refs> {
  return {
    modelName,
    actionName,
    refs,
    refsUsed: collectWorkbookRefs(refs),
    commands,
    ops,
    changed: commands.map((command) => ({
      kind: command.kind,
      target: command.target,
      message: commandMessage(command),
    })),
    checks,
  }
}

export function planWorkbookAction<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
  actionName: string,
): WorkbookActionPlanResult<Refs> {
  const action: WorkbookAction<Refs> | undefined = model.actions[actionName]
  if (action === undefined) {
    return {
      status: 'failed',
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
    return failedPlan<Refs>('find_failed', errorMessage(error), checks)
  }

  const context: WorkbookActionContext<Refs> = { refs, workbook }
  try {
    pushReturnedChecks(checks, model.checks?.(context))
  } catch (error) {
    return failedPlan<Refs>('checks_failed', errorMessage(error), checks)
  }

  try {
    action(context)
  } catch (error) {
    return failedPlan<Refs>('action_failed', errorMessage(error), checks)
  }

  return {
    status: 'planned',
    plan: createActionPlan(model.name, actionName, refs, commands, ops, checks),
  }
}

export function buildWorkbookActionPlan<Refs, Actions extends WorkbookActionMap<Refs>, ActionName extends keyof Actions & string>(
  model: WorkbookModel<Refs, Actions>,
  actionName: ActionName,
): WorkbookActionPlan<Refs> {
  const result = planWorkbookAction(model, actionName)
  if (result.status === 'failed') {
    const [error] = result.errors
    throw new Error(error?.message ?? `Workbook model ${model.name} failed to plan action ${actionName}`)
  }
  return result.plan
}
