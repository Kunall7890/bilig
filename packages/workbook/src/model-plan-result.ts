import { WorkbookActionInputError, type WorkbookActionInput, type WorkbookActionInputIssue } from './input.js'
import type { WorkbookActionPlan, WorkbookActionPlanResult, WorkbookModelInspection } from './model.js'
import type { WorkbookCheckResult, WorkbookRunError, WorkbookRunErrorCode } from './result.js'

export function freezeModelInspection(inspection: WorkbookModelInspection): WorkbookModelInspection {
  return freezeData(inspection, new WeakSet())
}

export function plannedActionPlanResult<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookActionPlanResult<Refs> {
  return freezePlanResult({
    status: 'planned',
    plan,
  })
}

export function failedPlan<Refs>(
  modelName: string,
  actionName: string,
  code: WorkbookRunErrorCode,
  message: string,
  checks: readonly WorkbookCheckResult[] = [],
  input?: WorkbookActionInput,
): WorkbookActionPlanResult<Refs> {
  return freezePlanResult({
    status: 'failed',
    modelName,
    actionName,
    ...inputProperty(input),
    checks,
    errors: [Object.freeze({ code, message })],
  })
}

export function failedInvalidModelPlan<Refs>(
  actionName: string,
  input: WorkbookActionInput | undefined,
  error: unknown,
  modelName = 'unknown-model',
): WorkbookActionPlanResult<Refs> {
  return freezePlanResult({
    status: 'failed',
    modelName,
    actionName,
    ...inputProperty(input),
    checks: [],
    errors: [invalidModelError(error)],
  })
}

export function failedActionInputPlan<Refs>(modelName: string, actionName: string, error: unknown): WorkbookActionPlanResult<Refs> {
  return freezePlanResult({
    status: 'failed',
    modelName,
    actionName,
    checks: [],
    errors: [invalidActionInputError(error)],
  })
}

export function failedActionInputIssuesPlan<Refs>(
  modelName: string,
  actionName: string,
  issues: readonly WorkbookActionInputIssue[],
  input?: WorkbookActionInput,
): WorkbookActionPlanResult<Refs> {
  return freezePlanResult({
    status: 'failed',
    modelName,
    actionName,
    ...inputProperty(input),
    checks: [],
    errors: issues.map(actionInputError),
  })
}

export function failedActionNotFoundPlan<Refs>(
  modelName: string,
  actionName: string,
  input?: WorkbookActionInput,
): WorkbookActionPlanResult<Refs> {
  return freezePlanResult({
    status: 'failed',
    modelName,
    actionName,
    ...inputProperty(input),
    checks: [],
    errors: [actionNotFound(modelName, actionName)],
  })
}

export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export function inputProperty(input: WorkbookActionInput | undefined): { readonly input: WorkbookActionInput } | {} {
  return input === undefined ? {} : { input }
}

function freezePlanResult<Refs>(result: WorkbookActionPlanResult<Refs>): WorkbookActionPlanResult<Refs> {
  return freezeData(result, new WeakSet())
}

function freezeData<T>(value: T, seen: WeakSet<object>): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeData(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}

function actionNotFound(modelName: string, actionName: string): WorkbookRunError {
  return Object.freeze({
    code: 'action_not_found',
    message: `Workbook model ${modelName} does not define action ${actionName}`,
  })
}

function actionInputError(issue: WorkbookActionInputIssue): WorkbookRunError {
  return Object.freeze({
    code: 'invalid_action_input',
    message: issue.message,
    path: issue.path,
    issueCode: issue.code,
  })
}

function invalidActionInputError(error: unknown): WorkbookRunError {
  return Object.freeze({
    code: 'invalid_action_input',
    message: errorMessage(error),
    path: error instanceof WorkbookActionInputError ? error.path : 'input',
    issueCode: 'invalid_action_input',
  })
}

function invalidModelError(error: unknown): WorkbookRunError {
  return Object.freeze({
    code: 'invalid_model',
    message: errorMessage(error),
  })
}
