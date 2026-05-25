import type { WorkbookRunResultDescription } from './describe.js'
import { describeRunResult } from './describe.js'
import type { WorkbookActionPlan } from './model.js'
import type { WorkbookExecutablePlan, WorkbookPlanData, WorkbookPlanDataRefs } from './plan-data.js'
import type { WorkbookRunError, WorkbookRunErrorCode, WorkbookRunResult } from './result.js'
import { checkRuntimeAdapter, type WorkbookRuntimeAdapterIssue } from './requirements.js'
import { runWorkbookPlan, type WorkbookRunAdapter, type WorkbookRunOptions } from './run.js'
import { checkWorkbookRunResultDescription } from './run-description.js'

export type WorkbookRunAdapterCheckIssueCode = WorkbookRunErrorCode | 'invalid_run_result_description'

export interface WorkbookRunAdapterCheckIssue {
  readonly code: WorkbookRunAdapterCheckIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookRunAdapterCheckOptions = Omit<WorkbookRunOptions, 'strict'>

export type WorkbookRunAdapterPassedResult = {
  readonly status: 'passed'
  readonly result: Extract<WorkbookRunResult, { readonly status: 'done' }>
  readonly description: Extract<WorkbookRunResultDescription, { readonly status: 'done' }>
  readonly issues: readonly []
}

export type WorkbookRunAdapterFailedResult = {
  readonly status: 'failed'
  readonly result?: WorkbookRunResult
  readonly description?: WorkbookRunResultDescription
  readonly errors: readonly WorkbookRunError[]
  readonly issues: readonly WorkbookRunAdapterCheckIssue[]
}

export type WorkbookRunAdapterCheckResult = WorkbookRunAdapterPassedResult | WorkbookRunAdapterFailedResult

function adapterCheckIssue(code: WorkbookRunAdapterCheckIssueCode, path: string, message: string): WorkbookRunAdapterCheckIssue {
  return Object.freeze({
    code,
    path,
    message,
  })
}

function failedAdapterCheck(args: {
  readonly result?: WorkbookRunResult
  readonly description?: WorkbookRunResultDescription
  readonly errors?: readonly WorkbookRunError[]
  readonly issues: readonly WorkbookRunAdapterCheckIssue[]
}): WorkbookRunAdapterFailedResult {
  return Object.freeze({
    status: 'failed',
    ...(args.result !== undefined ? { result: args.result } : {}),
    ...(args.description !== undefined ? { description: args.description } : {}),
    errors: Object.freeze([...(args.errors ?? [])]),
    issues: Object.freeze([...args.issues]),
  })
}

function runErrorIssue(error: WorkbookRunError): WorkbookRunAdapterCheckIssue {
  return adapterCheckIssue(error.code, error.path ?? 'result', error.message)
}

function runtimeAdapterIssue(issue: WorkbookRuntimeAdapterIssue): WorkbookRunAdapterCheckIssue {
  if (issue.code === 'invalid_requirements') {
    return adapterCheckIssue('invalid_plan_data', issue.path ?? 'plan', issue.message)
  }
  return adapterCheckIssue('adapter_missing_capability', issue.method === undefined ? 'adapter' : `adapter.${issue.method}`, issue.message)
}

function strictOptions(options: WorkbookRunAdapterCheckOptions): WorkbookRunOptions {
  return Object.freeze({
    ...options,
    strict: true,
  })
}

export function checkWorkbookRunAdapter<Refs>(
  plan: WorkbookActionPlan<Refs>,
  adapter: WorkbookRunAdapter<Refs>,
  options?: WorkbookRunAdapterCheckOptions,
): Promise<WorkbookRunAdapterCheckResult>
export function checkWorkbookRunAdapter(
  plan: WorkbookPlanData,
  adapter: WorkbookRunAdapter<WorkbookPlanDataRefs>,
  options?: WorkbookRunAdapterCheckOptions,
): Promise<WorkbookRunAdapterCheckResult>
export async function checkWorkbookRunAdapter(
  plan: WorkbookExecutablePlan,
  adapter: WorkbookRunAdapter,
  options: WorkbookRunAdapterCheckOptions = {},
): Promise<WorkbookRunAdapterCheckResult> {
  const adapterCheck = checkRuntimeAdapter(plan, adapter)
  if (adapterCheck.status === 'invalid') {
    return failedAdapterCheck({
      issues: adapterCheck.issues.map(runtimeAdapterIssue),
    })
  }

  const result = await runWorkbookPlan(plan, adapter, strictOptions(options))
  const description = describeRunResult(result)
  const descriptionCheck = checkWorkbookRunResultDescription(description)
  if (descriptionCheck.status === 'invalid') {
    return failedAdapterCheck({
      result,
      description,
      issues: descriptionCheck.issues.map((issue) => adapterCheckIssue('invalid_run_result_description', issue.path, issue.message)),
    })
  }

  if (result.status === 'failed') {
    return failedAdapterCheck({
      result,
      description,
      errors: result.errors,
      issues: result.errors.map(runErrorIssue),
    })
  }
  if (description.status !== 'done') {
    return failedAdapterCheck({
      result,
      description,
      issues: [
        adapterCheckIssue(
          'invalid_run_result_description',
          'status',
          'Workbook run adapter returned a done result with a non-done description',
        ),
      ],
    })
  }

  return Object.freeze({
    status: 'passed',
    result,
    description,
    issues: Object.freeze([] as const),
  })
}

export function assertWorkbookRunAdapter<Refs>(
  plan: WorkbookActionPlan<Refs>,
  adapter: WorkbookRunAdapter<Refs>,
  options?: WorkbookRunAdapterCheckOptions,
): Promise<WorkbookRunAdapterPassedResult['description']>
export function assertWorkbookRunAdapter(
  plan: WorkbookPlanData,
  adapter: WorkbookRunAdapter<WorkbookPlanDataRefs>,
  options?: WorkbookRunAdapterCheckOptions,
): Promise<WorkbookRunAdapterPassedResult['description']>
export async function assertWorkbookRunAdapter(
  plan: WorkbookExecutablePlan,
  adapter: WorkbookRunAdapter,
  options: WorkbookRunAdapterCheckOptions = {},
): Promise<WorkbookRunAdapterPassedResult['description']> {
  const check = await checkWorkbookRunAdapter(plan, adapter, options)
  if (check.status === 'passed') {
    return check.description
  }
  const [firstIssue] = check.issues
  throw new Error(firstIssue?.message ?? 'Workbook run adapter check failed')
}
