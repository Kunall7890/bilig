import type { WorkbookActionInput } from './input.js'
import {
  planWorkbookAction,
  type WorkbookActionMap,
  type WorkbookActionPlan,
  type WorkbookActionPlanResult,
  type WorkbookModel,
} from './model.js'
import { toPlanData, workbookPlanId, type WorkbookPlanData, type WorkbookPlanId } from './plan-data.js'
import { describeRuntimeRequirements, type WorkbookRuntimeRequirements } from './requirements.js'
import type { WorkbookRunError } from './result.js'
import { verifyPlan, type WorkbookPlanIssue, type WorkbookPlanVerification } from './verify.js'

export type WorkbookActionPreparation<Refs = unknown> =
  | {
      readonly status: 'prepared'
      readonly plan: WorkbookActionPlan<Refs>
      readonly planData: WorkbookPlanData
      readonly planId: WorkbookPlanId
      readonly requirements: WorkbookRuntimeRequirements
      readonly verification: WorkbookPlanVerification
    }
  | {
      readonly status: 'failed'
      readonly planning: WorkbookActionPlanResult<Refs>
      readonly errors: readonly WorkbookRunError[]
      readonly issues: readonly WorkbookPlanIssue[]
      readonly verification?: WorkbookPlanVerification
    }

function failedPreparation<Refs>(
  planning: WorkbookActionPlanResult<Refs>,
  errors: readonly WorkbookRunError[],
  issues: readonly WorkbookPlanIssue[] = [],
  verification?: WorkbookPlanVerification,
): WorkbookActionPreparation<Refs> {
  return Object.freeze({
    status: 'failed',
    planning,
    errors: Object.freeze([...errors]),
    issues: Object.freeze([...issues]),
    ...(verification !== undefined ? { verification } : {}),
  })
}

function errorsForIssues(issues: readonly WorkbookPlanIssue[]): readonly WorkbookRunError[] {
  return Object.freeze(
    issues.map((issue) =>
      Object.freeze({
        code: issue.code,
        message: issue.message,
        path: issue.path,
        issueCode: issue.code,
      }),
    ),
  )
}

export function prepareWorkbookAction<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
  actionName: string,
  input?: WorkbookActionInput,
): WorkbookActionPreparation<Refs>
export function prepareWorkbookAction(model: unknown, actionName: string, input?: WorkbookActionInput): WorkbookActionPreparation
export function prepareWorkbookAction(model: unknown, actionName: unknown, input?: WorkbookActionInput): WorkbookActionPreparation {
  const planning = planWorkbookAction(model, actionName, input)
  if (planning.status === 'failed') {
    return failedPreparation(planning, planning.errors)
  }

  const verification = verifyPlan(planning.plan)
  if (verification.status === 'invalid') {
    return failedPreparation(planning, errorsForIssues(verification.issues), verification.issues, verification)
  }

  const planData = toPlanData(planning.plan)
  return Object.freeze({
    status: 'prepared',
    plan: planning.plan,
    planData,
    planId: workbookPlanId(planData),
    requirements: describeRuntimeRequirements(planData),
    verification,
  })
}
