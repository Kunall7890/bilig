import { parseFormula } from '@bilig/formula'
import { describePlanResult, describeRef, type WorkbookActionPlanResultDescription, type WorkbookRefDescription } from './describe.js'
import type { WorkbookRef } from './find.js'
import {
  inspectModel,
  planWorkbookAction,
  type WorkbookActionCommand,
  type WorkbookActionMap,
  type WorkbookActionPlan,
  type WorkbookModel,
} from './model.js'
import type { WorkbookOp } from './ops.js'

type WorkbookConcreteCommandOp = Extract<WorkbookOp, { kind: 'setCellFormula' | 'setCellValue' | 'clearCell' }>

export type WorkbookPlanIssueCode =
  | 'duplicate_ref'
  | 'command_target_not_resolved'
  | 'formula_input_not_resolved'
  | 'invalid_formula'
  | 'change_target_not_resolved'
  | 'check_target_not_resolved'
  | 'check_ref_not_resolved'
  | 'missing_concrete_op'

export interface WorkbookPlanIssue {
  readonly code: WorkbookPlanIssueCode
  readonly message: string
  readonly path: string
  readonly ref?: WorkbookRefDescription
}

export interface WorkbookPlanVerification {
  readonly status: 'valid' | 'invalid'
  readonly modelName: string
  readonly actionName: string
  readonly issues: readonly WorkbookPlanIssue[]
}

export interface WorkbookModelActionVerification {
  readonly actionName: string
  readonly planning: WorkbookActionPlanResultDescription
  readonly verification?: WorkbookPlanVerification
}

export interface WorkbookModelVerification {
  readonly status: 'valid' | 'invalid'
  readonly modelName: string
  readonly actions: readonly WorkbookModelActionVerification[]
}

function refKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
}

function describeProblemRef(ref: WorkbookRef): WorkbookRefDescription {
  return describeRef(ref)
}

function issue(input: {
  readonly code: WorkbookPlanIssueCode
  readonly message: string
  readonly path: string
  readonly ref?: WorkbookRef
}): WorkbookPlanIssue {
  return {
    code: input.code,
    message: input.message,
    path: input.path,
    ...(input.ref !== undefined ? { ref: describeProblemRef(input.ref) } : {}),
  }
}

function concreteSingleCell(target: WorkbookRef): { sheetName: string; address: string } | null {
  if (target.kind !== 'range') {
    return null
  }
  const range = target.range
  return range.startAddress === range.endAddress ? { sheetName: range.sheetName, address: range.startAddress } : null
}

function expectedConcreteOp(command: WorkbookActionCommand): WorkbookConcreteCommandOp | null {
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
    case 'clear':
      return {
        kind: 'clearCell',
        sheetName: target.sheetName,
        address: target.address,
      }
    case 'format':
      return null
  }
}

function opMatches(expected: WorkbookConcreteCommandOp, actual: WorkbookOp): boolean {
  if (expected.kind !== actual.kind) {
    return false
  }
  switch (expected.kind) {
    case 'setCellFormula':
      return (
        actual.kind === 'setCellFormula' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.formula === expected.formula
      )
    case 'setCellValue':
      return (
        actual.kind === 'setCellValue' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.value === expected.value
      )
    case 'clearCell':
      return actual.kind === 'clearCell' && actual.sheetName === expected.sheetName && actual.address === expected.address
    default:
      return false
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export function verifyPlan<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookPlanVerification {
  const issues: WorkbookPlanIssue[] = []
  const knownRefs = new Set<string>()

  plan.refsUsed.forEach((ref, index) => {
    const key = refKey(ref)
    if (knownRefs.has(key)) {
      issues.push(
        issue({
          code: 'duplicate_ref',
          path: `refsUsed[${index}]`,
          ref,
          message: `${ref.label} appears more than once in refsUsed`,
        }),
      )
      return
    }
    knownRefs.add(key)
  })

  function hasRef(ref: WorkbookRef): boolean {
    return knownRefs.has(refKey(ref))
  }

  plan.commands.forEach((command, commandIndex) => {
    if (!hasRef(command.target)) {
      issues.push(
        issue({
          code: 'command_target_not_resolved',
          path: `commands[${commandIndex}].target`,
          ref: command.target,
          message: `${command.target.label} is used as a command target but is missing from refsUsed`,
        }),
      )
    }

    if (command.kind === 'writeFormula') {
      try {
        parseFormula(command.formula)
      } catch (error) {
        issues.push({
          code: 'invalid_formula',
          path: `commands[${commandIndex}].formula`,
          message: `Formula for ${command.target.label} is not parseable: ${errorMessage(error)}`,
        })
      }

      command.inputs.forEach((input, inputIndex) => {
        if (!hasRef(input)) {
          issues.push(
            issue({
              code: 'formula_input_not_resolved',
              path: `commands[${commandIndex}].inputs[${inputIndex}]`,
              ref: input,
              message: `${input.label} is used as a formula input but is missing from refsUsed`,
            }),
          )
        }
      })
    }

    const expectedOp = expectedConcreteOp(command)
    if (expectedOp !== null && !plan.ops.some((op) => opMatches(expectedOp, op))) {
      issues.push(
        issue({
          code: 'missing_concrete_op',
          path: `commands[${commandIndex}]`,
          ref: command.target,
          message: `${command.target.label} has no matching concrete workbook op`,
        }),
      )
    }
  })

  plan.changed.forEach((change, changeIndex) => {
    if (change.target !== undefined && !hasRef(change.target)) {
      issues.push(
        issue({
          code: 'change_target_not_resolved',
          path: `changed[${changeIndex}].target`,
          ref: change.target,
          message: `${change.target.label} appears in changed but is missing from refsUsed`,
        }),
      )
    }
  })

  plan.checks.forEach((check, checkIndex) => {
    if (check.target !== undefined && !hasRef(check.target)) {
      issues.push(
        issue({
          code: 'check_target_not_resolved',
          path: `checks[${checkIndex}].target`,
          ref: check.target,
          message: `${check.target.label} appears in checks but is missing from refsUsed`,
        }),
      )
    }

    check.refs?.forEach((ref, refIndex) => {
      if (!hasRef(ref)) {
        issues.push(
          issue({
            code: 'check_ref_not_resolved',
            path: `checks[${checkIndex}].refs[${refIndex}]`,
            ref,
            message: `${ref.label} appears in checks but is missing from refsUsed`,
          }),
        )
      }
    })
  })

  return {
    status: issues.length === 0 ? 'valid' : 'invalid',
    modelName: plan.modelName,
    actionName: plan.actionName,
    issues,
  }
}

export function verifyModel<Refs, Actions extends WorkbookActionMap<Refs>>(model: WorkbookModel<Refs, Actions>): WorkbookModelVerification {
  const actions = inspectModel(model).actions.map((actionName): WorkbookModelActionVerification => {
    const planning = planWorkbookAction(model, actionName)
    const describedPlanning = describePlanResult(planning)
    if (planning.status === 'failed') {
      return {
        actionName,
        planning: describedPlanning,
      }
    }
    return {
      actionName,
      planning: describedPlanning,
      verification: verifyPlan(planning.plan),
    }
  })

  const isValid = actions.every((action) => action.planning.status === 'planned' && action.verification?.status === 'valid')

  return {
    status: isValid ? 'valid' : 'invalid',
    modelName: model.name,
    actions,
  }
}
