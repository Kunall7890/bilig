import type { CellRangeRef } from '@bilig/protocol'
import { parseFormula } from '@bilig/formula'
import {
  describePlanResult,
  describeRef,
  type WorkbookActionPlanDescription,
  type WorkbookActionPlanResultDescription,
  type WorkbookRefDescription,
} from './describe.js'
import { collectWorkbookRefs, type WorkbookRef } from './find.js'
import type { WorkbookFormulaLabel } from './formula.js'
import { isWorkbookOp } from './guards.js'
import {
  inspectModel,
  planWorkbookAction,
  type WorkbookActionCommand,
  type WorkbookActionMap,
  type WorkbookActionPlan,
  type WorkbookModel,
} from './model.js'
import { describeRuntimeRequirements, type WorkbookRuntimeRequirements } from './requirements.js'
import {
  getOwnActionInput,
  hasOwnActionInput,
  WorkbookActionInputError,
  normalizeWorkbookActionInput,
  type WorkbookActionInput,
} from './input.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookCheckExpectation } from './result.js'
import { hydratePlanData } from './plan-data.js'

type WorkbookConcreteCommandOp = Extract<WorkbookOp, { kind: 'setCellFormula' | 'setCellValue' | 'setCellFormat' | 'clearCell' }>

export type WorkbookPlanIssueCode =
  | 'invalid_action_input'
  | 'ref_not_in_refs'
  | 'duplicate_ref'
  | 'command_target_not_resolved'
  | 'formula_input_not_resolved'
  | 'formula_input_not_labeled'
  | 'formula_label_not_resolved'
  | 'formula_label_not_used'
  | 'invalid_formula'
  | 'change_target_not_resolved'
  | 'check_status_not_planned'
  | 'check_target_not_resolved'
  | 'check_ref_not_resolved'
  | 'check_expectation_input_not_resolved'
  | 'check_expectation_input_not_labeled'
  | 'check_expectation_label_not_resolved'
  | 'check_expectation_label_not_used'
  | 'invalid_check_expectation_formula'
  | 'invalid_workbook_op'
  | 'op_target_mismatch'
  | 'missing_concrete_op'
  | 'missing_workbook_op'

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
  readonly requirements?: WorkbookRuntimeRequirements
  readonly verification?: WorkbookPlanVerification
}

export interface WorkbookModelVerification {
  readonly status: 'valid' | 'invalid'
  readonly modelName: string
  readonly actions: readonly WorkbookModelActionVerification[]
}

export interface WorkbookModelVerificationOptions {
  readonly inputs?: Partial<Record<string, WorkbookActionInput>>
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
  if (command.kind === 'op') {
    return null
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
    case 'clear':
      return {
        kind: 'clearCell',
        sheetName: target.sheetName,
        address: target.address,
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
  }
}

function commandTarget(command: WorkbookActionCommand): WorkbookRef | undefined {
  return command.target
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
    case 'setCellFormat':
      return (
        actual.kind === 'setCellFormat' &&
        actual.sheetName === expected.sheetName &&
        actual.address === expected.address &&
        actual.format === expected.format
      )
    case 'clearCell':
      return actual.kind === 'clearCell' && actual.sheetName === expected.sheetName && actual.address === expected.address
    default:
      return false
  }
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
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

function workbookOpMatches(expected: WorkbookOp, actual: WorkbookOp): boolean {
  return canonicalJson(actual) === canonicalJson(expected)
}

function singleCellRange(sheetName: string, address: string): CellRangeRef {
  return {
    sheetName,
    startAddress: address,
    endAddress: address,
  }
}

function tableRange(table: { readonly sheetName: string; readonly startAddress: string; readonly endAddress: string }): CellRangeRef {
  return {
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
  }
}

function rangesEqual(left: CellRangeRef, right: CellRangeRef): boolean {
  return left.sheetName === right.sheetName && left.startAddress === right.startAddress && left.endAddress === right.endAddress
}

function opTargetRange(op: WorkbookOp) {
  if (
    op.kind === 'setCellValue' ||
    op.kind === 'setCellFormula' ||
    op.kind === 'setCellFormat' ||
    op.kind === 'clearCell' ||
    op.kind === 'deleteCommentThread' ||
    op.kind === 'deleteNote' ||
    op.kind === 'deleteHyperlink' ||
    op.kind === 'upsertCommentThread' ||
    op.kind === 'upsertNote' ||
    op.kind === 'upsertHyperlink' ||
    op.kind === 'upsertSpillRange' ||
    op.kind === 'deleteSpillRange' ||
    op.kind === 'upsertPivotTable' ||
    op.kind === 'deletePivotTable' ||
    op.kind === 'upsertChart' ||
    op.kind === 'upsertImage' ||
    op.kind === 'upsertShape'
  ) {
    if (op.kind === 'upsertCommentThread') {
      return singleCellRange(op.thread.sheetName, op.thread.address)
    }
    if (op.kind === 'upsertNote') {
      return singleCellRange(op.note.sheetName, op.note.address)
    }
    if (op.kind === 'upsertHyperlink') {
      return singleCellRange(op.hyperlink.sheetName, op.hyperlink.address)
    }
    if (op.kind === 'upsertChart') {
      return singleCellRange(op.chart.sheetName, op.chart.address)
    }
    if (op.kind === 'upsertImage') {
      return singleCellRange(op.image.sheetName, op.image.address)
    }
    if (op.kind === 'upsertShape') {
      return singleCellRange(op.shape.sheetName, op.shape.address)
    }
    return singleCellRange(op.sheetName, op.address)
  }

  if (op.kind === 'mergeCells' || op.kind === 'unmergeCells' || op.kind === 'setStyleRange' || op.kind === 'setFormatRange') {
    return op.range
  }

  if (
    op.kind === 'setFilter' ||
    op.kind === 'clearFilter' ||
    op.kind === 'setSort' ||
    op.kind === 'clearSort' ||
    op.kind === 'clearDataValidation'
  ) {
    return {
      ...op.range,
      sheetName: op.sheetName,
    }
  }

  if (op.kind === 'setDataValidation') {
    return op.validation.range
  }
  if (op.kind === 'upsertConditionalFormat') {
    return op.format.range
  }
  if (op.kind === 'upsertRangeProtection') {
    return op.protection.range
  }
  if (op.kind === 'upsertTable') {
    return tableRange(op.table)
  }

  return null
}

function opTargetMismatch(op: WorkbookOp, target: WorkbookRef | undefined): string | null {
  if (target === undefined) {
    return null
  }

  const range = opTargetRange(op)
  if (range === null) {
    return null
  }

  if (target.kind !== 'range') {
    return `Low-level workbook op ${op.kind} targets ${range.sheetName}!${range.startAddress}:${range.endAddress}, but ${target.label} is not a range ref`
  }

  if (!rangesEqual(range, target.range)) {
    return `Low-level workbook op ${op.kind} targets ${range.sheetName}!${range.startAddress}:${range.endAddress}, but command target is ${target.label}`
  }

  return null
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function formulaLabelsForCommand(
  command: Extract<WorkbookActionCommand, { readonly kind: 'writeFormula' }>,
): readonly WorkbookFormulaLabel[] {
  return (command as { readonly labels?: readonly WorkbookFormulaLabel[] }).labels ?? []
}

function formulaLabelsForExpectation(
  expectation: Extract<WorkbookCheckExpectation, { readonly kind: 'formulaEquals' }>,
): readonly WorkbookFormulaLabel[] {
  return (expectation as { readonly labels?: readonly WorkbookFormulaLabel[] }).labels ?? []
}

function hasLabelForInput(labels: readonly WorkbookFormulaLabel[], input: WorkbookRef): boolean {
  const inputKey = refKey(input)
  return labels.some((label) => refKey(label.ref) === inputKey)
}

export function verifyPlan<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookPlanVerification {
  const issues: WorkbookPlanIssue[] = []
  const knownRefs = new Set<string>()
  const refsInShape = new Set(collectWorkbookRefs(plan.refs).map(refKey))

  if (hasOwnActionInput(plan)) {
    try {
      normalizeWorkbookActionInput(getOwnActionInput(plan))
    } catch (error) {
      issues.push({
        code: 'invalid_action_input',
        path: error instanceof WorkbookActionInputError ? error.path : 'input',
        message: errorMessage(error),
      })
    }
  }

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
    if (!refsInShape.has(key)) {
      issues.push(
        issue({
          code: 'ref_not_in_refs',
          path: `refsUsed[${index}]`,
          ref,
          message: `${ref.label} appears in refsUsed but is not discoverable from refs`,
        }),
      )
    }
  })

  function hasRef(ref: WorkbookRef): boolean {
    return knownRefs.has(refKey(ref))
  }

  plan.commands.forEach((command, commandIndex) => {
    const target = commandTarget(command)

    if (target !== undefined && !hasRef(target)) {
      issues.push(
        issue({
          code: 'command_target_not_resolved',
          path: `commands[${commandIndex}].target`,
          ref: target,
          message: `${target.label} is used as a command target but is missing from refsUsed`,
        }),
      )
    }

    if (command.kind === 'writeFormula') {
      const labels = formulaLabelsForCommand(command)
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
        if (!hasLabelForInput(labels, input)) {
          issues.push(
            issue({
              code: 'formula_input_not_labeled',
              path: `commands[${commandIndex}].inputs[${inputIndex}]`,
              ref: input,
              message: `${input.label} is used as a formula input but has no formula label`,
            }),
          )
        }
      })

      const inputKeys = new Set(command.inputs.map(refKey))
      labels.forEach((label, labelIndex) => {
        if (!hasRef(label.ref) && !inputKeys.has(refKey(label.ref))) {
          issues.push(
            issue({
              code: 'formula_label_not_resolved',
              path: `commands[${commandIndex}].labels[${labelIndex}].ref`,
              ref: label.ref,
              message: `${label.ref.label} is used as a formula label but is missing from refsUsed`,
            }),
          )
        }
        if (!command.formula.includes(label.name)) {
          issues.push(
            issue({
              code: 'formula_label_not_used',
              path: `commands[${commandIndex}].labels[${labelIndex}].name`,
              ref: label.ref,
              message: `Formula label ${label.name} for ${label.ref.label} does not appear in the formula`,
            }),
          )
        }
      })
    }

    if (command.kind === 'op') {
      if (!isWorkbookOp(command.op)) {
        issues.push({
          code: 'invalid_workbook_op',
          path: `commands[${commandIndex}].op`,
          message: 'Low-level workbook op is not a valid WorkbookOp',
        })
      } else {
        const mismatch = opTargetMismatch(command.op, command.target)
        if (mismatch !== null) {
          issues.push(
            issue({
              code: 'op_target_mismatch',
              path: `commands[${commandIndex}].target`,
              ...(command.target !== undefined ? { ref: command.target } : {}),
              message: mismatch,
            }),
          )
        }
        if (!plan.ops.some((op) => workbookOpMatches(command.op, op))) {
          issues.push({
            code: 'missing_workbook_op',
            path: `commands[${commandIndex}].op`,
            message: `Low-level workbook op ${command.op.kind} is missing from plan ops`,
          })
        }
      }
    }

    const expectedOp = expectedConcreteOp(command)
    if (expectedOp !== null && !plan.ops.some((op) => opMatches(expectedOp, op))) {
      const expectedTarget = commandTarget(command)
      issues.push(
        issue({
          code: 'missing_concrete_op',
          path: `commands[${commandIndex}]`,
          ...(expectedTarget !== undefined ? { ref: expectedTarget } : {}),
          message: `${expectedTarget?.label ?? 'Command'} has no matching concrete workbook op`,
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
    if (check.status !== 'planned') {
      issues.push({
        code: 'check_status_not_planned',
        path: `checks[${checkIndex}].status`,
        message: `${check.target?.label ?? check.kind} check ${check.kind} must start planned before runtime proof`,
      })
    }

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

    if (check.expectation?.kind === 'formulaEquals') {
      const expectation = check.expectation
      const labels = formulaLabelsForExpectation(expectation)
      try {
        parseFormula(expectation.formula)
      } catch (error) {
        issues.push({
          code: 'invalid_check_expectation_formula',
          path: `checks[${checkIndex}].expectation.formula`,
          message: `Formula expectation for ${check.target?.label ?? check.kind} is not parseable: ${errorMessage(error)}`,
        })
      }

      expectation.inputs.forEach((input, inputIndex) => {
        if (!hasRef(input)) {
          issues.push(
            issue({
              code: 'check_expectation_input_not_resolved',
              path: `checks[${checkIndex}].expectation.inputs[${inputIndex}]`,
              ref: input,
              message: `${input.label} appears in a formula expectation but is missing from refsUsed`,
            }),
          )
        }
        if (!hasLabelForInput(labels, input)) {
          issues.push(
            issue({
              code: 'check_expectation_input_not_labeled',
              path: `checks[${checkIndex}].expectation.inputs[${inputIndex}]`,
              ref: input,
              message: `${input.label} appears in a formula expectation but has no formula label`,
            }),
          )
        }
      })

      const inputKeys = new Set(expectation.inputs.map(refKey))
      labels.forEach((label, labelIndex) => {
        if (!hasRef(label.ref) && !inputKeys.has(refKey(label.ref))) {
          issues.push(
            issue({
              code: 'check_expectation_label_not_resolved',
              path: `checks[${checkIndex}].expectation.labels[${labelIndex}].ref`,
              ref: label.ref,
              message: `${label.ref.label} appears in a formula expectation label but is missing from refsUsed`,
            }),
          )
        }
        if (!expectation.formula.includes(label.name)) {
          issues.push(
            issue({
              code: 'check_expectation_label_not_used',
              path: `checks[${checkIndex}].expectation.labels[${labelIndex}].name`,
              ref: label.ref,
              message: `Formula expectation label ${label.name} for ${label.ref.label} does not appear in the formula`,
            }),
          )
        }
      })
    }
  })

  return {
    status: issues.length === 0 ? 'valid' : 'invalid',
    modelName: plan.modelName,
    actionName: plan.actionName,
    issues,
  }
}

export function verifyPlanData(plan: WorkbookActionPlanDescription): WorkbookPlanVerification {
  return verifyPlan(hydratePlanData(plan))
}

export function verifyModel<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
  options: WorkbookModelVerificationOptions = {},
): WorkbookModelVerification {
  const inspection = inspectModel(model)
  const actions = inspection.actions.map((actionName): WorkbookModelActionVerification => {
    const planning = planWorkbookAction(model, actionName, options.inputs?.[actionName])
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
      requirements: describeRuntimeRequirements(planning.plan),
      verification: verifyPlan(planning.plan),
    }
  })

  const isValid = actions.every((action) => action.planning.status === 'planned' && action.verification?.status === 'valid')

  return {
    status: isValid ? 'valid' : 'invalid',
    modelName: inspection.name,
    actions,
  }
}
