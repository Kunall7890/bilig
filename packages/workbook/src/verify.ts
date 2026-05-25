import type { CellRangeRef } from '@bilig/protocol'
import { parseFormula } from '@bilig/formula'
import {
  describePlan,
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
  type WorkbookActionPlan,
  type WorkbookModelInspection,
} from './model.js'
import { isObjectRecord, optionalDataProperty } from './data-properties.js'
import { describeRuntimeRequirements, type WorkbookRuntimeRequirements } from './requirements.js'
import {
  getOwnActionInput,
  hasOwnActionInput,
  WorkbookActionInputError,
  normalizeWorkbookActionInput,
  type WorkbookActionInput,
} from './input.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookChangeSummary, WorkbookCheckExpectation, WorkbookCheckResult, WorkbookRunError } from './result.js'
import { hydratePlanData } from './plan-data.js'
import { expectedConcreteCommandOp, workbookConcreteOpMatches, workbookOpMatches } from './command-ops.js'
import { formulaUsesLabel } from './formula-usage.js'
import { normalizeWorkbookActionFormatOptions } from './model-action-validation.js'

export type WorkbookPlanIssueCode =
  | 'invalid_plan'
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
  readonly errors?: readonly WorkbookRunError[]
  readonly actions: readonly WorkbookModelActionVerification[]
}

export interface WorkbookModelVerificationOptions {
  readonly inputs?: Partial<Record<string, WorkbookActionInput>>
}

type WorkbookVerificationInputs =
  | {
      readonly status: 'valid'
      readonly inputs?: object
    }
  | {
      readonly status: 'invalid'
      readonly error: WorkbookRunError
    }

type WorkbookVerificationActionInput =
  | {
      readonly status: 'valid'
      readonly input?: WorkbookActionInput
    }
  | {
      readonly status: 'invalid'
      readonly error: WorkbookRunError
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
  return Object.freeze({
    code: input.code,
    message: input.message,
    path: input.path,
    ...(input.ref !== undefined ? { ref: describeProblemRef(input.ref) } : {}),
  })
}

function freezePlanIssues(issues: readonly WorkbookPlanIssue[]): readonly WorkbookPlanIssue[] {
  return Object.freeze(issues.map((entry) => Object.freeze(entry)))
}

function commandTarget(command: WorkbookActionCommand): WorkbookRef | undefined {
  return command.target
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

function errorPath(error: unknown): string {
  return error instanceof WorkbookActionInputError ? error.path : 'input'
}

function childPath(path: string, key: string): string {
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(key) ? `${path}.${key}` : `${path}[${JSON.stringify(key)}]`
}

function prefixedInputPath(path: string, error: unknown): string {
  const inputPath = errorPath(error)
  return inputPath === 'input' ? path : `${path}${inputPath.slice('input'.length)}`
}

function invalidActionInputRunError(error: unknown, path: string): WorkbookRunError {
  return Object.freeze({
    code: 'invalid_action_input',
    message: errorMessage(error),
    path: prefixedInputPath(path, error),
    issueCode: 'invalid_action_input',
  })
}

function invalidVerificationInput(message: string, path: string): WorkbookVerificationInputs {
  return Object.freeze({
    status: 'invalid',
    error: Object.freeze({
      code: 'invalid_action_input',
      message,
      path,
      issueCode: 'invalid_action_input',
    }),
  })
}

function readVerificationInputs(options: unknown): WorkbookVerificationInputs {
  if (options === undefined) {
    return { status: 'valid' }
  }
  if (!isObjectRecord(options)) {
    return invalidVerificationInput('Workbook model verification options must be an object', 'options')
  }

  let inputsValue: unknown
  try {
    const inputs = optionalDataProperty(options, 'inputs', 'Workbook model verification options inputs')
    if (inputs.status === 'missing' || inputs.value === undefined) {
      return { status: 'valid' }
    }
    inputsValue = inputs.value
  } catch (error) {
    return invalidVerificationInput(errorMessage(error), 'options.inputs')
  }

  if (!isObjectRecord(inputsValue)) {
    return invalidVerificationInput('Workbook model verification options inputs must be an object', 'options.inputs')
  }

  return {
    status: 'valid',
    inputs: inputsValue,
  }
}

function readVerificationActionInput(inputs: object | undefined, actionName: string): WorkbookVerificationActionInput {
  if (inputs === undefined) {
    return { status: 'valid' }
  }

  const path = childPath('inputs', actionName)
  const descriptor = Object.getOwnPropertyDescriptor(inputs, actionName)
  if (descriptor === undefined) {
    return { status: 'valid' }
  }
  if (!('value' in descriptor)) {
    return {
      status: 'invalid',
      error: {
        code: 'invalid_action_input',
        message: `Workbook model verification input for action ${actionName} must be a data property`,
        path,
        issueCode: 'invalid_action_input',
      },
    }
  }
  if (descriptor.value === undefined) {
    return { status: 'valid' }
  }

  try {
    return {
      status: 'valid',
      input: normalizeWorkbookActionInput(descriptor.value),
    }
  } catch (error) {
    return {
      status: 'invalid',
      error: invalidActionInputRunError(error, path),
    }
  }
}

function failedVerificationAction(modelName: string, actionName: string, error: WorkbookRunError): WorkbookModelActionVerification {
  return Object.freeze({
    actionName,
    planning: describePlanResult({
      status: 'failed',
      modelName,
      actionName,
      errors: Object.freeze([error]),
      checks: Object.freeze([]),
    }),
  })
}

function ownPlanValue<T extends object, K extends keyof T>(value: T, key: K, path: string): T[K] {
  if (typeof value !== 'object' || value === null) {
    throw new Error('Workbook verification plan must be an object')
  }
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined || !('value' in descriptor)) {
    throw new Error(`Workbook verification ${path} must be a data property`)
  }
  return descriptor.value
}

function ownPlanArray<T extends object, K extends keyof T>(value: T, key: K, path: string): T[K] {
  const array = ownPlanValue(value, key, path)
  if (!Array.isArray(array)) {
    throw new Error(`Workbook verification ${path} must be an array`)
  }
  return array
}

function ownPlanString<T extends object, K extends keyof T>(value: T, key: K, path: string): T[K] {
  const text = ownPlanValue(value, key, path)
  if (typeof text !== 'string') {
    throw new Error(`Workbook verification ${path} must be a string`)
  }
  return text
}

function safePlanText(plan: unknown, key: string, fallback: string): string {
  if (typeof plan !== 'object' || plan === null) {
    return fallback
  }
  const descriptor = Object.getOwnPropertyDescriptor(plan, key)
  return descriptor !== undefined && 'value' in descriptor && typeof descriptor.value === 'string' ? descriptor.value : fallback
}

function invalidPlanVerification(plan: unknown, error: unknown): WorkbookPlanVerification {
  return Object.freeze({
    status: 'invalid',
    modelName: safePlanText(plan, 'modelName', 'unknown-model'),
    actionName: safePlanText(plan, 'actionName', 'unknown-action'),
    issues: freezePlanIssues([
      Object.freeze({
        code: 'invalid_plan',
        path: 'plan',
        message: `Workbook plan is invalid: ${errorMessage(error)}`,
      }),
    ]),
  })
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

function formatCommandOption(
  command: Extract<WorkbookActionCommand, { readonly kind: 'format' }>,
  key: 'style' | 'numberFormat',
  path: string,
): { readonly status: 'absent' } | { readonly status: 'present'; readonly value: unknown } {
  const descriptor = Object.getOwnPropertyDescriptor(command, key)
  if (descriptor === undefined) {
    return { status: 'absent' }
  }
  if (!('value' in descriptor)) {
    throw new Error(`${path} must be a data property`)
  }
  if (descriptor.value === undefined) {
    throw new Error(`${path} must not be undefined`)
  }
  return { status: 'present', value: descriptor.value }
}

function pushFormatCommandIssues(
  issues: WorkbookPlanIssue[],
  command: Extract<WorkbookActionCommand, { readonly kind: 'format' }>,
  commandIndex: number,
): void {
  try {
    const style = formatCommandOption(command, 'style', `commands[${commandIndex}].style`)
    const numberFormat = formatCommandOption(command, 'numberFormat', `commands[${commandIndex}].numberFormat`)
    normalizeWorkbookActionFormatOptions({
      ...(style.status === 'present' ? { style: style.value } : {}),
      ...(numberFormat.status === 'present' ? { numberFormat: numberFormat.value } : {}),
    })
  } catch (error) {
    issues.push({
      code: 'invalid_plan',
      path: `commands[${commandIndex}]`,
      message: `Workbook format command is invalid: ${errorMessage(error)}`,
    })
  }
}

export function verifyPlan<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookPlanVerification {
  let modelName: string
  let actionName: string
  let refs: Refs
  let refsUsed: readonly WorkbookRef[]
  let commands: readonly WorkbookActionCommand[]
  let ops: readonly WorkbookOp[]
  let changed: readonly WorkbookChangeSummary[]
  let checks: readonly WorkbookCheckResult[]

  try {
    modelName = ownPlanString(plan, 'modelName', 'plan.modelName')
    actionName = ownPlanString(plan, 'actionName', 'plan.actionName')
    refs = ownPlanValue(plan, 'refs', 'plan.refs')
    refsUsed = ownPlanArray(plan, 'refsUsed', 'plan.refsUsed')
    commands = ownPlanArray(plan, 'commands', 'plan.commands')
    ops = ownPlanArray(plan, 'ops', 'plan.ops')
    changed = ownPlanArray(plan, 'changed', 'plan.changed')
    checks = ownPlanArray(plan, 'checks', 'plan.checks')
    describePlan(plan)
  } catch (error) {
    return invalidPlanVerification(plan, error)
  }

  const issues: WorkbookPlanIssue[] = []
  const knownRefs = new Set<string>()
  const refsInShape = new Set(collectWorkbookRefs(refs).map(refKey))

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

  refsUsed.forEach((ref, index) => {
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

  commands.forEach((command, commandIndex) => {
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
      let formulaAst: ReturnType<typeof parseFormula> | undefined
      try {
        formulaAst = parseFormula(command.formula)
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
        if (formulaAst === undefined || !formulaUsesLabel(formulaAst, label.name)) {
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
        if (!ops.some((op) => workbookOpMatches(command.op, op))) {
          issues.push({
            code: 'missing_workbook_op',
            path: `commands[${commandIndex}].op`,
            message: `Low-level workbook op ${command.op.kind} is missing from plan ops`,
          })
        }
      }
    }

    if (command.kind === 'format') {
      pushFormatCommandIssues(issues, command, commandIndex)
    }

    const expectedOp = expectedConcreteCommandOp(command)
    if (expectedOp !== null && !ops.some((op) => workbookConcreteOpMatches(expectedOp, op))) {
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

  changed.forEach((change, changeIndex) => {
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

  checks.forEach((check, checkIndex) => {
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
      let formulaAst: ReturnType<typeof parseFormula> | undefined
      try {
        formulaAst = parseFormula(expectation.formula)
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
        if (formulaAst === undefined || !formulaUsesLabel(formulaAst, label.name)) {
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

  return Object.freeze({
    status: issues.length === 0 ? 'valid' : 'invalid',
    modelName,
    actionName,
    issues: freezePlanIssues(issues),
  })
}

export function verifyPlanData(plan: WorkbookActionPlanDescription): WorkbookPlanVerification {
  return verifyPlan(hydratePlanData(plan))
}

export function verifyModel(model: unknown, options?: WorkbookModelVerificationOptions): WorkbookModelVerification
export function verifyModel(model: unknown, options?: unknown): WorkbookModelVerification
export function verifyModel(model: unknown, options: unknown = {}): WorkbookModelVerification {
  let inspection: WorkbookModelInspection
  try {
    inspection = inspectModel(model)
  } catch (error) {
    return Object.freeze({
      status: 'invalid',
      modelName: invalidModelName(model),
      errors: Object.freeze([invalidModelRunError(error)]),
      actions: Object.freeze([]),
    })
  }
  const inputsResult = readVerificationInputs(options)
  if (inputsResult.status === 'invalid') {
    return Object.freeze({
      status: 'invalid',
      modelName: inspection.name,
      errors: Object.freeze([inputsResult.error]),
      actions: Object.freeze([]),
    })
  }

  const actions = Object.freeze(
    inspection.actions.map((actionName): WorkbookModelActionVerification => {
      const actionInput = readVerificationActionInput(inputsResult.inputs, actionName)
      if (actionInput.status === 'invalid') {
        return failedVerificationAction(inspection.name, actionName, actionInput.error)
      }
      const planning = planWorkbookAction(model, actionName, actionInput.input)
      const describedPlanning = describePlanResult(planning)
      if (planning.status === 'failed') {
        return Object.freeze({
          actionName,
          planning: describedPlanning,
        })
      }
      return Object.freeze({
        actionName,
        planning: describedPlanning,
        requirements: describeRuntimeRequirements(planning.plan),
        verification: verifyPlan(planning.plan),
      })
    }),
  )

  const isValid = actions.every((action) => action.planning.status === 'planned' && action.verification?.status === 'valid')

  return Object.freeze({
    status: isValid ? 'valid' : 'invalid',
    modelName: inspection.name,
    actions,
  })
}

function invalidModelRunError(error: unknown): WorkbookRunError {
  return Object.freeze({
    code: 'invalid_model',
    message: error instanceof Error ? error.message : String(error),
  })
}

function invalidModelName(model: unknown): string {
  if (!isModelNameRecord(model)) {
    return 'unknown-model'
  }
  const descriptor = Object.getOwnPropertyDescriptor(model, 'name')
  if (descriptor === undefined || !('value' in descriptor) || typeof descriptor.value !== 'string') {
    return 'unknown-model'
  }
  const name = descriptor.value.trim()
  return name !== '' && name === descriptor.value ? name : 'unknown-model'
}

function isModelNameRecord(value: unknown): value is object {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}
