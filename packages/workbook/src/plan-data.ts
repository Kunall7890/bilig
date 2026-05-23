import { isLiteralInput } from '@bilig/protocol'
import {
  describePlan,
  type WorkbookActionCommandDescription,
  type WorkbookActionPlanDescription,
  type WorkbookChangeSummaryDescription,
  type WorkbookCheckExpectationDescription,
  type WorkbookCheckResultDescription,
  type WorkbookFormulaLabelDescription,
  type WorkbookRefDescription,
} from './describe.js'
import { hydrateWorkbookRef, isWorkbookRefData, type WorkbookRef } from './find.js'
import type { WorkbookFormulaLabel } from './formula.js'
import { isWorkbookOp } from './guards.js'
import { isWorkbookActionInput, normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import type { WorkbookActionCommand, WorkbookActionPlan } from './model.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookChangeSummary, WorkbookCheckExpectation, WorkbookCheckResult, WorkbookCheckStatus } from './result.js'

export type WorkbookPlanData = WorkbookActionPlanDescription

export interface WorkbookPlanDataRefs {
  readonly refsUsed: readonly WorkbookRef[]
}

export type WorkbookExecutablePlan<Refs = unknown> = WorkbookActionPlan<Refs> | WorkbookPlanData

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function hasString(value: Record<string, unknown>, key: string): boolean {
  return typeof value[key] === 'string'
}

function hasOptionalString(value: Record<string, unknown>, key: string): boolean {
  const entry = value[key]
  return entry === undefined || typeof entry === 'string'
}

function isWorkbookRefDescription(value: unknown): value is WorkbookRefDescription {
  return isWorkbookRefData(value)
}

function isFormulaLabelData(value: unknown): value is WorkbookFormulaLabelDescription {
  return isRecord(value) && hasString(value, 'name') && isWorkbookRefDescription(value['ref'])
}

function isFormulaLabelDataArray(value: unknown): value is readonly WorkbookFormulaLabelDescription[] {
  return Array.isArray(value) && value.every(isFormulaLabelData)
}

function isRefDataArray(value: unknown): value is readonly WorkbookRefDescription[] {
  return Array.isArray(value) && value.every(isWorkbookRefDescription)
}

function isWorkbookActionCommandData(value: unknown): value is WorkbookActionCommandDescription {
  if (!isRecord(value) || !hasString(value, 'kind')) {
    return false
  }
  switch (value['kind']) {
    case 'writeFormula':
      return (
        isWorkbookRefDescription(value['target']) &&
        typeof value['formula'] === 'string' &&
        isRefDataArray(value['inputs']) &&
        isFormulaLabelDataArray(value['labels'])
      )
    case 'writeValue':
      return isWorkbookRefDescription(value['target']) && isLiteralInput(value['value'])
    case 'format':
      return (
        isWorkbookRefDescription(value['target']) &&
        (value['style'] === undefined || isRecord(value['style'])) &&
        (value['numberFormat'] === undefined || typeof value['numberFormat'] === 'string' || value['numberFormat'] === null)
      )
    case 'clear':
      return isWorkbookRefDescription(value['target'])
    case 'op':
      return (
        isWorkbookOp(value['op']) &&
        (value['target'] === undefined || isWorkbookRefDescription(value['target'])) &&
        hasOptionalString(value, 'message')
      )
    default:
      return false
  }
}

function isChangeData(value: unknown): value is WorkbookChangeSummaryDescription {
  return (
    isRecord(value) &&
    hasString(value, 'kind') &&
    hasString(value, 'message') &&
    (value['target'] === undefined || isWorkbookRefDescription(value['target']))
  )
}

function isCheckStatus(value: unknown): value is WorkbookCheckStatus {
  return value === 'planned' || value === 'passed' || value === 'failed'
}

function isCheckExpectationData(value: unknown): value is WorkbookCheckExpectationDescription {
  if (!isRecord(value) || !hasString(value, 'kind')) {
    return false
  }
  switch (value['kind']) {
    case 'valueEquals':
      return isLiteralInput(value['value'])
    case 'formulaEquals':
      return typeof value['formula'] === 'string' && isRefDataArray(value['inputs']) && isFormulaLabelDataArray(value['labels'])
    default:
      return false
  }
}

function isCheckData(value: unknown): value is WorkbookCheckResultDescription {
  return (
    isRecord(value) &&
    isCheckStatus(value['status']) &&
    hasString(value, 'kind') &&
    hasString(value, 'message') &&
    (value['target'] === undefined || isWorkbookRefDescription(value['target'])) &&
    (value['refs'] === undefined || isRefDataArray(value['refs'])) &&
    (value['expectation'] === undefined || isCheckExpectationData(value['expectation'])) &&
    (value['proof'] === undefined || isWorkbookActionInput(value['proof']))
  )
}

export function isPlanData(value: unknown): value is WorkbookPlanData {
  return (
    isRecord(value) &&
    hasString(value, 'modelName') &&
    hasString(value, 'actionName') &&
    (value['input'] === undefined || isWorkbookActionInput(value['input'])) &&
    isRefDataArray(value['refsUsed']) &&
    Array.isArray(value['commands']) &&
    value['commands'].every(isWorkbookActionCommandData) &&
    Array.isArray(value['ops']) &&
    value['ops'].every(isWorkbookOp) &&
    Array.isArray(value['changed']) &&
    value['changed'].every(isChangeData) &&
    Array.isArray(value['checks']) &&
    value['checks'].every(isCheckData)
  )
}

function hydrateRef(ref: WorkbookRefDescription): WorkbookRef {
  return hydrateWorkbookRef(ref)
}

function hydrateFormulaLabel(label: WorkbookFormulaLabelDescription): WorkbookFormulaLabel {
  return Object.freeze({
    name: label.name,
    ref: hydrateRef(label.ref),
  })
}

function hydrateCommand(command: WorkbookActionCommandDescription): WorkbookActionCommand {
  switch (command.kind) {
    case 'writeFormula':
      return Object.freeze({
        kind: 'writeFormula',
        target: hydrateRef(command.target),
        formula: command.formula,
        inputs: Object.freeze(command.inputs.map(hydrateRef)),
        labels: Object.freeze(command.labels.map(hydrateFormulaLabel)),
      })
    case 'writeValue':
      return Object.freeze({
        kind: 'writeValue',
        target: hydrateRef(command.target),
        value: command.value,
      })
    case 'format':
      return Object.freeze({
        kind: 'format',
        target: hydrateRef(command.target),
        ...(command.style !== undefined ? { style: Object.freeze(structuredClone(command.style)) } : {}),
        ...(command.numberFormat !== undefined ? { numberFormat: command.numberFormat } : {}),
      })
    case 'clear':
      return Object.freeze({
        kind: 'clear',
        target: hydrateRef(command.target),
      })
    case 'op':
      return Object.freeze({
        kind: 'op',
        op: Object.freeze(structuredClone(command.op)),
        ...(command.target !== undefined ? { target: hydrateRef(command.target) } : {}),
        ...(command.message !== undefined ? { message: command.message } : {}),
      })
  }
}

function hydrateChange(change: WorkbookChangeSummaryDescription): WorkbookChangeSummary {
  return Object.freeze({
    kind: change.kind,
    ...(change.target !== undefined ? { target: hydrateRef(change.target) } : {}),
    message: change.message,
  })
}

function hydrateExpectation(expectation: WorkbookCheckExpectationDescription): WorkbookCheckExpectation {
  switch (expectation.kind) {
    case 'valueEquals':
      return Object.freeze({
        kind: 'valueEquals',
        value: expectation.value,
      })
    case 'formulaEquals':
      return Object.freeze({
        kind: 'formulaEquals',
        formula: expectation.formula,
        inputs: Object.freeze(expectation.inputs.map(hydrateRef)),
        labels: Object.freeze(expectation.labels.map(hydrateFormulaLabel)),
      })
  }
}

function hydrateCheck(check: WorkbookCheckResultDescription): WorkbookCheckResult {
  const proof: WorkbookActionInput | undefined = check.proof === undefined ? undefined : normalizeWorkbookActionInput(check.proof)
  return Object.freeze({
    status: check.status,
    kind: check.kind,
    ...(check.target !== undefined ? { target: hydrateRef(check.target) } : {}),
    ...(check.refs !== undefined ? { refs: Object.freeze(check.refs.map(hydrateRef)) } : {}),
    message: check.message,
    ...(check.expectation !== undefined ? { expectation: hydrateExpectation(check.expectation) } : {}),
    ...(proof !== undefined ? { proof } : {}),
  })
}

function cloneOp(op: WorkbookOp): WorkbookOp {
  return Object.freeze(structuredClone(op))
}

export function toPlanData<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookPlanData {
  return describePlan(plan)
}

export function hydratePlanData(data: WorkbookPlanData): WorkbookActionPlan<WorkbookPlanDataRefs> {
  if (!isPlanData(data)) {
    throw new Error('Workbook plan data is invalid')
  }
  const refsUsed = Object.freeze(data.refsUsed.map(hydrateRef))
  const input = data.input === undefined ? undefined : normalizeWorkbookActionInput(data.input)
  return Object.freeze({
    modelName: data.modelName,
    actionName: data.actionName,
    ...(input !== undefined ? { input } : {}),
    refs: Object.freeze({ refsUsed }),
    refsUsed,
    commands: Object.freeze(data.commands.map(hydrateCommand)),
    ops: Object.freeze(data.ops.map(cloneOp)),
    changed: Object.freeze(data.changed.map(hydrateChange)),
    checks: Object.freeze(data.checks.map(hydrateCheck)),
  })
}

export function isHydratedPlan<Refs>(value: WorkbookExecutablePlan<Refs>): value is WorkbookActionPlan<Refs> {
  return isRecord(value) && isRecord(value['refs'])
}

export function executablePlan<Refs>(
  plan: WorkbookExecutablePlan<Refs>,
): WorkbookActionPlan<Refs> | WorkbookActionPlan<WorkbookPlanDataRefs> {
  return isHydratedPlan(plan) ? plan : hydratePlanData(plan)
}
