import type { CellRangeRef, CellStylePatch, LiteralInput } from '@bilig/protocol'
import type {
  WorkbookColumnRef,
  WorkbookNameRef,
  WorkbookRangeRef,
  WorkbookRef,
  WorkbookRefKind,
  WorkbookRowOperator,
  WorkbookRowsRef,
  WorkbookTableRef,
} from './find.js'
import {
  inspectModel,
  type WorkbookActionInspection,
  type WorkbookActionCommand,
  type WorkbookActionMap,
  type WorkbookActionPlan,
  type WorkbookActionPlanResult,
  type WorkbookModel,
} from './model.js'
import type { WorkbookActionInput } from './input.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookPlanId } from './plan-data.js'
import type {
  WorkbookChangeSummary,
  WorkbookCheckExpectation,
  WorkbookCheckResult,
  WorkbookCheckStatus,
  WorkbookRunApplySummary,
  WorkbookRunErrorCode,
  WorkbookRunResult,
  WorkbookRunUnverified,
  WorkbookRunUnverifiedKind,
  WorkbookUndoRef,
} from './result.js'

export interface WorkbookModelDescription {
  readonly name: string
  readonly description?: string
  readonly actions: readonly string[]
  readonly actionDetails: readonly WorkbookActionInspection[]
  readonly hasChecks: boolean
}

export interface WorkbookBaseRefDescription {
  readonly kind: WorkbookRefKind
  readonly id: string
  readonly label: string
}

export interface WorkbookRangeRefDescription extends WorkbookBaseRefDescription {
  readonly kind: 'range'
  readonly range: CellRangeRef
}

export interface WorkbookNameRefDescription extends WorkbookBaseRefDescription {
  readonly kind: 'name'
  readonly name: string
}

export interface WorkbookTableRefDescription extends WorkbookBaseRefDescription {
  readonly kind: 'table'
  readonly name?: string
  readonly sheetName?: string
  readonly headers?: readonly string[]
}

export interface WorkbookColumnRefDescription extends WorkbookBaseRefDescription {
  readonly kind: 'column'
  readonly table: WorkbookTableRefDescription
  readonly rows?: WorkbookRowsRefDescription
  readonly name: string
}

export interface WorkbookRowsRefDescription extends WorkbookBaseRefDescription {
  readonly kind: 'rows'
  readonly sheetName?: string
  readonly table?: WorkbookTableRefDescription
  readonly where: {
    readonly column: string
    readonly op: WorkbookRowOperator
    readonly value: LiteralInput
  }
}

export type WorkbookRefDescription =
  | WorkbookRangeRefDescription
  | WorkbookNameRefDescription
  | WorkbookTableRefDescription
  | WorkbookColumnRefDescription
  | WorkbookRowsRefDescription

export type WorkbookActionCommandDescription =
  | {
      readonly kind: 'writeFormula'
      readonly target: WorkbookRefDescription
      readonly formula: string
      readonly inputs: readonly WorkbookRefDescription[]
      readonly labels: readonly WorkbookFormulaLabelDescription[]
    }
  | {
      readonly kind: 'writeValue'
      readonly target: WorkbookRefDescription
      readonly value: LiteralInput
    }
  | {
      readonly kind: 'format'
      readonly target: WorkbookRefDescription
      readonly style?: CellStylePatch
      readonly numberFormat?: string | null
    }
  | {
      readonly kind: 'clear'
      readonly target: WorkbookRefDescription
    }
  | {
      readonly kind: 'op'
      readonly op: WorkbookOp
      readonly target?: WorkbookRefDescription
      readonly message?: string
    }

export interface WorkbookChangeSummaryDescription {
  readonly kind: string
  readonly target?: WorkbookRefDescription
  readonly message: string
}

export interface WorkbookCheckResultDescription {
  readonly status: WorkbookCheckStatus
  readonly kind: string
  readonly target?: WorkbookRefDescription
  readonly refs?: readonly WorkbookRefDescription[]
  readonly message: string
  readonly expectation?: WorkbookCheckExpectationDescription
  readonly proof?: WorkbookActionInput
}

export type WorkbookCheckExpectationDescription =
  | {
      readonly kind: 'valueEquals'
      readonly value: LiteralInput
    }
  | {
      readonly kind: 'formulaEquals'
      readonly formula: string
      readonly inputs: readonly WorkbookRefDescription[]
      readonly labels: readonly WorkbookFormulaLabelDescription[]
    }

export interface WorkbookFormulaLabelDescription {
  readonly name: string
  readonly ref: WorkbookRefDescription
}

export interface WorkbookActionPlanDescription {
  readonly modelName: string
  readonly actionName: string
  readonly input?: WorkbookActionInput
  readonly refsUsed: readonly WorkbookRefDescription[]
  readonly commands: readonly WorkbookActionCommandDescription[]
  readonly ops: readonly WorkbookOp[]
  readonly changed: readonly WorkbookChangeSummaryDescription[]
  readonly checks: readonly WorkbookCheckResultDescription[]
}

export type WorkbookActionPlanResultDescription =
  | {
      readonly status: 'planned'
      readonly plan: WorkbookActionPlanDescription
    }
  | {
      readonly status: 'failed'
      readonly modelName: string
      readonly actionName: string
      readonly input?: WorkbookActionInput
      readonly errors: readonly WorkbookRunErrorDescription[]
      readonly checks: readonly WorkbookCheckResultDescription[]
    }

export interface WorkbookUndoRefDescription {
  readonly id: string
  readonly ops?: readonly WorkbookOp[]
}

export interface WorkbookRunApplySummaryDescription {
  readonly matched: boolean | null
  readonly planId?: WorkbookPlanId
  readonly baseRevision?: number
  readonly revision?: number
  readonly previewOps?: readonly WorkbookOp[]
  readonly appliedOps?: readonly WorkbookOp[]
  readonly commandReceipts?: readonly {
    readonly commandIndex: number
    readonly commandKind: string
    readonly commandDigest: string
    readonly previewOps: readonly WorkbookOp[]
    readonly appliedOps: readonly WorkbookOp[]
    readonly resolvedRefs?: WorkbookActionInput
    readonly proof?: WorkbookActionInput
  }[]
  readonly proof?: WorkbookActionInput
}

export interface WorkbookRunUnverifiedDescription {
  readonly kind: WorkbookRunUnverifiedKind
  readonly message: string
}

export type WorkbookRunResultDescription =
  | {
      readonly status: 'done'
      readonly apply?: WorkbookRunApplySummaryDescription
      readonly changed: readonly WorkbookChangeSummaryDescription[]
      readonly checks: readonly WorkbookCheckResultDescription[]
      readonly undo?: WorkbookUndoRefDescription
      readonly unverified?: readonly WorkbookRunUnverifiedDescription[]
    }
  | {
      readonly status: 'failed'
      readonly errors: readonly WorkbookRunErrorDescription[]
      readonly apply?: WorkbookRunApplySummaryDescription
      readonly changed: readonly WorkbookChangeSummaryDescription[]
      readonly checks: readonly WorkbookCheckResultDescription[]
      readonly undo?: WorkbookUndoRefDescription
      readonly unverified?: readonly WorkbookRunUnverifiedDescription[]
    }

export interface WorkbookRunErrorDescription {
  readonly code: WorkbookRunErrorCode
  readonly message: string
  readonly path?: string
  readonly issueCode?: string
}

export function describeModel<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
): WorkbookModelDescription {
  return inspectModel(model)
}

interface WorkbookDataDescriptor<T> {
  readonly enumerable?: boolean
  readonly value: T
}

function isDataDescriptor<T>(descriptor: PropertyDescriptor | undefined): descriptor is WorkbookDataDescriptor<T> {
  return descriptor !== undefined && 'value' in descriptor
}

function ownDataValue<T extends object, K extends keyof T>(value: T, key: K, path: string): T[K] | undefined {
  if (typeof value !== 'object' || value === null) {
    throw new Error(`Workbook description ${path} parent must be an object`)
  }
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return undefined
  }
  if (!isDataDescriptor<T[K]>(descriptor)) {
    throw new Error(`Workbook description ${path} must be a data property`)
  }
  return descriptor.value
}

function requiredOwnDataValue<T extends object, K extends keyof T>(value: T, key: K, path: string): T[K] {
  const data = ownDataValue(value, key, path)
  if (data === undefined) {
    throw new Error(`Workbook description ${path} must be a data property`)
  }
  return data
}

function mapArrayData<T, Result>(
  value: readonly T[],
  path: string,
  mapper: (entry: T, index: number, entryPath: string) => Result,
): readonly Result[] {
  if (!Array.isArray(value)) {
    throw new Error(`Workbook description ${path} must be an array`)
  }
  const mapped: Result[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (!isDataDescriptor<T>(descriptor) || !descriptor.enumerable) {
      throw new Error(`Workbook description ${path}[${index}] must be a data property`)
    }
    mapped.push(mapper(descriptor.value, index, `${path}[${index}]`))
  }
  return mapped
}

function cloneDescriptionData<T>(value: T, path: string, seen?: WeakMap<object, unknown>): T
function cloneDescriptionData(value: unknown, path: string, seen = new WeakMap<object, unknown>()): unknown {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  const existing = seen.get(value)
  if (existing !== undefined) {
    return existing
  }
  if (Array.isArray(value)) {
    const cloned: unknown[] = []
    seen.set(value, cloned)
    mapArrayData(value, path, (entry, _index, entryPath) => {
      cloned.push(cloneDescriptionData(entry, entryPath, seen))
      return undefined
    })
    return cloned
  }
  const cloned: Record<string, unknown> = {}
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (!descriptor.enumerable) {
      return
    }
    if (!isDataDescriptor<unknown>(descriptor)) {
      throw new Error(`Workbook description ${path}.${key} must be a data property`)
    }
    Object.defineProperty(cloned, key, {
      configurable: true,
      enumerable: true,
      value: cloneDescriptionData(descriptor.value, `${path}.${key}`, seen),
      writable: true,
    })
  })
  return cloned
}

function hasRefKind<K extends WorkbookRefKind>(ref: WorkbookRef, kind: K): ref is Extract<WorkbookRef, { readonly kind: K }> {
  return requiredOwnDataValue(ref, 'kind', 'ref.kind') === kind
}

function describeRangeRef(ref: WorkbookRangeRef): WorkbookRangeRefDescription {
  const range = requiredOwnDataValue(ref, 'range', 'ref.range')
  return {
    kind: 'range',
    id: requiredOwnDataValue(ref, 'id', 'ref.id'),
    label: requiredOwnDataValue(ref, 'label', 'ref.label'),
    range: {
      sheetName: requiredOwnDataValue(range, 'sheetName', 'ref.range.sheetName'),
      startAddress: requiredOwnDataValue(range, 'startAddress', 'ref.range.startAddress'),
      endAddress: requiredOwnDataValue(range, 'endAddress', 'ref.range.endAddress'),
    },
  }
}

function describeNameRef(ref: WorkbookNameRef): WorkbookNameRefDescription {
  return {
    kind: 'name',
    id: requiredOwnDataValue(ref, 'id', 'ref.id'),
    label: requiredOwnDataValue(ref, 'label', 'ref.label'),
    name: requiredOwnDataValue(ref, 'name', 'ref.name'),
  }
}

function describeTableRef(ref: WorkbookTableRef): WorkbookTableRefDescription {
  const name = ownDataValue(ref, 'name', 'ref.name')
  const sheetName = ownDataValue(ref, 'sheetName', 'ref.sheetName')
  const headers = ownDataValue(ref, 'headers', 'ref.headers')
  return {
    kind: 'table',
    id: requiredOwnDataValue(ref, 'id', 'ref.id'),
    label: requiredOwnDataValue(ref, 'label', 'ref.label'),
    ...(name !== undefined ? { name } : {}),
    ...(sheetName !== undefined ? { sheetName } : {}),
    ...(headers !== undefined ? { headers: mapArrayData(headers, 'ref.headers', (header) => header) } : {}),
  }
}

function describeColumnRef(ref: WorkbookColumnRef): WorkbookColumnRefDescription {
  const rows = ownDataValue(ref, 'rows', 'ref.rows')
  return {
    kind: 'column',
    id: requiredOwnDataValue(ref, 'id', 'ref.id'),
    label: requiredOwnDataValue(ref, 'label', 'ref.label'),
    table: describeTableRef(requiredOwnDataValue(ref, 'table', 'ref.table')),
    ...(rows !== undefined ? { rows: describeRowsRef(rows) } : {}),
    name: requiredOwnDataValue(ref, 'name', 'ref.name'),
  }
}

function describeRowsRef(ref: WorkbookRowsRef): WorkbookRowsRefDescription {
  const sheetName = ownDataValue(ref, 'sheetName', 'ref.sheetName')
  const table = ownDataValue(ref, 'table', 'ref.table')
  const where = requiredOwnDataValue(ref, 'where', 'ref.where')
  return {
    kind: 'rows',
    id: requiredOwnDataValue(ref, 'id', 'ref.id'),
    label: requiredOwnDataValue(ref, 'label', 'ref.label'),
    ...(sheetName !== undefined ? { sheetName } : {}),
    ...(table !== undefined ? { table: describeTableRef(table) } : {}),
    where: {
      column: requiredOwnDataValue(where, 'column', 'ref.where.column'),
      op: requiredOwnDataValue(where, 'op', 'ref.where.op'),
      value: requiredOwnDataValue(where, 'value', 'ref.where.value'),
    },
  }
}

export function describeRef(ref: WorkbookRef): WorkbookRefDescription {
  if (hasRefKind(ref, 'range')) {
    return describeRangeRef(ref)
  }
  if (hasRefKind(ref, 'name')) {
    return describeNameRef(ref)
  }
  if (hasRefKind(ref, 'table')) {
    return describeTableRef(ref)
  }
  if (hasRefKind(ref, 'column')) {
    return describeColumnRef(ref)
  }
  if (hasRefKind(ref, 'rows')) {
    return describeRowsRef(ref)
  }
  throw new Error(`Unsupported workbook ref kind: ${String(requiredOwnDataValue(ref, 'kind', 'ref.kind'))}`)
}

function hasCommandKind<K extends WorkbookActionCommand['kind']>(
  command: WorkbookActionCommand,
  kind: K,
): command is Extract<WorkbookActionCommand, { readonly kind: K }> {
  return requiredOwnDataValue(command, 'kind', 'command.kind') === kind
}

function hasExpectationKind<K extends WorkbookCheckExpectation['kind']>(
  expectation: WorkbookCheckExpectation,
  kind: K,
): expectation is Extract<WorkbookCheckExpectation, { readonly kind: K }> {
  return requiredOwnDataValue(expectation, 'kind', 'expectation.kind') === kind
}

function describeCommand(command: WorkbookActionCommand): WorkbookActionCommandDescription {
  if (hasCommandKind(command, 'writeFormula')) {
    const inputs = requiredOwnDataValue(command, 'inputs', 'command.inputs')
    const labels = requiredOwnDataValue(command, 'labels', 'command.labels')
    return {
      kind: 'writeFormula',
      target: describeRef(requiredOwnDataValue(command, 'target', 'command.target')),
      formula: requiredOwnDataValue(command, 'formula', 'command.formula'),
      inputs: mapArrayData(inputs, 'command.inputs', describeRef),
      labels: mapArrayData(labels, 'command.labels', describeFormulaLabel),
    }
  }
  if (hasCommandKind(command, 'writeValue')) {
    return {
      kind: 'writeValue',
      target: describeRef(requiredOwnDataValue(command, 'target', 'command.target')),
      value: requiredOwnDataValue(command, 'value', 'command.value'),
    }
  }
  if (hasCommandKind(command, 'format')) {
    const style = ownDataValue(command, 'style', 'command.style')
    const numberFormat = ownDataValue(command, 'numberFormat', 'command.numberFormat')
    return {
      kind: 'format',
      target: describeRef(requiredOwnDataValue(command, 'target', 'command.target')),
      ...(style !== undefined ? { style: cloneDescriptionData(style, 'command.style') } : {}),
      ...(numberFormat !== undefined ? { numberFormat } : {}),
    }
  }
  if (hasCommandKind(command, 'clear')) {
    return {
      kind: 'clear',
      target: describeRef(requiredOwnDataValue(command, 'target', 'command.target')),
    }
  }
  if (hasCommandKind(command, 'op')) {
    const target = ownDataValue(command, 'target', 'command.target')
    const message = ownDataValue(command, 'message', 'command.message')
    return {
      kind: 'op',
      op: cloneDescriptionData(requiredOwnDataValue(command, 'op', 'command.op'), 'command.op'),
      ...(target !== undefined ? { target: describeRef(target) } : {}),
      ...(message !== undefined ? { message } : {}),
    }
  }
  throw new Error(`Unsupported workbook command kind: ${String(requiredOwnDataValue(command, 'kind', 'command.kind'))}`)
}

function describeFormulaLabel(label: { readonly name: string; readonly ref: WorkbookRef }): WorkbookFormulaLabelDescription {
  return {
    name: requiredOwnDataValue(label, 'name', 'label.name'),
    ref: describeRef(requiredOwnDataValue(label, 'ref', 'label.ref')),
  }
}

function describeChange(change: WorkbookChangeSummary): WorkbookChangeSummaryDescription {
  const target = ownDataValue(change, 'target', 'change.target')
  return {
    kind: requiredOwnDataValue(change, 'kind', 'change.kind'),
    ...(target !== undefined ? { target: describeRef(target) } : {}),
    message: requiredOwnDataValue(change, 'message', 'change.message'),
  }
}

function describeCheck(check: WorkbookCheckResult): WorkbookCheckResultDescription {
  const target = ownDataValue(check, 'target', 'check.target')
  const refs = ownDataValue(check, 'refs', 'check.refs')
  const expectation = ownDataValue(check, 'expectation', 'check.expectation')
  const proof = ownDataValue(check, 'proof', 'check.proof')
  return {
    status: requiredOwnDataValue(check, 'status', 'check.status'),
    kind: requiredOwnDataValue(check, 'kind', 'check.kind'),
    ...(target !== undefined ? { target: describeRef(target) } : {}),
    ...(refs !== undefined ? { refs: mapArrayData(refs, 'check.refs', describeRef) } : {}),
    message: requiredOwnDataValue(check, 'message', 'check.message'),
    ...(expectation !== undefined ? { expectation: describeExpectation(expectation) } : {}),
    ...(proof !== undefined ? { proof: cloneDescriptionData(proof, 'check.proof') } : {}),
  }
}

function describeExpectation(expectation: WorkbookCheckExpectation): WorkbookCheckExpectationDescription {
  if (hasExpectationKind(expectation, 'valueEquals')) {
    return {
      kind: 'valueEquals',
      value: requiredOwnDataValue(expectation, 'value', 'expectation.value'),
    }
  }
  if (hasExpectationKind(expectation, 'formulaEquals')) {
    const inputs = requiredOwnDataValue(expectation, 'inputs', 'expectation.inputs')
    const labels = requiredOwnDataValue(expectation, 'labels', 'expectation.labels')
    return {
      kind: 'formulaEquals',
      formula: requiredOwnDataValue(expectation, 'formula', 'expectation.formula'),
      inputs: mapArrayData(inputs, 'expectation.inputs', describeRef),
      labels: mapArrayData(labels, 'expectation.labels', describeFormulaLabel),
    }
  }
  throw new Error(`Unsupported workbook expectation kind: ${String(requiredOwnDataValue(expectation, 'kind', 'expectation.kind'))}`)
}

export function describePlan<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookActionPlanDescription {
  const input = ownDataValue(plan, 'input', 'plan.input')
  return {
    modelName: requiredOwnDataValue(plan, 'modelName', 'plan.modelName'),
    actionName: requiredOwnDataValue(plan, 'actionName', 'plan.actionName'),
    ...(input !== undefined ? { input: cloneDescriptionData(input, 'plan.input') } : {}),
    refsUsed: mapArrayData(requiredOwnDataValue(plan, 'refsUsed', 'plan.refsUsed'), 'plan.refsUsed', describeRef),
    commands: mapArrayData(requiredOwnDataValue(plan, 'commands', 'plan.commands'), 'plan.commands', describeCommand),
    ops: mapArrayData(requiredOwnDataValue(plan, 'ops', 'plan.ops'), 'plan.ops', (op, _index, entryPath) =>
      cloneDescriptionData(op, entryPath),
    ),
    changed: mapArrayData(requiredOwnDataValue(plan, 'changed', 'plan.changed'), 'plan.changed', describeChange),
    checks: mapArrayData(requiredOwnDataValue(plan, 'checks', 'plan.checks'), 'plan.checks', describeCheck),
  }
}

function describeError(error: WorkbookRunErrorDescription): WorkbookRunErrorDescription {
  const path = ownDataValue(error, 'path', 'error.path')
  const issueCode = ownDataValue(error, 'issueCode', 'error.issueCode')
  return {
    code: requiredOwnDataValue(error, 'code', 'error.code'),
    message: requiredOwnDataValue(error, 'message', 'error.message'),
    ...(path !== undefined ? { path } : {}),
    ...(issueCode !== undefined ? { issueCode } : {}),
  }
}

function describeUndo(undo: WorkbookUndoRef): WorkbookUndoRefDescription {
  const ops = ownDataValue(undo, 'ops', 'undo.ops')
  return {
    id: requiredOwnDataValue(undo, 'id', 'undo.id'),
    ...(ops !== undefined ? { ops: mapArrayData(ops, 'undo.ops', (op, _index, entryPath) => cloneDescriptionData(op, entryPath)) } : {}),
  }
}

function describeApply(apply: WorkbookRunApplySummary): WorkbookRunApplySummaryDescription {
  const planId = ownDataValue(apply, 'planId', 'apply.planId')
  const baseRevision = ownDataValue(apply, 'baseRevision', 'apply.baseRevision')
  const revision = ownDataValue(apply, 'revision', 'apply.revision')
  const previewOps = ownDataValue(apply, 'previewOps', 'apply.previewOps')
  const appliedOps = ownDataValue(apply, 'appliedOps', 'apply.appliedOps')
  const commandReceipts = ownDataValue(apply, 'commandReceipts', 'apply.commandReceipts')
  const proof = ownDataValue(apply, 'proof', 'apply.proof')
  return {
    matched: requiredOwnDataValue(apply, 'matched', 'apply.matched'),
    ...(planId !== undefined ? { planId } : {}),
    ...(baseRevision !== undefined ? { baseRevision } : {}),
    ...(revision !== undefined ? { revision } : {}),
    ...(previewOps !== undefined
      ? { previewOps: mapArrayData(previewOps, 'apply.previewOps', (op, _index, entryPath) => cloneDescriptionData(op, entryPath)) }
      : {}),
    ...(appliedOps !== undefined
      ? { appliedOps: mapArrayData(appliedOps, 'apply.appliedOps', (op, _index, entryPath) => cloneDescriptionData(op, entryPath)) }
      : {}),
    ...(commandReceipts !== undefined
      ? {
          commandReceipts: mapArrayData(commandReceipts, 'apply.commandReceipts', (receipt, _index, entryPath) => {
            const receiptPreviewOps = requiredOwnDataValue(receipt, 'previewOps', `${entryPath}.previewOps`)
            const receiptAppliedOps = requiredOwnDataValue(receipt, 'appliedOps', `${entryPath}.appliedOps`)
            const resolvedRefs = ownDataValue(receipt, 'resolvedRefs', `${entryPath}.resolvedRefs`)
            const receiptProof = ownDataValue(receipt, 'proof', `${entryPath}.proof`)
            return {
              commandIndex: requiredOwnDataValue(receipt, 'commandIndex', `${entryPath}.commandIndex`),
              commandKind: requiredOwnDataValue(receipt, 'commandKind', `${entryPath}.commandKind`),
              commandDigest: requiredOwnDataValue(receipt, 'commandDigest', `${entryPath}.commandDigest`),
              previewOps: mapArrayData(receiptPreviewOps, `${entryPath}.previewOps`, (op, _opIndex, opPath) =>
                cloneDescriptionData(op, opPath),
              ),
              appliedOps: mapArrayData(receiptAppliedOps, `${entryPath}.appliedOps`, (op, _opIndex, opPath) =>
                cloneDescriptionData(op, opPath),
              ),
              ...(resolvedRefs !== undefined ? { resolvedRefs: cloneDescriptionData(resolvedRefs, `${entryPath}.resolvedRefs`) } : {}),
              ...(receiptProof !== undefined ? { proof: cloneDescriptionData(receiptProof, `${entryPath}.proof`) } : {}),
            }
          }),
        }
      : {}),
    ...(proof !== undefined ? { proof: cloneDescriptionData(proof, 'apply.proof') } : {}),
  }
}

function describeUnverified(entry: WorkbookRunUnverified): WorkbookRunUnverifiedDescription {
  return {
    kind: requiredOwnDataValue(entry, 'kind', 'unverified.kind'),
    message: requiredOwnDataValue(entry, 'message', 'unverified.message'),
  }
}

function hasPlanResultStatus<Refs, K extends WorkbookActionPlanResult<Refs>['status']>(
  result: WorkbookActionPlanResult<Refs>,
  status: K,
): result is Extract<WorkbookActionPlanResult<Refs>, { readonly status: K }> {
  return requiredOwnDataValue(result, 'status', 'result.status') === status
}

function hasRunResultStatus<K extends WorkbookRunResult['status']>(
  result: WorkbookRunResult,
  status: K,
): result is Extract<WorkbookRunResult, { readonly status: K }> {
  return requiredOwnDataValue(result, 'status', 'result.status') === status
}

export function describePlanResult<Refs>(result: WorkbookActionPlanResult<Refs>): WorkbookActionPlanResultDescription {
  if (hasPlanResultStatus(result, 'planned')) {
    return {
      status: 'planned',
      plan: describePlan(requiredOwnDataValue(result, 'plan', 'result.plan')),
    }
  }
  const input = ownDataValue(result, 'input', 'result.input')
  return {
    status: 'failed',
    modelName: requiredOwnDataValue(result, 'modelName', 'result.modelName'),
    actionName: requiredOwnDataValue(result, 'actionName', 'result.actionName'),
    ...(input !== undefined ? { input: cloneDescriptionData(input, 'result.input') } : {}),
    errors: mapArrayData(requiredOwnDataValue(result, 'errors', 'result.errors'), 'result.errors', describeError),
    checks: mapArrayData(requiredOwnDataValue(result, 'checks', 'result.checks'), 'result.checks', describeCheck),
  }
}

export function describeRunResult(result: WorkbookRunResult): WorkbookRunResultDescription {
  const apply = ownDataValue(result, 'apply', 'result.apply')
  const undo = ownDataValue(result, 'undo', 'result.undo')
  const unverified = ownDataValue(result, 'unverified', 'result.unverified')
  if (hasRunResultStatus(result, 'done')) {
    return {
      status: 'done',
      ...(apply !== undefined ? { apply: describeApply(apply) } : {}),
      changed: mapArrayData(requiredOwnDataValue(result, 'changed', 'result.changed'), 'result.changed', describeChange),
      checks: mapArrayData(requiredOwnDataValue(result, 'checks', 'result.checks'), 'result.checks', describeCheck),
      ...(undo !== undefined ? { undo: describeUndo(undo) } : {}),
      ...(unverified !== undefined ? { unverified: mapArrayData(unverified, 'result.unverified', describeUnverified) } : {}),
    }
  }
  return {
    status: 'failed',
    errors: mapArrayData(requiredOwnDataValue(result, 'errors', 'result.errors'), 'result.errors', describeError),
    ...(apply !== undefined ? { apply: describeApply(apply) } : {}),
    changed: mapArrayData(requiredOwnDataValue(result, 'changed', 'result.changed'), 'result.changed', describeChange),
    checks: mapArrayData(requiredOwnDataValue(result, 'checks', 'result.checks'), 'result.checks', describeCheck),
    ...(undo !== undefined ? { undo: describeUndo(undo) } : {}),
    ...(unverified !== undefined ? { unverified: mapArrayData(unverified, 'result.unverified', describeUnverified) } : {}),
  }
}
