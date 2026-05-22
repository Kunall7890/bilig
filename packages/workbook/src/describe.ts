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
import type { WorkbookReceiptProof, WorkbookRenderedReceipt, WorkbookRunReceipt } from './receipt.js'
import type { WorkbookRuntimePreview, WorkbookRuntimeRequirement } from './requirements.js'
import type {
  WorkbookAppliedSummary,
  WorkbookChangeSummary,
  WorkbookCheckExpectation,
  WorkbookCheckProof,
  WorkbookCheckResult,
  WorkbookCheckStatus,
  WorkbookRunError,
  WorkbookRunErrorCode,
  WorkbookRunResult,
  WorkbookUndoRef,
} from './result.js'
import type { WorkbookPreviewResult } from './run.js'

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
  readonly table: WorkbookTableRefDescription
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
  readonly proof?: WorkbookCheckProofDescription
}

export type WorkbookCheckExpectationDescription =
  | {
      readonly kind: 'valueEquals'
      readonly value: LiteralInput
    }
  | {
      readonly kind: 'valuesEqual'
      readonly values: readonly (readonly LiteralInput[])[]
    }
  | {
      readonly kind: 'formulaEquals'
      readonly formula: string
      readonly inputs: readonly WorkbookRefDescription[]
    }
  | {
      readonly kind: 'formulasEqual'
      readonly formulas: readonly (readonly (string | null)[])[]
    }

export type WorkbookCheckProofDescription = WorkbookCheckProof

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

export interface WorkbookRuntimeRequirementDescription extends WorkbookRuntimeRequirement {}

export interface WorkbookRuntimePreviewDescription {
  readonly modelName: string
  readonly actionName: string
  readonly requirements: readonly WorkbookRuntimeRequirementDescription[]
  readonly materializedOps: readonly WorkbookOp[]
}

export type WorkbookPreviewResultDescription =
  | {
      readonly status: 'previewed'
      readonly preview: WorkbookRuntimePreviewDescription
    }
  | {
      readonly status: 'failed'
      readonly errors: readonly WorkbookRunErrorDescription[]
      readonly checks: readonly WorkbookCheckResultDescription[]
    }

export interface WorkbookUndoRefDescription {
  readonly id: string
  readonly ops?: readonly WorkbookOp[]
}

export type WorkbookRunResultDescription =
  | {
      readonly status: 'done'
      readonly changed: readonly WorkbookChangeSummaryDescription[]
      readonly checks: readonly WorkbookCheckResultDescription[]
      readonly undo?: WorkbookUndoRefDescription
      readonly applied?: WorkbookAppliedSummaryDescription
      readonly receipt?: WorkbookRunReceiptDescription
    }
  | {
      readonly status: 'failed'
      readonly errors: readonly WorkbookRunErrorDescription[]
      readonly checks: readonly WorkbookCheckResultDescription[]
      readonly undo?: WorkbookUndoRefDescription
      readonly receipt?: WorkbookRunReceiptDescription
    }

export interface WorkbookRunErrorDescription {
  readonly code: WorkbookRunErrorCode
  readonly message: string
  readonly path?: string
  readonly target?: WorkbookRefDescription
  readonly check?: WorkbookCheckResultDescription
  readonly expected?: WorkbookActionInput
  readonly actual?: WorkbookActionInput
}

export interface WorkbookAppliedSummaryDescription {
  readonly opCount: number
  readonly ops?: readonly WorkbookOp[]
}

export interface WorkbookReceiptProofDescription extends Omit<WorkbookReceiptProof, 'target'> {
  readonly target?: WorkbookRefDescription
}

export interface WorkbookRenderedReceiptDescription extends Omit<WorkbookRenderedReceipt, 'diffs'> {
  readonly diffs?: readonly WorkbookChangeSummaryDescription[]
}

export interface WorkbookRunReceiptDescription extends Omit<WorkbookRunReceipt, 'rendered' | 'proof' | 'undo'> {
  readonly rendered?: WorkbookRenderedReceiptDescription
  readonly proof: readonly WorkbookReceiptProofDescription[]
  readonly undo?: WorkbookUndoRefDescription
}

export function describeModel<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
): WorkbookModelDescription {
  return inspectModel(model)
}

function describeRangeRef(ref: WorkbookRangeRef): WorkbookRangeRefDescription {
  return {
    kind: 'range',
    id: ref.id,
    label: ref.label,
    range: { ...ref.range },
  }
}

function describeNameRef(ref: WorkbookNameRef): WorkbookNameRefDescription {
  return {
    kind: 'name',
    id: ref.id,
    label: ref.label,
    name: ref.name,
  }
}

function describeTableRef(ref: WorkbookTableRef): WorkbookTableRefDescription {
  return {
    kind: 'table',
    id: ref.id,
    label: ref.label,
    ...(ref.name !== undefined ? { name: ref.name } : {}),
    ...(ref.sheetName !== undefined ? { sheetName: ref.sheetName } : {}),
    ...(ref.headers !== undefined ? { headers: [...ref.headers] } : {}),
  }
}

function describeColumnRef(ref: WorkbookColumnRef): WorkbookColumnRefDescription {
  return {
    kind: 'column',
    id: ref.id,
    label: ref.label,
    table: describeTableRef(ref.table),
    ...(ref.rows !== undefined ? { rows: describeRowsRef(ref.rows) } : {}),
    name: ref.name,
  }
}

function describeRowsRef(ref: WorkbookRowsRef): WorkbookRowsRefDescription {
  return {
    kind: 'rows',
    id: ref.id,
    label: ref.label,
    table: describeTableRef(ref.table),
    where: { ...ref.where },
  }
}

export function describeRef(ref: WorkbookRef): WorkbookRefDescription {
  switch (ref.kind) {
    case 'range':
      return describeRangeRef(ref)
    case 'name':
      return describeNameRef(ref)
    case 'table':
      return describeTableRef(ref)
    case 'column':
      return describeColumnRef(ref)
    case 'rows':
      return describeRowsRef(ref)
  }
}

function describeCommand(command: WorkbookActionCommand): WorkbookActionCommandDescription {
  switch (command.kind) {
    case 'writeFormula':
      return {
        kind: 'writeFormula',
        target: describeRef(command.target),
        formula: command.formula,
        inputs: command.inputs.map(describeRef),
      }
    case 'writeValue':
      return {
        kind: 'writeValue',
        target: describeRef(command.target),
        value: command.value,
      }
    case 'format':
      return {
        kind: 'format',
        target: describeRef(command.target),
        ...(command.style !== undefined ? { style: command.style } : {}),
        ...(command.numberFormat !== undefined ? { numberFormat: command.numberFormat } : {}),
      }
    case 'clear':
      return {
        kind: 'clear',
        target: describeRef(command.target),
      }
    case 'op':
      return {
        kind: 'op',
        op: command.op,
        ...(command.target !== undefined ? { target: describeRef(command.target) } : {}),
        ...(command.message !== undefined ? { message: command.message } : {}),
      }
  }
}

function describeChange(change: WorkbookChangeSummary): WorkbookChangeSummaryDescription {
  return {
    kind: change.kind,
    ...(change.target !== undefined ? { target: describeRef(change.target) } : {}),
    message: change.message,
  }
}

function describeCheck(check: WorkbookCheckResult): WorkbookCheckResultDescription {
  return {
    status: check.status,
    kind: check.kind,
    ...(check.target !== undefined ? { target: describeRef(check.target) } : {}),
    ...(check.refs !== undefined ? { refs: check.refs.map(describeRef) } : {}),
    message: check.message,
    ...(check.expectation !== undefined ? { expectation: describeExpectation(check.expectation) } : {}),
    ...(check.proof !== undefined ? { proof: describeProof(check.proof) } : {}),
  }
}

function describeProof(proof: WorkbookCheckProof): WorkbookCheckProofDescription {
  if (proof.kind === 'values') {
    return {
      kind: 'values',
      values: proof.values.map((row) => [...row]),
    }
  }
  if (proof.kind === 'formulas') {
    return {
      kind: 'formulas',
      formulas: proof.formulas.map((row) => [...row]),
    }
  }
  return { ...proof }
}

function describeExpectation(expectation: WorkbookCheckExpectation): WorkbookCheckExpectationDescription {
  switch (expectation.kind) {
    case 'valueEquals':
      return {
        kind: 'valueEquals',
        value: expectation.value,
      }
    case 'valuesEqual':
      return {
        kind: 'valuesEqual',
        values: expectation.values.map((row) => [...row]),
      }
    case 'formulaEquals':
      return {
        kind: 'formulaEquals',
        formula: expectation.formula,
        inputs: expectation.inputs.map(describeRef),
      }
    case 'formulasEqual':
      return {
        kind: 'formulasEqual',
        formulas: expectation.formulas.map((row) => [...row]),
      }
  }
}

export function describePlan<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookActionPlanDescription {
  return {
    modelName: plan.modelName,
    actionName: plan.actionName,
    ...(plan.input !== undefined ? { input: plan.input } : {}),
    refsUsed: plan.refsUsed.map(describeRef),
    commands: plan.commands.map(describeCommand),
    ops: [...plan.ops],
    changed: plan.changed.map(describeChange),
    checks: plan.checks.map(describeCheck),
  }
}

function describeError(error: WorkbookRunError): WorkbookRunErrorDescription {
  return {
    code: error.code,
    message: error.message,
    ...(error.path !== undefined ? { path: error.path } : {}),
    ...(error.target !== undefined ? { target: describeRef(error.target) } : {}),
    ...(error.check !== undefined ? { check: describeCheck(error.check) } : {}),
    ...(error.expected !== undefined ? { expected: error.expected } : {}),
    ...(error.actual !== undefined ? { actual: error.actual } : {}),
  }
}

function describeApplied(applied: WorkbookAppliedSummary): WorkbookAppliedSummaryDescription {
  return {
    opCount: applied.opCount,
    ...(applied.ops !== undefined ? { ops: [...applied.ops] } : {}),
  }
}

function describeReceiptProof(proof: WorkbookReceiptProof): WorkbookReceiptProofDescription {
  return {
    kind: proof.kind,
    status: proof.status,
    message: proof.message,
    ...(proof.revision !== undefined ? { revision: proof.revision } : {}),
    ...(proof.target !== undefined ? { target: describeRef(proof.target) } : {}),
    ...(proof.data !== undefined ? { data: proof.data } : {}),
  }
}

function describeRenderedReceipt(rendered: WorkbookRenderedReceipt): WorkbookRenderedReceiptDescription {
  return {
    ...(rendered.revision !== undefined ? { revision: rendered.revision } : {}),
    ...(rendered.diffs !== undefined ? { diffs: rendered.diffs.map(describeChange) } : {}),
    ...(rendered.message !== undefined ? { message: rendered.message } : {}),
  }
}

function describeRunReceipt(receipt: WorkbookRunReceipt): WorkbookRunReceiptDescription {
  return {
    ...(receipt.commandId !== undefined ? { commandId: receipt.commandId } : {}),
    ...(receipt.idempotencyKey !== undefined ? { idempotencyKey: receipt.idempotencyKey } : {}),
    modelName: receipt.modelName,
    actionName: receipt.actionName,
    ...(receipt.baseRevision !== undefined ? { baseRevision: receipt.baseRevision } : {}),
    ...(receipt.appliedRevision !== undefined ? { appliedRevision: receipt.appliedRevision } : {}),
    ...(receipt.calculatedRevision !== undefined ? { calculatedRevision: receipt.calculatedRevision } : {}),
    ...(receipt.renderedRevision !== undefined ? { renderedRevision: receipt.renderedRevision } : {}),
    ...(receipt.rendered !== undefined ? { rendered: describeRenderedReceipt(receipt.rendered) } : {}),
    previewed: receipt.previewed,
    applied: receipt.applied,
    verified: receipt.verified,
    checkCount: receipt.checkCount,
    passedCheckCount: receipt.passedCheckCount,
    failedCheckCount: receipt.failedCheckCount,
    unverifiedCheckCount: receipt.unverifiedCheckCount,
    proof: receipt.proof.map(describeReceiptProof),
    ...(receipt.warnings !== undefined ? { warnings: [...receipt.warnings] } : {}),
    ...(receipt.undo !== undefined ? { undo: describeUndo(receipt.undo) } : {}),
  }
}

function describeUndo(undo: WorkbookUndoRef): WorkbookUndoRefDescription {
  return {
    id: undo.id,
    ...(undo.ops !== undefined ? { ops: [...undo.ops] } : {}),
  }
}

export function describePlanResult<Refs>(result: WorkbookActionPlanResult<Refs>): WorkbookActionPlanResultDescription {
  if (result.status === 'planned') {
    return {
      status: 'planned',
      plan: describePlan(result.plan),
    }
  }
  return {
    status: 'failed',
    modelName: result.modelName,
    actionName: result.actionName,
    ...(result.input !== undefined ? { input: result.input } : {}),
    errors: result.errors.map(describeError),
    checks: result.checks.map(describeCheck),
  }
}

function describeRuntimeRequirement(requirement: WorkbookRuntimeRequirement): WorkbookRuntimeRequirementDescription {
  return structuredClone(requirement)
}

function describeRuntimePreview(preview: WorkbookRuntimePreview): WorkbookRuntimePreviewDescription {
  return {
    modelName: preview.modelName,
    actionName: preview.actionName,
    requirements: preview.requirements.map(describeRuntimeRequirement),
    materializedOps: preview.materializedOps.map((op) => structuredClone(op)),
  }
}

export function describePreviewResult(result: WorkbookPreviewResult): WorkbookPreviewResultDescription {
  if (result.status === 'previewed') {
    return {
      status: 'previewed',
      preview: describeRuntimePreview(result.preview),
    }
  }
  return {
    status: 'failed',
    errors: result.errors.map(describeError),
    checks: result.checks.map(describeCheck),
  }
}

export function describeRunResult(result: WorkbookRunResult): WorkbookRunResultDescription {
  if (result.status === 'done') {
    return {
      status: 'done',
      changed: result.changed.map(describeChange),
      checks: result.checks.map(describeCheck),
      ...(result.undo !== undefined ? { undo: describeUndo(result.undo) } : {}),
      ...(result.applied !== undefined ? { applied: describeApplied(result.applied) } : {}),
      ...(result.receipt !== undefined ? { receipt: describeRunReceipt(result.receipt) } : {}),
    }
  }
  return {
    status: 'failed',
    errors: result.errors.map(describeError),
    checks: result.checks.map(describeCheck),
    ...(result.undo !== undefined ? { undo: describeUndo(result.undo) } : {}),
    ...(result.receipt !== undefined ? { receipt: describeRunReceipt(result.receipt) } : {}),
  }
}
