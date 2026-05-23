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
  readonly previewOps?: readonly WorkbookOp[]
  readonly appliedOps?: readonly WorkbookOp[]
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
    ...(ref.sheetName !== undefined ? { sheetName: ref.sheetName } : {}),
    ...(ref.table !== undefined ? { table: describeTableRef(ref.table) } : {}),
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
        labels: command.labels.map(describeFormulaLabel),
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

function describeFormulaLabel(label: { readonly name: string; readonly ref: WorkbookRef }): WorkbookFormulaLabelDescription {
  return {
    name: label.name,
    ref: describeRef(label.ref),
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
    ...(check.proof !== undefined ? { proof: check.proof } : {}),
  }
}

function describeExpectation(expectation: WorkbookCheckExpectation): WorkbookCheckExpectationDescription {
  switch (expectation.kind) {
    case 'valueEquals':
      return {
        kind: 'valueEquals',
        value: expectation.value,
      }
    case 'formulaEquals':
      return {
        kind: 'formulaEquals',
        formula: expectation.formula,
        inputs: expectation.inputs.map(describeRef),
        labels: expectation.labels.map(describeFormulaLabel),
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

function describeError(error: WorkbookRunErrorDescription): WorkbookRunErrorDescription {
  return {
    code: error.code,
    message: error.message,
  }
}

function describeUndo(undo: WorkbookUndoRef): WorkbookUndoRefDescription {
  return {
    id: undo.id,
    ...(undo.ops !== undefined ? { ops: [...undo.ops] } : {}),
  }
}

function describeApply(apply: WorkbookRunApplySummary): WorkbookRunApplySummaryDescription {
  return {
    matched: apply.matched,
    ...(apply.previewOps !== undefined ? { previewOps: [...apply.previewOps] } : {}),
    ...(apply.appliedOps !== undefined ? { appliedOps: [...apply.appliedOps] } : {}),
    ...(apply.proof !== undefined ? { proof: apply.proof } : {}),
  }
}

function describeUnverified(entry: WorkbookRunUnverified): WorkbookRunUnverifiedDescription {
  return {
    kind: entry.kind,
    message: entry.message,
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

export function describeRunResult(result: WorkbookRunResult): WorkbookRunResultDescription {
  if (result.status === 'done') {
    return {
      status: 'done',
      ...(result.apply !== undefined ? { apply: describeApply(result.apply) } : {}),
      changed: result.changed.map(describeChange),
      checks: result.checks.map(describeCheck),
      ...(result.undo !== undefined ? { undo: describeUndo(result.undo) } : {}),
      ...(result.unverified !== undefined ? { unverified: result.unverified.map(describeUnverified) } : {}),
    }
  }
  return {
    status: 'failed',
    errors: result.errors.map(describeError),
    ...(result.apply !== undefined ? { apply: describeApply(result.apply) } : {}),
    changed: result.changed.map(describeChange),
    checks: result.checks.map(describeCheck),
    ...(result.undo !== undefined ? { undo: describeUndo(result.undo) } : {}),
    ...(result.unverified !== undefined ? { unverified: result.unverified.map(describeUnverified) } : {}),
  }
}
