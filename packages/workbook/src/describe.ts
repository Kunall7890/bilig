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
  type WorkbookActionCommand,
  type WorkbookActionMap,
  type WorkbookActionPlan,
  type WorkbookActionPlanResult,
  type WorkbookModel,
} from './model.js'
import type { WorkbookOp } from './ops.js'
import type { WorkbookChangeSummary, WorkbookCheckResult, WorkbookCheckStatus } from './result.js'

export interface WorkbookModelDescription {
  readonly name: string
  readonly actions: readonly string[]
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
      readonly numberFormat?: string
    }
  | {
      readonly kind: 'clear'
      readonly target: WorkbookRefDescription
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
}

export interface WorkbookActionPlanDescription {
  readonly modelName: string
  readonly actionName: string
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
      readonly errors: readonly WorkbookRunErrorDescription[]
      readonly checks: readonly WorkbookCheckResultDescription[]
    }

export interface WorkbookRunErrorDescription {
  readonly code: string
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
  }
}

export function describePlan<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookActionPlanDescription {
  return {
    modelName: plan.modelName,
    actionName: plan.actionName,
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
    errors: result.errors.map(describeError),
    checks: result.checks.map(describeCheck),
  }
}
