import { describeRef, type WorkbookRefDescription } from './describe.js'
import type { WorkbookRef } from './find.js'
import type { WorkbookActionCommand, WorkbookActionPlan } from './model.js'
import type { WorkbookOp } from './ops.js'
import { hydratePlanData, isHydratedPlan, type WorkbookExecutablePlan } from './plan-data.js'
import type { WorkbookCheckResult } from './result.js'

type WorkbookConcreteCommandOp = Extract<WorkbookOp, { kind: 'setCellFormula' | 'setCellValue' | 'setCellFormat' | 'clearCell' }>

export type WorkbookRuntimeRequirementKind = 'apply' | 'read' | 'verify'

export type WorkbookRuntimeCapability = 'writeFormula' | 'writeValue' | 'format' | 'clear' | 'applyOp' | 'read' | 'verifyCheck'

export interface WorkbookRuntimeRequirement {
  readonly kind: WorkbookRuntimeRequirementKind
  readonly capability: WorkbookRuntimeCapability
  readonly message: string
  readonly commandIndex?: number
  readonly checkIndex?: number
  readonly opIndex?: number
  readonly opKind?: string
  readonly checkKind?: string
  readonly target?: WorkbookRefDescription
  readonly refs?: readonly WorkbookRefDescription[]
}

export interface WorkbookRuntimeRequirements {
  readonly modelName: string
  readonly actionName: string
  readonly requirements: readonly WorkbookRuntimeRequirement[]
}

function describedRef(ref: WorkbookRef | undefined): { readonly target: WorkbookRefDescription } | {} {
  return ref === undefined ? {} : { target: describeRef(ref) }
}

function describedRefs(refs: readonly WorkbookRef[] | undefined): { readonly refs: readonly WorkbookRefDescription[] } | {} {
  return refs === undefined || refs.length === 0 ? {} : { refs: refs.map(describeRef) }
}

function commandCapability(command: WorkbookActionCommand): WorkbookRuntimeCapability {
  switch (command.kind) {
    case 'writeFormula':
      return 'writeFormula'
    case 'writeValue':
      return 'writeValue'
    case 'format':
      return 'format'
    case 'clear':
      return 'clear'
    case 'op':
      return 'applyOp'
  }
}

function commandRefs(command: WorkbookActionCommand): readonly WorkbookRef[] | undefined {
  return command.kind === 'writeFormula' ? command.inputs : undefined
}

function commandMessage(command: WorkbookActionCommand): string {
  switch (command.kind) {
    case 'writeFormula':
      return `Apply formula write to ${command.target.label}`
    case 'writeValue':
      return `Apply value write to ${command.target.label}`
    case 'format':
      return `Apply format to ${command.target.label}`
    case 'clear':
      return `Apply clear to ${command.target.label}`
    case 'op':
      return command.target === undefined
        ? `Apply workbook op ${command.op.kind}`
        : `Apply workbook op ${command.op.kind} to ${command.target.label}`
  }
}

function commandRequirement(command: WorkbookActionCommand, commandIndex: number): WorkbookRuntimeRequirement {
  return {
    kind: 'apply',
    capability: commandCapability(command),
    commandIndex,
    ...(command.kind === 'op' ? { opKind: command.op.kind } : {}),
    ...describedRef(command.target),
    ...describedRefs(commandRefs(command)),
    message: commandMessage(command),
  }
}

function concreteSingleCell(target: WorkbookRef): { readonly sheetName: string; readonly address: string } | null {
  if (target.kind !== 'range') {
    return null
  }
  const range = target.range
  return range.startAddress === range.endAddress ? { sheetName: range.sheetName, address: range.startAddress } : null
}

function commandConcreteOp(command: WorkbookActionCommand): WorkbookConcreteCommandOp | null {
  if (command.kind === 'op') {
    return command.op.kind === 'setCellFormula' ||
      command.op.kind === 'setCellValue' ||
      command.op.kind === 'setCellFormat' ||
      command.op.kind === 'clearCell'
      ? command.op
      : null
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
    case 'clear':
      return {
        kind: 'clearCell',
        sheetName: target.sheetName,
        address: target.address,
      }
  }
}

function readRequirement(check: WorkbookCheckResult, checkIndex: number): WorkbookRuntimeRequirement | null {
  if (check.expectation === undefined) {
    return null
  }
  return {
    kind: 'read',
    capability: 'read',
    checkIndex,
    checkKind: check.kind,
    ...describedRef(check.target),
    ...describedRefs(check.expectation.kind === 'formulaEquals' ? check.expectation.inputs : undefined),
    message: `Read ${check.target?.label ?? check.kind} for ${check.expectation.kind}`,
  }
}

function verifyRequirement(check: WorkbookCheckResult, checkIndex: number): WorkbookRuntimeRequirement | null {
  if (check.expectation !== undefined) {
    return null
  }
  return {
    kind: 'verify',
    capability: 'verifyCheck',
    checkIndex,
    checkKind: check.kind,
    ...describedRef(check.target),
    ...describedRefs(check.refs),
    message: `Verify ${check.kind} for ${check.target?.label ?? check.kind}`,
  }
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

function opKey(op: WorkbookOp): string {
  return JSON.stringify(canonicalValue(op))
}

function commandCoveredOpKeys(commands: readonly WorkbookActionCommand[]): ReadonlySet<string> {
  const keys = new Set<string>()
  commands.forEach((command) => {
    if (command.kind === 'op') {
      keys.add(opKey(command.op))
      return
    }
    const concreteOp = commandConcreteOp(command)
    if (concreteOp !== null) {
      keys.add(opKey(concreteOp))
    }
  })
  return keys
}

function describeLiveRuntimeRequirements<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRuntimeRequirements {
  const requirements: WorkbookRuntimeRequirement[] = plan.commands.map(commandRequirement)
  const commandCoveredOps = commandCoveredOpKeys(plan.commands)

  plan.ops.forEach((op, opIndex) => {
    if (commandCoveredOps.has(opKey(op))) {
      return
    }
    requirements.push({
      kind: 'apply',
      capability: 'applyOp',
      opIndex,
      opKind: op.kind,
      message: `Apply workbook op ${op.kind}`,
    })
  })

  plan.checks.forEach((check, checkIndex) => {
    const read = readRequirement(check, checkIndex)
    if (read !== null) {
      requirements.push(read)
    }
  })

  plan.checks.forEach((check, checkIndex) => {
    const verify = verifyRequirement(check, checkIndex)
    if (verify !== null) {
      requirements.push(verify)
    }
  })

  return {
    modelName: plan.modelName,
    actionName: plan.actionName,
    requirements,
  }
}

export function describeRuntimeRequirements<Refs>(plan: WorkbookExecutablePlan<Refs>): WorkbookRuntimeRequirements {
  if (isHydratedPlan(plan)) {
    return describeLiveRuntimeRequirements(plan)
  }
  return describeLiveRuntimeRequirements(hydratePlanData(plan))
}
