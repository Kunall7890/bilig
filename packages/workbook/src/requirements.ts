import { describeRef, type WorkbookCheckExpectationDescription, type WorkbookRefDescription } from './describe.js'
import type { WorkbookRef } from './find.js'
import type { WorkbookActionCommand, WorkbookActionPlan } from './model.js'
import type { EngineOp, WorkbookOp } from './ops.js'
import type { WorkbookCheckExpectation, WorkbookCheckResult } from './result.js'

export type WorkbookRuntimeRequirementKind = 'apply' | 'read' | 'verify'

export type WorkbookRuntimeCapability = 'writeFormula' | 'writeValue' | 'format' | 'clear' | 'applyOp' | 'read' | 'verifyCheck'

export const workbookRuntimeCapabilities = Object.freeze([
  'writeFormula',
  'writeValue',
  'format',
  'clear',
  'applyOp',
  'read',
  'verifyCheck',
] satisfies readonly WorkbookRuntimeCapability[])

export type WorkbookRuntimeMaterialization = 'concreteOp' | 'adapterMaterialization' | 'providedOp'

export interface WorkbookRuntimeRequirement {
  readonly kind: WorkbookRuntimeRequirementKind
  readonly capability: WorkbookRuntimeCapability
  readonly path: string
  readonly message: string
  readonly materialization?: WorkbookRuntimeMaterialization
  readonly commandIndex?: number
  readonly checkIndex?: number
  readonly opIndex?: number
  readonly opIndexes?: readonly number[]
  readonly opKind?: string
  readonly checkKind?: string
  readonly target?: WorkbookRefDescription
  readonly refs?: readonly WorkbookRefDescription[]
  readonly expectation?: WorkbookCheckExpectationDescription
}

export interface WorkbookRuntimeRequirements {
  readonly modelName: string
  readonly actionName: string
  readonly requirements: readonly WorkbookRuntimeRequirement[]
}

export interface WorkbookRuntimePreview {
  readonly modelName: string
  readonly actionName: string
  readonly requirements: readonly WorkbookRuntimeRequirement[]
  readonly materializedOps: readonly EngineOp[]
}

export interface WorkbookRuntimeCapabilityIssue {
  readonly capability: WorkbookRuntimeCapability
  readonly path: string
  readonly message: string
  readonly requirement: WorkbookRuntimeRequirement
}

export interface WorkbookRuntimeCapabilityVerification {
  readonly status: 'supported' | 'unsupported'
  readonly missing: readonly WorkbookRuntimeCapabilityIssue[]
}

export function isWorkbookRuntimeCapability(value: unknown): value is WorkbookRuntimeCapability {
  return typeof value === 'string' && workbookRuntimeCapabilities.some((capability) => capability === value)
}

export function verifyRuntimeRequirements(
  requirements: WorkbookRuntimeRequirements,
  capabilities: readonly WorkbookRuntimeCapability[],
): WorkbookRuntimeCapabilityVerification {
  const supported = new Set(capabilities)
  const missing = requirements.requirements.flatMap((requirement): WorkbookRuntimeCapabilityIssue[] => {
    if (supported.has(requirement.capability)) {
      return []
    }
    return [
      {
        capability: requirement.capability,
        path: requirement.path,
        message: `Runtime is missing ${requirement.capability} for ${requirement.path}: ${requirement.message}`,
        requirement,
      },
    ]
  })
  return {
    status: missing.length === 0 ? 'supported' : 'unsupported',
    missing,
  }
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

function expectationDescription(expectation: WorkbookCheckExpectation): WorkbookCheckExpectationDescription {
  if (expectation.kind === 'valueEquals') {
    return {
      kind: 'valueEquals',
      value: expectation.value,
    }
  }
  if (expectation.kind === 'valuesEqual') {
    return {
      kind: 'valuesEqual',
      values: expectation.values.map((row) => [...row]),
    }
  }
  if (expectation.kind === 'formulaEquals') {
    return {
      kind: 'formulaEquals',
      formula: expectation.formula,
      inputs: expectation.inputs.map(describeRef),
    }
  }
  return {
    kind: 'formulasEqual',
    formulas: expectation.formulas.map((row) => [...row]),
  }
}

function concreteSingleCell(target: WorkbookRef): { sheetName: string; address: string } | null {
  if (target.kind !== 'range') {
    return null
  }
  const range = target.range
  return range.startAddress === range.endAddress ? { sheetName: range.sheetName, address: range.startAddress } : null
}

function expectedConcreteOp(command: WorkbookActionCommand): WorkbookOp | null {
  if (command.kind === 'op') {
    return command.op
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

function opIndexesFor(command: WorkbookActionCommand, ops: readonly WorkbookOp[]): readonly number[] {
  const expected = expectedConcreteOp(command)
  if (expected === null) {
    return []
  }
  const expectedKey = opKey(expected)
  return ops.flatMap((op, index) => (opKey(op) === expectedKey ? [index] : []))
}

function commandMaterialization(command: WorkbookActionCommand, opIndexes: readonly number[]): WorkbookRuntimeMaterialization {
  if (command.kind === 'op') {
    return 'providedOp'
  }
  return opIndexes.length === 0 ? 'adapterMaterialization' : 'concreteOp'
}

function commandRequirement(command: WorkbookActionCommand, commandIndex: number, ops: readonly WorkbookOp[]): WorkbookRuntimeRequirement {
  const opIndexes = opIndexesFor(command, ops)
  return {
    kind: 'apply',
    capability: commandCapability(command),
    path: `commands[${String(commandIndex)}]`,
    commandIndex,
    materialization: commandMaterialization(command, opIndexes),
    ...(opIndexes.length > 0 ? { opIndexes } : {}),
    ...(command.kind === 'op' ? { opKind: command.op.kind } : {}),
    ...describedRef(command.target),
    ...describedRefs(commandRefs(command)),
    message: commandMessage(command),
  }
}

function readRequirement(check: WorkbookCheckResult, checkIndex: number): WorkbookRuntimeRequirement | null {
  if (check.expectation === undefined) {
    return null
  }
  return {
    kind: 'read',
    capability: 'read',
    path: `checks[${String(checkIndex)}]`,
    checkIndex,
    checkKind: check.kind,
    ...describedRef(check.target),
    ...describedRefs(check.expectation.kind === 'formulaEquals' ? check.expectation.inputs : undefined),
    expectation: expectationDescription(check.expectation),
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
    path: `checks[${String(checkIndex)}]`,
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

export function describeRuntimeRequirements<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRuntimeRequirements {
  const requirements: WorkbookRuntimeRequirement[] = plan.commands.map((command, commandIndex) =>
    commandRequirement(command, commandIndex, plan.ops),
  )
  const commandBackedOpIndexes = new Set(plan.commands.flatMap((command) => opIndexesFor(command, plan.ops)))

  plan.ops.forEach((op, opIndex) => {
    if (commandBackedOpIndexes.has(opIndex)) {
      return
    }
    requirements.push({
      kind: 'apply',
      capability: 'applyOp',
      path: `ops[${String(opIndex)}]`,
      opIndex,
      materialization: 'providedOp',
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
