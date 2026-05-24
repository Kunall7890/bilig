import { isLiteralInput, type LiteralInput } from '@bilig/protocol'
import { isWorkbookRef, type WorkbookRef } from './find.js'
import { formula, type WorkbookFormulaOperand } from './formula.js'
import { isObject, optionalDataProperty, requiredDataProperty } from './data-properties.js'
import type { WorkbookCheckExpectation, WorkbookCheckResult } from './result.js'

export interface WorkbookCustomCheckOptions {
  readonly kind: string
  readonly message: string
  readonly target?: WorkbookRef
  readonly refs?: readonly WorkbookRef[]
}

export interface WorkbookReadbackCheckOptions {
  readonly message?: string
}

export interface WorkbookCheckApi {
  readonly exists: (target: WorkbookRef) => WorkbookCheckResult
  readonly noFormulaErrors: (target: WorkbookRef) => WorkbookCheckResult
  readonly valueEquals: (target: WorkbookRef, value: LiteralInput, options?: WorkbookReadbackCheckOptions) => WorkbookCheckResult
  readonly formulaEquals: (
    target: WorkbookRef,
    value: WorkbookFormulaOperand,
    options?: WorkbookReadbackCheckOptions,
  ) => WorkbookCheckResult
  readonly custom: (options: WorkbookCustomCheckOptions) => WorkbookCheckResult
}

export type WorkbookBuiltInCheckKind = 'exists' | 'noFormulaErrors' | 'valueEquals' | 'formulaEquals'

export const builtInWorkbookCheckKinds = Object.freeze([
  'exists',
  'noFormulaErrors',
  'valueEquals',
  'formulaEquals',
] satisfies readonly WorkbookBuiltInCheckKind[])

export function isBuiltInWorkbookCheckKind(value: unknown): value is WorkbookBuiltInCheckKind {
  return typeof value === 'string' && builtInWorkbookCheckKinds.some((kind) => kind === value)
}

export function createWorkbookCheckResult(kind: string, target: WorkbookRef, message: string): WorkbookCheckResult {
  return createWorkbookCheck({ kind, target, message })
}

interface WorkbookCheckBuildOptions extends WorkbookCustomCheckOptions {
  readonly expectation?: WorkbookCheckExpectation
}

function requiredText(value: unknown, name: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Workbook check ${name} must be a string`)
  }
  const trimmed = value.trim()
  if (trimmed === '') {
    throw new Error(`Workbook check ${name} cannot be empty`)
  }
  return trimmed
}

function checkedLiteralInput(value: LiteralInput): LiteralInput {
  if (!isLiteralInput(value)) {
    throw new Error('Workbook readback value must be a finite JSON literal')
  }
  return value
}

function checkTarget(kind: string, target: unknown): WorkbookRef {
  if (!isWorkbookRef(target)) {
    throw new Error(`Workbook check ${kind} target must be a workbook ref`)
  }
  return target
}

function dataArrayValues(value: unknown, label: string): readonly unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be an array`)
  }
  const entries: unknown[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      throw new Error(`${label}[${index}] must be a data property`)
    }
    entries.push(descriptor.value)
  }
  return entries
}

function checkRefs(value: unknown, label: string): readonly WorkbookRef[] {
  return Object.freeze(
    dataArrayValues(value, label).map((entry, index) => {
      if (!isWorkbookRef(entry)) {
        throw new Error(`${label}[${index}] must be a workbook ref`)
      }
      return entry
    }),
  )
}

function refKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
}

function uniqueRefs(refs: readonly WorkbookRef[] | undefined): readonly WorkbookRef[] | undefined {
  if (refs === undefined) {
    return undefined
  }
  const seen = new Set<string>()
  const unique: WorkbookRef[] = []
  for (const ref of refs) {
    const key = refKey(ref)
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    unique.push(ref)
  }
  return unique.length === 0 ? undefined : unique
}

function assertOnlyDataProperties(value: unknown, path: string, seen = new WeakSet<object>()): void {
  if (typeof value !== 'object' || value === null) {
    return
  }
  if (seen.has(value)) {
    return
  }
  seen.add(value)
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key === 'symbol') {
      throw new Error(`${path}[${String(key)}] must be a data property`)
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    const child = Array.isArray(value) && /^\d+$/.test(key) ? `${path}[${key}]` : `${path}.${key}`
    if (descriptor === undefined || !('value' in descriptor)) {
      throw new Error(`${child} must be a data property`)
    }
    assertOnlyDataProperties(descriptor.value, child, seen)
  }
}

function readbackOptions(kind: string, options: unknown): WorkbookReadbackCheckOptions {
  if (!isObject(options) || Array.isArray(options)) {
    throw new Error(`Workbook check ${kind} options must be an object`)
  }
  assertOnlyDataProperties(options, `Workbook check ${kind} options`)
  const message = optionalDataProperty(options, 'message', `Workbook check ${kind} options.message`)
  if (message.status === 'missing' || message.value === undefined) {
    return Object.freeze({})
  }
  return Object.freeze({
    message: requiredText(message.value, `${kind} options.message`),
  })
}

function customOptions(options: unknown): WorkbookCustomCheckOptions {
  if (!isObject(options) || Array.isArray(options)) {
    throw new Error('Workbook custom check options must be an object')
  }
  assertOnlyDataProperties(options, 'Workbook custom check options')
  const kind = requiredText(requiredDataProperty(options, 'kind', 'Workbook custom check kind'), 'kind')
  const message = requiredText(requiredDataProperty(options, 'message', 'Workbook custom check message'), 'message')
  const target = optionalDataProperty(options, 'target', 'Workbook custom check target')
  const refs = optionalDataProperty(options, 'refs', 'Workbook custom check refs')
  return Object.freeze({
    kind,
    message,
    ...(target.status === 'present' && target.value !== undefined ? { target: checkTarget(kind, target.value) } : {}),
    ...(refs.status === 'present' && refs.value !== undefined ? { refs: checkRefs(refs.value, 'Workbook custom check refs') } : {}),
  })
}

function createWorkbookCheck(options: WorkbookCheckBuildOptions): WorkbookCheckResult {
  const target = options.target === undefined ? undefined : checkTarget(options.kind, options.target)
  const refs = uniqueRefs(options.refs)
  const expectation =
    options.expectation === undefined
      ? undefined
      : options.expectation.kind === 'formulaEquals'
        ? Object.freeze({
            ...options.expectation,
            inputs: Object.freeze([...options.expectation.inputs]),
            labels: Object.freeze(options.expectation.labels.map((label) => Object.freeze({ ...label }))),
          })
        : Object.freeze({ ...options.expectation })
  return Object.freeze({
    status: 'planned',
    kind: requiredText(options.kind, 'kind'),
    ...(target !== undefined ? { target } : {}),
    ...(refs !== undefined ? { refs: Object.freeze([...refs]) } : {}),
    message: requiredText(options.message, 'message'),
    ...(expectation !== undefined ? { expectation } : {}),
  })
}

export function createWorkbookCustomCheck(options: WorkbookCustomCheckOptions): WorkbookCheckResult {
  const normalized = customOptions(options)
  const kind = normalized.kind
  if (isBuiltInWorkbookCheckKind(kind)) {
    throw new Error(`Workbook custom check kind ${kind} is reserved`)
  }
  return createWorkbookCheck(normalized)
}

export function createWorkbookCheckApi(record?: (check: WorkbookCheckResult) => void): WorkbookCheckApi {
  function planned(options: WorkbookCheckBuildOptions): WorkbookCheckResult {
    const check = createWorkbookCheck(options)
    record?.(check)
    return check
  }

  const api: WorkbookCheckApi = {
    exists(target) {
      const checkedTarget = checkTarget('exists', target)
      return planned({ kind: 'exists', target: checkedTarget, message: `${checkedTarget.label} exists` })
    },
    noFormulaErrors(target) {
      const checkedTarget = checkTarget('noFormulaErrors', target)
      return planned({ kind: 'noFormulaErrors', target: checkedTarget, message: `${checkedTarget.label} has no formula errors` })
    },
    valueEquals(target, value, options = {}) {
      const checkedTarget = checkTarget('valueEquals', target)
      const expected = checkedLiteralInput(value)
      const checkedOptions = readbackOptions('valueEquals', options)
      return planned({
        kind: 'valueEquals',
        target: checkedTarget,
        message: checkedOptions.message ?? `${checkedTarget.label} equals ${JSON.stringify(expected)}`,
        expectation: {
          kind: 'valueEquals',
          value: expected,
        },
      })
    },
    formulaEquals(target, value, options = {}) {
      const checkedTarget = checkTarget('formulaEquals', target)
      const checkedOptions = readbackOptions('formulaEquals', options)
      const source = formula.source(value)
      const inputs = formula.inputs(value)
      const labels = formula.labels(value)
      return planned({
        kind: 'formulaEquals',
        target: checkedTarget,
        message: checkedOptions.message ?? `${checkedTarget.label} formula equals ${source}`,
        expectation: {
          kind: 'formulaEquals',
          formula: source,
          inputs,
          labels,
        },
      })
    },
    custom(options) {
      const check = createWorkbookCustomCheck(options)
      record?.(check)
      return check
    },
  }
  return Object.freeze(api)
}

export const check: WorkbookCheckApi = createWorkbookCheckApi()
