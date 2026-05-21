import { isLiteralInput, type LiteralInput } from '@bilig/protocol'
import type { WorkbookRef } from './find.js'
import { formula, type WorkbookFormulaOperand } from './formula.js'
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

export function createWorkbookCheckResult(kind: string, target: WorkbookRef, message: string): WorkbookCheckResult {
  return createWorkbookCustomCheck({ kind, target, message })
}

interface WorkbookCheckBuildOptions extends WorkbookCustomCheckOptions {
  readonly expectation?: WorkbookCheckExpectation
}

function requiredText(value: string, name: string): string {
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

function createWorkbookCheck(options: WorkbookCheckBuildOptions): WorkbookCheckResult {
  const refs = uniqueRefs(options.refs)
  return {
    status: 'planned',
    kind: requiredText(options.kind, 'kind'),
    ...(options.target !== undefined ? { target: options.target } : {}),
    ...(refs !== undefined ? { refs } : {}),
    message: requiredText(options.message, 'message'),
    ...(options.expectation !== undefined ? { expectation: options.expectation } : {}),
  }
}

export function createWorkbookCustomCheck(options: WorkbookCustomCheckOptions): WorkbookCheckResult {
  return createWorkbookCheck(options)
}

export function createWorkbookCheckApi(record?: (check: WorkbookCheckResult) => void): WorkbookCheckApi {
  function planned(options: WorkbookCheckBuildOptions): WorkbookCheckResult {
    const check = createWorkbookCheck(options)
    record?.(check)
    return check
  }

  return {
    exists(target) {
      return planned({ kind: 'exists', target, message: `${target.label} exists` })
    },
    noFormulaErrors(target) {
      return planned({ kind: 'noFormulaErrors', target, message: `${target.label} has no formula errors` })
    },
    valueEquals(target, value, options = {}) {
      const expected = checkedLiteralInput(value)
      return planned({
        kind: 'valueEquals',
        target,
        message: options.message ?? `${target.label} equals ${JSON.stringify(expected)}`,
        expectation: {
          kind: 'valueEquals',
          value: expected,
        },
      })
    },
    formulaEquals(target, value, options = {}) {
      const source = formula.source(value)
      const inputs = formula.inputs(value)
      return planned({
        kind: 'formulaEquals',
        target,
        message: options.message ?? `${target.label} formula equals ${source}`,
        expectation: {
          kind: 'formulaEquals',
          formula: source,
          inputs,
        },
      })
    },
    custom(options) {
      return planned(options)
    },
  }
}

export const check: WorkbookCheckApi = createWorkbookCheckApi()
