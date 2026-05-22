import { isLiteralInput, type LiteralInput } from '@bilig/protocol'
import { isWorkbookRef, type WorkbookRef } from './find.js'
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
  readonly valuesEqual: (
    target: WorkbookRef,
    values: readonly (readonly LiteralInput[])[],
    options?: WorkbookReadbackCheckOptions,
  ) => WorkbookCheckResult
  readonly formulaEquals: (
    target: WorkbookRef,
    value: WorkbookFormulaOperand,
    options?: WorkbookReadbackCheckOptions,
  ) => WorkbookCheckResult
  readonly formulasEqual: (
    target: WorkbookRef,
    formulas: readonly (readonly (string | null)[])[],
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

function checkedRef(value: WorkbookRef, label: string): WorkbookRef {
  if (!isWorkbookRef(value)) {
    throw new Error(`Workbook check ${label} must be a WorkbookRef`)
  }
  return value
}

function checkedLiteralMatrix(values: readonly (readonly LiteralInput[])[]): readonly (readonly LiteralInput[])[] {
  if (!Array.isArray(values)) {
    throw new Error('Workbook readback values must be an array of rows')
  }
  let width: number | undefined
  return Object.freeze(
    values.map((row, rowIndex) => {
      if (!Array.isArray(row)) {
        throw new Error(`Workbook readback values row ${rowIndex.toString()} must be an array`)
      }
      width ??= row.length
      if (row.length !== width) {
        throw new Error('Workbook readback values must be rectangular')
      }
      return Object.freeze(row.map(checkedLiteralInput))
    }),
  )
}

function checkedFormulaMatrix(formulas: readonly (readonly (string | null)[])[]): readonly (readonly (string | null)[])[] {
  if (!Array.isArray(formulas)) {
    throw new Error('Workbook readback formulas must be an array of rows')
  }
  let width: number | undefined
  return Object.freeze(
    formulas.map((row, rowIndex) => {
      if (!Array.isArray(row)) {
        throw new Error(`Workbook readback formulas row ${rowIndex.toString()} must be an array`)
      }
      width ??= row.length
      if (row.length !== width) {
        throw new Error('Workbook readback formulas must be rectangular')
      }
      return Object.freeze(
        row.map((formulaText) => {
          if (formulaText !== null && typeof formulaText !== 'string') {
            throw new Error('Workbook readback formula must be a string or null')
          }
          return formulaText
        }),
      )
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
  if (!Array.isArray(refs)) {
    throw new Error('Workbook check refs must be an array')
  }
  const seen = new Set<string>()
  const unique: WorkbookRef[] = []
  for (const ref of refs) {
    checkedRef(ref, 'ref')
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
  const target = options.target === undefined ? undefined : checkedRef(options.target, 'target')
  return {
    status: 'planned',
    kind: requiredText(options.kind, 'kind'),
    ...(target !== undefined ? { target } : {}),
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
    valuesEqual(target, values, options = {}) {
      const expected = checkedLiteralMatrix(values)
      return planned({
        kind: 'valuesEqual',
        target,
        message: options.message ?? `${target.label} values equal ${JSON.stringify(expected)}`,
        expectation: {
          kind: 'valuesEqual',
          values: expected,
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
    formulasEqual(target, formulas, options = {}) {
      const expected = checkedFormulaMatrix(formulas)
      return planned({
        kind: 'formulasEqual',
        target,
        message: options.message ?? `${target.label} formulas equal ${JSON.stringify(expected)}`,
        expectation: {
          kind: 'formulasEqual',
          formulas: expected,
        },
      })
    },
    custom(options) {
      return planned(options)
    },
  }
}

export const check: WorkbookCheckApi = createWorkbookCheckApi()
