import type { LiteralInput } from '@bilig/protocol'
import type { WorkbookRef } from './find.js'
import type { WorkbookCheckResult } from './result.js'

export interface WorkbookRunReadback {
  readonly target: WorkbookRef
  readonly value?: LiteralInput
  readonly formula?: string | null
}

export type WorkbookReadbackIssueCode = 'readback_missing' | 'duplicate_readback' | 'value_mismatch' | 'formula_mismatch'

export const workbookReadbackIssueCodes = Object.freeze([
  'readback_missing',
  'duplicate_readback',
  'value_mismatch',
  'formula_mismatch',
] satisfies readonly WorkbookReadbackIssueCode[])

export function isWorkbookReadbackIssueCode(value: unknown): value is WorkbookReadbackIssueCode {
  return typeof value === 'string' && workbookReadbackIssueCodes.some((code) => code === value)
}

export interface WorkbookReadbackIssue {
  readonly code: WorkbookReadbackIssueCode
  readonly message: string
  readonly path?: string
  readonly check: WorkbookCheckResult
  readonly target?: WorkbookRef
  readonly expected?: LiteralInput
  readonly actual?: LiteralInput
}

export interface WorkbookReadbackVerification {
  readonly status: 'passed' | 'failed'
  readonly checks: readonly WorkbookCheckResult[]
  readonly issues: readonly WorkbookReadbackIssue[]
}

function refKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
}

function readbackKey(readback: WorkbookRunReadback): string {
  return refKey(readback.target)
}

function findReadback(readbacks: ReadonlyMap<string, WorkbookRunReadback>, target: WorkbookRef): WorkbookRunReadback | undefined {
  return readbacks.get(refKey(target))
}

function hasValue(readback: WorkbookRunReadback): readback is WorkbookRunReadback & { readonly value: LiteralInput } {
  return Object.prototype.hasOwnProperty.call(readback, 'value')
}

function hasFormula(readback: WorkbookRunReadback): readback is WorkbookRunReadback & { readonly formula: string | null } {
  return Object.prototype.hasOwnProperty.call(readback, 'formula')
}

function issue(input: WorkbookReadbackIssue): WorkbookReadbackIssue {
  return input
}

function missingReadback(check: WorkbookCheckResult): WorkbookReadbackIssue {
  return issue({
    code: 'readback_missing',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    message: `${check.target?.label ?? check.kind} has no readback`,
  })
}

function duplicateReadback(check: WorkbookCheckResult, readback: WorkbookRunReadback): WorkbookReadbackIssue {
  return issue({
    code: 'duplicate_readback',
    check,
    target: readback.target,
    message: `${readback.target.label} has more than one readback`,
  })
}

function valueMismatch(check: WorkbookCheckResult, expected: LiteralInput, actual?: LiteralInput): WorkbookReadbackIssue {
  return issue({
    code: 'value_mismatch',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.target?.label ?? check.kind} has no value readback`
        : `${check.target?.label ?? check.kind} expected value ${JSON.stringify(expected)} but read ${JSON.stringify(actual)}`,
  })
}

function formulaMismatch(check: WorkbookCheckResult, expected: string, actual: string | null | undefined): WorkbookReadbackIssue {
  return issue({
    code: 'formula_mismatch',
    check,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.target?.label ?? check.kind} has no formula readback`
        : `${check.target?.label ?? check.kind} expected formula ${expected} but read ${actual ?? 'null'}`,
  })
}

function checked(check: WorkbookCheckResult, status: WorkbookCheckResult['status']): WorkbookCheckResult {
  return {
    ...check,
    status,
  }
}

function verifyCheck(
  check: WorkbookCheckResult,
  readbacks: ReadonlyMap<string, WorkbookRunReadback>,
): { readonly check: WorkbookCheckResult; readonly issue?: WorkbookReadbackIssue } {
  if (check.expectation === undefined) {
    return { check }
  }

  if (check.target === undefined) {
    return {
      check: checked(check, 'failed'),
      issue: missingReadback(check),
    }
  }

  const readback = findReadback(readbacks, check.target)
  if (readback === undefined) {
    return {
      check: checked(check, 'failed'),
      issue: missingReadback(check),
    }
  }

  if (check.expectation.kind === 'valueEquals') {
    const actual = hasValue(readback) ? readback.value : undefined
    if (actual !== check.expectation.value) {
      return {
        check: checked(check, 'failed'),
        issue: valueMismatch(check, check.expectation.value, actual),
      }
    }
    return { check: checked(check, 'passed') }
  }

  const actual = hasFormula(readback) ? readback.formula : undefined
  if (actual !== check.expectation.formula) {
    return {
      check: checked(check, 'failed'),
      issue: formulaMismatch(check, check.expectation.formula, actual),
    }
  }
  return { check: checked(check, 'passed') }
}

export function verifyWorkbookReadbacks(
  checks: readonly WorkbookCheckResult[],
  readbacks: readonly WorkbookRunReadback[],
): WorkbookReadbackVerification {
  const readbackByTarget = new Map<string, WorkbookRunReadback>()
  const duplicateReadbacks = new Map<string, WorkbookRunReadback>()
  readbacks.forEach((readback) => {
    const key = readbackKey(readback)
    if (readbackByTarget.has(key)) {
      duplicateReadbacks.set(key, readback)
      return
    }
    readbackByTarget.set(key, readback)
  })

  const verifiedChecks: WorkbookCheckResult[] = []
  const issues: WorkbookReadbackIssue[] = []

  checks.forEach((check, checkIndex) => {
    if (check.target !== undefined) {
      const duplicate = duplicateReadbacks.get(refKey(check.target))
      if (duplicate !== undefined) {
        const failedCheck = checked(check, 'failed')
        verifiedChecks.push(failedCheck)
        issues.push({
          ...duplicateReadback(failedCheck, duplicate),
          path: `checks[${String(checkIndex)}]`,
        })
        return
      }
    }
    const result = verifyCheck(check, readbackByTarget)
    verifiedChecks.push(result.check)
    if (result.issue !== undefined) {
      issues.push({
        ...result.issue,
        path: `checks[${String(checkIndex)}]`,
      })
    }
  })

  return {
    status: issues.length === 0 ? 'passed' : 'failed',
    checks: verifiedChecks,
    issues,
  }
}
