import type { LiteralInput } from '@bilig/protocol'
import type { WorkbookRef } from './find.js'
import type { WorkbookCheckResult } from './result.js'

export interface WorkbookRunReadback {
  readonly target: WorkbookRef
  readonly value?: LiteralInput
  readonly formula?: string | null
}

export type WorkbookReadbackIssueCode = 'readback_missing' | 'readback_unexpected' | 'value_mismatch' | 'formula_mismatch'

export interface WorkbookReadbackIssue {
  readonly code: WorkbookReadbackIssueCode
  readonly message: string
  readonly check?: WorkbookCheckResult
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

function unexpectedReadback(readback: WorkbookRunReadback): WorkbookReadbackIssue {
  return issue({
    code: 'readback_unexpected',
    target: readback.target,
    message: `${readback.target.label} was returned by readback but was not requested`,
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

function checked(
  check: WorkbookCheckResult,
  status: WorkbookCheckResult['status'],
  proof?: WorkbookCheckResult['proof'],
): WorkbookCheckResult {
  return {
    ...check,
    status,
    ...(proof !== undefined ? { proof } : {}),
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
    return {
      check: checked(check, 'passed', {
        source: 'readback',
        value: actual,
      }),
    }
  }

  const actual = hasFormula(readback) ? readback.formula : undefined
  if (actual !== check.expectation.formula) {
    return {
      check: checked(check, 'failed'),
      issue: formulaMismatch(check, check.expectation.formula, actual),
    }
  }
  return {
    check: checked(check, 'passed', {
      source: 'readback',
      formula: actual,
    }),
  }
}

export function verifyWorkbookReadbacks(
  checks: readonly WorkbookCheckResult[],
  readbacks: readonly WorkbookRunReadback[],
): WorkbookReadbackVerification {
  const expectedTargets = new Set<string>()
  checks.forEach((check) => {
    if (check.expectation !== undefined && check.target !== undefined) {
      expectedTargets.add(refKey(check.target))
    }
  })

  const readbackByTarget = new Map<string, WorkbookRunReadback>()
  readbacks.forEach((readback) => {
    if (!readbackByTarget.has(readbackKey(readback))) {
      readbackByTarget.set(readbackKey(readback), readback)
    }
  })

  const verifiedChecks: WorkbookCheckResult[] = []
  const issues: WorkbookReadbackIssue[] = []

  readbackByTarget.forEach((readback, key) => {
    if (!expectedTargets.has(key)) {
      issues.push(unexpectedReadback(readback))
    }
  })

  checks.forEach((check) => {
    const result = verifyCheck(check, readbackByTarget)
    verifiedChecks.push(result.check)
    if (result.issue !== undefined) {
      issues.push(result.issue)
    }
  })

  return {
    status: issues.length === 0 ? 'passed' : 'failed',
    checks: verifiedChecks,
    issues,
  }
}
