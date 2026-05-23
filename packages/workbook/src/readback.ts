import { isLiteralInput, type LiteralInput } from '@bilig/protocol'
import { isWorkbookRef, type WorkbookRef } from './find.js'
import type { WorkbookFormulaLabel } from './formula.js'
import { isWorkbookActionInput } from './input.js'
import type { WorkbookCheckExpectation, WorkbookCheckResult } from './result.js'

export interface WorkbookRunReadback {
  readonly target: WorkbookRef
  readonly value?: LiteralInput
  readonly formula?: string | null
}

export type WorkbookReadbackIssueCode =
  | 'readback_invalid'
  | 'readback_missing'
  | 'readback_unexpected'
  | 'value_mismatch'
  | 'formula_mismatch'

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

function issue(input: WorkbookReadbackIssue): WorkbookReadbackIssue {
  return input
}

function invalidReadback(message: string): WorkbookReadbackIssue {
  return issue({
    code: 'readback_invalid',
    message,
  })
}

function missingReadback(check: SafeCheck): WorkbookReadbackIssue {
  return issue({
    code: 'readback_missing',
    check: check.result,
    ...(check.target !== undefined ? { target: check.target } : {}),
    message: `${check.label} has no readback`,
  })
}

function unexpectedReadback(readback: SafeReadback): WorkbookReadbackIssue {
  return issue({
    code: 'readback_unexpected',
    target: readback.target,
    message: `${readback.target.label} was returned by readback but was not requested`,
  })
}

function valueMismatch(check: SafeCheck, expected: LiteralInput, actual?: LiteralInput): WorkbookReadbackIssue {
  return issue({
    code: 'value_mismatch',
    check: check.result,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.label} has no value readback`
        : `${check.label} expected value ${JSON.stringify(expected)} but read ${JSON.stringify(actual)}`,
  })
}

function formulaMismatch(check: SafeCheck, expected: string, actual: string | null | undefined): WorkbookReadbackIssue {
  return issue({
    code: 'formula_mismatch',
    check: check.result,
    ...(check.target !== undefined ? { target: check.target } : {}),
    expected,
    ...(actual !== undefined ? { actual } : {}),
    message:
      actual === undefined
        ? `${check.label} has no formula readback`
        : `${check.label} expected formula ${expected} but read ${actual ?? 'null'}`,
  })
}

type DataValidation<T> =
  | {
      readonly status: 'valid'
      readonly value: T
    }
  | {
      readonly status: 'invalid'
      readonly issue: WorkbookReadbackIssue
    }

type OptionalData =
  | {
      readonly status: 'missing'
    }
  | {
      readonly status: 'present'
      readonly value: unknown
    }
  | {
      readonly status: 'invalid'
      readonly issue: WorkbookReadbackIssue
    }

type SafeReadbackData<T> =
  | {
      readonly status: 'present'
      readonly value: T
    }
  | {
      readonly status: 'missing'
    }

interface SafeCheck {
  readonly result: WorkbookCheckResult
  readonly kind: string
  readonly target?: WorkbookRef
  readonly expectation?: WorkbookCheckExpectation
  readonly label: string
}

interface SafeReadback {
  readonly result: WorkbookRunReadback
  readonly target: WorkbookRef
  readonly key: string
  readonly value: SafeReadbackData<LiteralInput>
  readonly formula: SafeReadbackData<string | null>
}

function valid<T>(value: T): DataValidation<T> {
  return {
    status: 'valid',
    value,
  }
}

function invalid<T>(message: string): DataValidation<T> {
  return {
    status: 'invalid',
    issue: invalidReadback(message),
  }
}

function isRecord(value: unknown): value is object {
  return typeof value === 'object' && value !== null
}

function optionalDataValue(value: object, key: string, path: string): OptionalData {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return { status: 'missing' }
  }
  if (!('value' in descriptor)) {
    return {
      status: 'invalid',
      issue: invalidReadback(`Workbook readback proof at ${path} must be a data property`),
    }
  }
  return {
    status: 'present',
    value: descriptor.value,
  }
}

function requiredDataValue(value: object, key: string, path: string): DataValidation<unknown> {
  const property = optionalDataValue(value, key, path)
  if (property.status === 'invalid') {
    return property
  }
  if (property.status === 'missing') {
    return invalid(`Workbook readback proof at ${path} is missing`)
  }
  return valid(property.value)
}

function arrayDataValues<T>(
  value: unknown,
  path: string,
  mapEntry: (entry: unknown, entryPath: string) => DataValidation<T>,
): DataValidation<readonly T[]> {
  if (!Array.isArray(value)) {
    return invalid(`Workbook readback proof at ${path} must be an array`)
  }

  const entries: T[] = []
  for (let index = 0; index < value.length; index += 1) {
    const entryPath = `${path}[${index}]`
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      return invalid(`Workbook readback proof at ${entryPath} is invalid`)
    }
    const entry = mapEntry(descriptor.value, entryPath)
    if (entry.status === 'invalid') {
      return entry
    }
    entries.push(entry.value)
  }
  return valid(entries)
}

function normalizeRef(value: unknown, path: string): DataValidation<WorkbookRef> {
  if (!isWorkbookRef(value)) {
    return invalid(`Workbook reference at ${path} is invalid`)
  }
  return valid(value)
}

function normalizeRefArray(value: unknown, path: string): DataValidation<readonly WorkbookRef[]> {
  return arrayDataValues(value, path, normalizeRef)
}

function normalizeFormulaLabel(value: unknown, path: string): DataValidation<WorkbookFormulaLabel> {
  if (!isRecord(value)) {
    return invalid(`Workbook formula label at ${path} is invalid`)
  }
  const name = requiredDataValue(value, 'name', `${path}.name`)
  if (name.status === 'invalid') {
    return name
  }
  if (typeof name.value !== 'string') {
    return invalid(`Workbook formula label at ${path}.name is invalid`)
  }
  const ref = requiredDataValue(value, 'ref', `${path}.ref`)
  if (ref.status === 'invalid') {
    return ref
  }
  const checkedRef = normalizeRef(ref.value, `${path}.ref`)
  if (checkedRef.status === 'invalid') {
    return checkedRef
  }
  return valid({
    name: name.value,
    ref: checkedRef.value,
  })
}

function normalizeFormulaLabels(value: unknown, path: string): DataValidation<readonly WorkbookFormulaLabel[]> {
  return arrayDataValues(value, path, normalizeFormulaLabel)
}

function normalizeExpectation(value: unknown, path: string): DataValidation<WorkbookCheckExpectation> {
  if (!isRecord(value)) {
    return invalid(`Workbook check expectation at ${path} is invalid`)
  }

  const kind = requiredDataValue(value, 'kind', `${path}.kind`)
  if (kind.status === 'invalid') {
    return kind
  }

  if (kind.value === 'valueEquals') {
    const expected = requiredDataValue(value, 'value', `${path}.value`)
    if (expected.status === 'invalid') {
      return expected
    }
    if (!isLiteralInput(expected.value)) {
      return invalid(`Workbook check expectation at ${path}.value is invalid`)
    }
    return valid({
      kind: 'valueEquals',
      value: expected.value,
    })
  }

  if (kind.value === 'formulaEquals') {
    const formula = requiredDataValue(value, 'formula', `${path}.formula`)
    if (formula.status === 'invalid') {
      return formula
    }
    if (typeof formula.value !== 'string') {
      return invalid(`Workbook check expectation at ${path}.formula is invalid`)
    }

    const inputs = optionalDataValue(value, 'inputs', `${path}.inputs`)
    if (inputs.status === 'invalid') {
      return inputs
    }
    const labels = optionalDataValue(value, 'labels', `${path}.labels`)
    if (labels.status === 'invalid') {
      return labels
    }

    const normalizedInputs = inputs.status === 'present' ? normalizeRefArray(inputs.value, `${path}.inputs`) : undefined
    if (normalizedInputs?.status === 'invalid') {
      return normalizedInputs
    }
    const normalizedLabels = labels.status === 'present' ? normalizeFormulaLabels(labels.value, `${path}.labels`) : undefined
    if (normalizedLabels?.status === 'invalid') {
      return normalizedLabels
    }

    return valid({
      kind: 'formulaEquals',
      formula: formula.value,
      inputs: normalizedInputs?.value ?? [],
      labels: normalizedLabels?.value ?? [],
    })
  }

  return invalid(`Workbook check expectation at ${path}.kind is invalid`)
}

function normalizeCheck(value: unknown, path: string): DataValidation<SafeCheck> {
  if (!isRecord(value)) {
    return invalid(`Workbook check at ${path} is invalid`)
  }

  const status = requiredDataValue(value, 'status', `${path}.status`)
  if (status.status === 'invalid') {
    return status
  }
  if (status.value !== 'planned' && status.value !== 'passed' && status.value !== 'failed') {
    return invalid(`Workbook check at ${path}.status is invalid`)
  }

  const kind = requiredDataValue(value, 'kind', `${path}.kind`)
  if (kind.status === 'invalid') {
    return kind
  }
  if (typeof kind.value !== 'string') {
    return invalid(`Workbook check at ${path}.kind is invalid`)
  }

  const message = requiredDataValue(value, 'message', `${path}.message`)
  if (message.status === 'invalid') {
    return message
  }
  if (typeof message.value !== 'string') {
    return invalid(`Workbook check at ${path}.message is invalid`)
  }

  const target = optionalDataValue(value, 'target', `${path}.target`)
  if (target.status === 'invalid') {
    return target
  }
  const checkedTarget = target.status === 'present' ? normalizeRef(target.value, `${path}.target`) : undefined
  if (checkedTarget?.status === 'invalid') {
    return checkedTarget
  }

  const refs = optionalDataValue(value, 'refs', `${path}.refs`)
  if (refs.status === 'invalid') {
    return refs
  }
  const checkedRefs = refs.status === 'present' ? normalizeRefArray(refs.value, `${path}.refs`) : undefined
  if (checkedRefs?.status === 'invalid') {
    return checkedRefs
  }

  const expectation = optionalDataValue(value, 'expectation', `${path}.expectation`)
  if (expectation.status === 'invalid') {
    return expectation
  }
  const checkedExpectation = expectation.status === 'present' ? normalizeExpectation(expectation.value, `${path}.expectation`) : undefined
  if (checkedExpectation?.status === 'invalid') {
    return checkedExpectation
  }

  const proof = optionalDataValue(value, 'proof', `${path}.proof`)
  if (proof.status === 'invalid') {
    return proof
  }
  let proofValue: WorkbookCheckResult['proof'] | undefined
  if (proof.status === 'present') {
    const candidate = proof.value
    if (!isWorkbookActionInput(candidate)) {
      return invalid(`Workbook check at ${path}.proof is invalid`)
    }
    proofValue = candidate
  }

  const result: WorkbookCheckResult = {
    status: status.value,
    kind: kind.value,
    ...(checkedTarget !== undefined ? { target: checkedTarget.value } : {}),
    ...(checkedRefs !== undefined ? { refs: checkedRefs.value } : {}),
    message: message.value,
    ...(checkedExpectation !== undefined ? { expectation: checkedExpectation.value } : {}),
    ...(proofValue !== undefined ? { proof: proofValue } : {}),
  }

  return valid({
    result,
    kind: kind.value,
    ...(checkedTarget !== undefined ? { target: checkedTarget.value } : {}),
    ...(checkedExpectation !== undefined ? { expectation: checkedExpectation.value } : {}),
    label: checkedTarget?.value.label ?? kind.value,
  })
}

function normalizeReadback(value: unknown, path: string): DataValidation<SafeReadback> {
  if (!isRecord(value)) {
    return invalid(`Workbook readback at ${path} is invalid`)
  }

  const target = requiredDataValue(value, 'target', `${path}.target`)
  if (target.status === 'invalid') {
    return target
  }
  const checkedTarget = normalizeRef(target.value, `${path}.target`)
  if (checkedTarget.status === 'invalid') {
    return checkedTarget
  }

  const valueReadback = optionalDataValue(value, 'value', `${path}.value`)
  if (valueReadback.status === 'invalid') {
    return valueReadback
  }
  let safeValue: SafeReadbackData<LiteralInput> = { status: 'missing' }
  if (valueReadback.status === 'present') {
    const candidate = valueReadback.value
    if (!isLiteralInput(candidate)) {
      return invalid(`Workbook readback at ${path}.value is invalid`)
    }
    safeValue = {
      status: 'present',
      value: candidate,
    }
  }

  const formulaReadback = optionalDataValue(value, 'formula', `${path}.formula`)
  if (formulaReadback.status === 'invalid') {
    return formulaReadback
  }
  let safeFormula: SafeReadbackData<string | null> = { status: 'missing' }
  if (formulaReadback.status === 'present') {
    const candidate = formulaReadback.value
    if (typeof candidate !== 'string' && candidate !== null) {
      return invalid(`Workbook readback at ${path}.formula is invalid`)
    }
    safeFormula = {
      status: 'present',
      value: candidate,
    }
  }

  const result: WorkbookRunReadback = {
    target: checkedTarget.value,
    ...(safeValue.status === 'present' ? { value: safeValue.value } : {}),
    ...(safeFormula.status === 'present' ? { formula: safeFormula.value } : {}),
  }

  return valid({
    result,
    target: checkedTarget.value,
    key: refKey(checkedTarget.value),
    value: safeValue,
    formula: safeFormula,
  })
}

function checked(check: SafeCheck, status: WorkbookCheckResult['status'], proof?: WorkbookCheckResult['proof']): WorkbookCheckResult {
  return {
    ...check.result,
    status,
    ...(proof !== undefined ? { proof } : {}),
  }
}

function verifyCheck(
  check: SafeCheck,
  readbacks: ReadonlyMap<string, SafeReadback>,
): { readonly check: WorkbookCheckResult; readonly issue?: WorkbookReadbackIssue } {
  if (check.expectation === undefined) {
    return { check: check.result }
  }

  if (check.target === undefined) {
    return {
      check: checked(check, 'failed'),
      issue: missingReadback(check),
    }
  }

  const readback = readbacks.get(refKey(check.target))
  if (readback === undefined) {
    return {
      check: checked(check, 'failed'),
      issue: missingReadback(check),
    }
  }

  if (check.expectation.kind === 'valueEquals') {
    const actual = readback.value.status === 'present' ? readback.value.value : undefined
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

  const actual = readback.formula.status === 'present' ? readback.formula.value : undefined
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
  const checkValidation = arrayDataValues(checks, 'checks', normalizeCheck)
  if (checkValidation.status === 'invalid') {
    return {
      status: 'failed',
      checks: [],
      issues: [checkValidation.issue],
    }
  }

  const readbackValidation = arrayDataValues(readbacks, 'readbacks', normalizeReadback)
  if (readbackValidation.status === 'invalid') {
    return {
      status: 'failed',
      checks: checkValidation.value.map((check) => check.result),
      issues: [readbackValidation.issue],
    }
  }

  const expectedTargets = new Set<string>()
  checkValidation.value.forEach((check) => {
    if (check.expectation !== undefined && check.target !== undefined) {
      expectedTargets.add(refKey(check.target))
    }
  })

  const readbackByTarget = new Map<string, SafeReadback>()
  readbackValidation.value.forEach((readback) => {
    if (!readbackByTarget.has(readback.key)) {
      readbackByTarget.set(readback.key, readback)
    }
  })

  const verifiedChecks: WorkbookCheckResult[] = []
  const issues: WorkbookReadbackIssue[] = []

  readbackByTarget.forEach((readback, key) => {
    if (!expectedTargets.has(key)) {
      issues.push(unexpectedReadback(readback))
    }
  })

  checkValidation.value.forEach((check) => {
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
