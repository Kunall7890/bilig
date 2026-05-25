import { parseFormula, serializeFormula } from '@bilig/formula'
import { isLiteralInput, type LiteralInput } from '@bilig/protocol'
import { hydrateWorkbookRef, isWorkbookRef, toWorkbookRefData, type WorkbookRef } from './find.js'
import { materializeFormulaLabels, type WorkbookFormulaLabelReplacement } from './formula-usage.js'
import type { WorkbookFormulaLabel } from './formula.js'
import { normalizeWorkbookActionInput } from './input.js'
import type { WorkbookCheckExpectation, WorkbookCheckResult } from './result.js'

export interface WorkbookRunReadback {
  readonly target: WorkbookRef
  readonly value?: LiteralInput
  readonly formula?: string | null
  readonly formulaLabels?: readonly WorkbookFormulaLabelReplacement[]
}

export type WorkbookReadbackIssueCode =
  | 'readback_invalid'
  | 'readback_duplicate'
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
  return Object.freeze(input)
}

function invalidReadback(message: string): WorkbookReadbackIssue {
  return issue({
    code: 'readback_invalid',
    message,
  })
}

function duplicateReadback(readback: SafeReadback): WorkbookReadbackIssue {
  return issue({
    code: 'readback_duplicate',
    target: readback.target,
    message: `${readback.target.label} was returned by readback more than once`,
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
  readonly formulaLabels: SafeReadbackData<readonly WorkbookFormulaLabelReplacement[]>
}

function valid<T>(value: T): DataValidation<T> {
  return Object.freeze({
    status: 'valid',
    value,
  })
}

function invalid<T>(message: string): DataValidation<T> {
  return Object.freeze({
    status: 'invalid',
    issue: invalidReadback(message),
  })
}

function isRecord(value: unknown): value is object {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function optionalDataValue(value: object, key: string, path: string): OptionalData {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return Object.freeze({ status: 'missing' })
  }
  if (!('value' in descriptor)) {
    return Object.freeze({
      status: 'invalid',
      issue: invalidReadback(`Workbook readback proof at ${path} must be a data property`),
    })
  }
  return Object.freeze({
    status: 'present',
    value: descriptor.value,
  })
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
  return valid(Object.freeze(entries))
}

function normalizeRef(value: unknown, path: string): DataValidation<WorkbookRef> {
  if (!isWorkbookRef(value)) {
    return invalid(`Workbook reference at ${path} is invalid`)
  }
  return valid(hydrateWorkbookRef(toWorkbookRefData(value)))
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
  return valid(
    Object.freeze({
      name: name.value,
      ref: checkedRef.value,
    }),
  )
}

function normalizeFormulaLabels(value: unknown, path: string): DataValidation<readonly WorkbookFormulaLabel[]> {
  return arrayDataValues(value, path, normalizeFormulaLabel)
}

function normalizeFormulaLabelReplacement(value: unknown, path: string): DataValidation<WorkbookFormulaLabelReplacement> {
  if (!isRecord(value)) {
    return invalid(`Workbook formula label proof at ${path} is invalid`)
  }
  const name = requiredDataValue(value, 'name', `${path}.name`)
  if (name.status === 'invalid') {
    return name
  }
  if (typeof name.value !== 'string') {
    return invalid(`Workbook formula label proof at ${path}.name is invalid`)
  }
  const source = requiredDataValue(value, 'source', `${path}.source`)
  if (source.status === 'invalid') {
    return source
  }
  if (typeof source.value !== 'string') {
    return invalid(`Workbook formula label proof at ${path}.source is invalid`)
  }
  const normalizedSource = normalizeFormulaProofSource(source.value, `${path}.source`)
  if (normalizedSource.status === 'invalid') {
    return normalizedSource
  }
  return valid(
    Object.freeze({
      name: name.value,
      source: normalizedSource.value,
    }),
  )
}

function normalizeFormulaLabelReplacements(value: unknown, path: string): DataValidation<readonly WorkbookFormulaLabelReplacement[]> {
  return arrayDataValues(value, path, normalizeFormulaLabelReplacement)
}

function normalizeFormulaProofSource(value: string, path: string): DataValidation<string> {
  const trimmed = value.trim()
  const source = trimmed.startsWith('=') ? trimmed.slice(1).trim() : trimmed
  if (source.length === 0) {
    return invalid(`Workbook formula proof at ${path} cannot be empty`)
  }
  try {
    return valid(serializeFormula(parseFormula(source)))
  } catch (error) {
    return invalid(`Workbook formula proof at ${path} is invalid: ${error instanceof Error ? error.message : String(error)}`)
  }
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
    return valid(
      Object.freeze({
        kind: 'valueEquals',
        value: expected.value,
      }),
    )
  }

  if (kind.value === 'formulaEquals') {
    const formula = requiredDataValue(value, 'formula', `${path}.formula`)
    if (formula.status === 'invalid') {
      return formula
    }
    if (typeof formula.value !== 'string') {
      return invalid(`Workbook check expectation at ${path}.formula is invalid`)
    }
    const normalizedFormula = normalizeFormulaProofSource(formula.value, `${path}.formula`)
    if (normalizedFormula.status === 'invalid') {
      return normalizedFormula
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

    return valid(
      Object.freeze({
        kind: 'formulaEquals',
        formula: normalizedFormula.value,
        inputs: normalizedInputs?.value ?? [],
        labels: normalizedLabels?.value ?? [],
      }),
    )
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
    try {
      proofValue = normalizeWorkbookActionInput(proof.value)
    } catch {
      return invalid(`Workbook check at ${path}.proof is invalid`)
    }
  }

  const result: WorkbookCheckResult = Object.freeze({
    status: status.value,
    kind: kind.value,
    ...(checkedTarget !== undefined ? { target: checkedTarget.value } : {}),
    ...(checkedRefs !== undefined ? { refs: checkedRefs.value } : {}),
    message: message.value,
    ...(checkedExpectation !== undefined ? { expectation: checkedExpectation.value } : {}),
    ...(proofValue !== undefined ? { proof: proofValue } : {}),
  })

  return valid(
    Object.freeze({
      result,
      kind: kind.value,
      ...(checkedTarget !== undefined ? { target: checkedTarget.value } : {}),
      ...(checkedExpectation !== undefined ? { expectation: checkedExpectation.value } : {}),
      label: checkedTarget?.value.label ?? kind.value,
    }),
  )
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
  let safeValue: SafeReadbackData<LiteralInput> = Object.freeze({ status: 'missing' })
  if (valueReadback.status === 'present') {
    const candidate = valueReadback.value
    if (!isLiteralInput(candidate)) {
      return invalid(`Workbook readback at ${path}.value is invalid`)
    }
    safeValue = Object.freeze({
      status: 'present',
      value: candidate,
    })
  }

  const formulaReadback = optionalDataValue(value, 'formula', `${path}.formula`)
  if (formulaReadback.status === 'invalid') {
    return formulaReadback
  }
  let safeFormula: SafeReadbackData<string | null> = Object.freeze({ status: 'missing' })
  if (formulaReadback.status === 'present') {
    const candidate = formulaReadback.value
    if (typeof candidate !== 'string' && candidate !== null) {
      return invalid(`Workbook readback at ${path}.formula is invalid`)
    }
    const normalizedFormula = candidate === null ? valid(null) : normalizeFormulaProofSource(candidate, `${path}.formula`)
    if (normalizedFormula.status === 'invalid') {
      return normalizedFormula
    }
    safeFormula = Object.freeze({
      status: 'present',
      value: normalizedFormula.value,
    })
  }

  const formulaLabels = optionalDataValue(value, 'formulaLabels', `${path}.formulaLabels`)
  if (formulaLabels.status === 'invalid') {
    return formulaLabels
  }
  let safeFormulaLabels: SafeReadbackData<readonly WorkbookFormulaLabelReplacement[]> = Object.freeze({ status: 'missing' })
  if (formulaLabels.status === 'present') {
    const normalizedLabels = normalizeFormulaLabelReplacements(formulaLabels.value, `${path}.formulaLabels`)
    if (normalizedLabels.status === 'invalid') {
      return normalizedLabels
    }
    safeFormulaLabels = Object.freeze({
      status: 'present',
      value: normalizedLabels.value,
    })
  }

  const result: WorkbookRunReadback = Object.freeze({
    target: checkedTarget.value,
    ...(safeValue.status === 'present' ? { value: safeValue.value } : {}),
    ...(safeFormula.status === 'present' ? { formula: safeFormula.value } : {}),
    ...(safeFormulaLabels.status === 'present' ? { formulaLabels: safeFormulaLabels.value } : {}),
  })

  return valid(
    Object.freeze({
      result,
      target: checkedTarget.value,
      key: refKey(checkedTarget.value),
      value: safeValue,
      formula: safeFormula,
      formulaLabels: safeFormulaLabels,
    }),
  )
}

function checked(check: SafeCheck, status: WorkbookCheckResult['status'], proof?: WorkbookCheckResult['proof']): WorkbookCheckResult {
  const normalizedProof = proof === undefined ? undefined : normalizeWorkbookActionInput(proof)
  return Object.freeze({
    ...check.result,
    status,
    ...(normalizedProof !== undefined ? { proof: normalizedProof } : {}),
  })
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
  let expected = check.expectation.formula
  const formulaLabels = readback.formulaLabels.status === 'present' ? readback.formulaLabels.value : undefined
  if (formulaLabels !== undefined && formulaLabels.length > 0) {
    try {
      expected = materializeFormulaLabels(check.expectation.formula, formulaLabels)
    } catch (error) {
      return {
        check: checked(check, 'failed'),
        issue: invalidReadback(`Workbook formula proof for ${check.label} is invalid: ${errorMessage(error)}`),
      }
    }
  }
  if (actual !== expected) {
    return {
      check: checked(check, 'failed'),
      issue: formulaMismatch(check, expected, actual),
    }
  }
  const proof = normalizeWorkbookActionInput({
    source: 'readback',
    formula: actual,
    expectedFormula: check.expectation.formula,
    materializedFormula: expected,
    ...(formulaLabels !== undefined ? { formulaLabels: formulaLabels.map((label) => ({ name: label.name, source: label.source })) } : {}),
  })
  return {
    check: checked(check, 'passed', proof),
  }
}

export function verifyWorkbookReadbacks(
  checks: readonly WorkbookCheckResult[],
  readbacks: readonly WorkbookRunReadback[],
): WorkbookReadbackVerification {
  const checkValidation = arrayDataValues(checks, 'checks', normalizeCheck)
  if (checkValidation.status === 'invalid') {
    return Object.freeze({
      status: 'failed',
      checks: Object.freeze([]),
      issues: Object.freeze([checkValidation.issue]),
    })
  }

  const readbackValidation = arrayDataValues(readbacks, 'readbacks', normalizeReadback)
  if (readbackValidation.status === 'invalid') {
    return Object.freeze({
      status: 'failed',
      checks: Object.freeze(checkValidation.value.map((check) => check.result)),
      issues: Object.freeze([readbackValidation.issue]),
    })
  }

  const expectedTargets = new Set<string>()
  checkValidation.value.forEach((check) => {
    if (check.expectation !== undefined && check.target !== undefined) {
      expectedTargets.add(refKey(check.target))
    }
  })

  const readbackByTarget = new Map<string, SafeReadback>()
  const issues: WorkbookReadbackIssue[] = []
  readbackValidation.value.forEach((readback) => {
    if (readbackByTarget.has(readback.key)) {
      issues.push(duplicateReadback(readback))
    } else {
      readbackByTarget.set(readback.key, readback)
    }
  })

  const verifiedChecks: WorkbookCheckResult[] = []

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

  return Object.freeze({
    status: issues.length === 0 ? 'passed' : 'failed',
    checks: Object.freeze(verifiedChecks),
    issues: Object.freeze(issues),
  })
}
