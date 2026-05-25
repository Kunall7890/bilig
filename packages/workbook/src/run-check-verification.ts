import type { WorkbookCheckResult, WorkbookRunError } from './result.js'
import { normalizeWorkbookActionInput } from './input.js'
import type { WorkbookActionPlan } from './model.js'
import { adapterVerifyChecksMethod } from './run-runtime-boundary.js'
import { canonicalJson, isRecord, ownValue } from './run-data.js'
import { errorMessage, runError } from './run-failure.js'
import type { WorkbookRunAdapter } from './run.js'

type CheckValidation =
  | {
      readonly status: 'valid'
      readonly checks: readonly WorkbookCheckResult[]
    }
  | {
      readonly status: 'invalid'
      readonly error: WorkbookRunError
    }

export async function verifyChecksWithAdapter<Refs>(
  checks: readonly WorkbookCheckResult[],
  plan: WorkbookActionPlan<Refs>,
  adapter: WorkbookRunAdapter<Refs>,
): Promise<{ readonly checks: readonly WorkbookCheckResult[]; readonly errors: readonly WorkbookRunError[] }> {
  const verifyChecks = adapterVerifyChecksMethod(adapter)
  if (verifyChecks === undefined) {
    return { checks, errors: [] }
  }

  const originalChecks = checks.map(cloneCheck)
  const originalContracts = originalChecks.map((check) => canonicalJson(checkContract(check)))
  const originalKinds = originalChecks.map((check) => check.kind)
  const verifierInput = originalChecks.map(cloneCheck)
  let verified: unknown
  try {
    verified = await Reflect.apply(verifyChecks, adapter, [verifierInput, plan])
  } catch (error) {
    return {
      checks: originalChecks,
      errors: [runError('check_verification_failed', errorMessage(error))],
    }
  }

  const validation = validateVerifiedChecks(originalContracts, originalKinds, verified)
  if (validation.status === 'invalid') {
    return {
      checks: originalChecks,
      errors: [validation.error],
    }
  }

  const verifiedChecks = validation.checks
  const failedChecks = verifiedChecks.filter((check) => check.status === 'failed')
  if (failedChecks.length > 0) {
    return {
      checks: verifiedChecks,
      errors: failedChecks.map((check) =>
        runError('check_failed', `${check.target?.label ?? check.kind} failed check ${check.kind}: ${check.message}`),
      ),
    }
  }

  return { checks: verifiedChecks, errors: [] }
}

export function unverifiedCheckErrors(checks: readonly WorkbookCheckResult[]): readonly WorkbookRunError[] {
  return checks
    .filter((check) => check.status === 'planned')
    .map((check) => runError('check_not_verified', `${checkLabel(check)} did not verify check ${check.kind}: ${check.message}`))
}

function validateVerifiedChecks(
  originalContracts: readonly string[],
  originalKinds: readonly string[],
  verified: unknown,
): CheckValidation {
  if (!Array.isArray(verified)) {
    return {
      status: 'invalid',
      error: runError('invalid_check_verification', 'Check verifier did not return a check array'),
    }
  }

  if (verified.length !== originalContracts.length) {
    return {
      status: 'invalid',
      error: runError(
        'invalid_check_verification',
        `Check verifier returned ${String(verified.length)} checks for ${String(originalContracts.length)} planned checks`,
      ),
    }
  }

  const verifiedChecks: WorkbookCheckResult[] = []
  for (let index = 0; index < originalContracts.length; index += 1) {
    const expectedContract = originalContracts[index]
    const expectedKind = originalKinds[index] ?? 'check'
    const descriptor = Object.getOwnPropertyDescriptor(verified, String(index))
    const actual = descriptor !== undefined && descriptor.enumerable && 'value' in descriptor ? descriptor.value : undefined
    if (expectedContract === undefined || !isRecord(actual)) {
      return {
        status: 'invalid',
        error: runError('invalid_check_verification', `Check verifier returned an invalid check at index ${String(index)}`),
      }
    }
    if (!isCheckStatus(ownValue(actual, 'status'))) {
      return {
        status: 'invalid',
        error: runError('invalid_check_verification', `Check verifier returned an invalid status at index ${String(index)}`),
      }
    }
    if (!isWorkbookCheckResult(actual)) {
      return {
        status: 'invalid',
        error: runError('invalid_check_verification', `Check verifier returned an invalid check at index ${String(index)}`),
      }
    }
    if (!checkContractMatches(expectedContract, actual)) {
      return {
        status: 'invalid',
        error: runError(
          'invalid_check_verification',
          `Check verifier changed the check contract at index ${String(index)} for ${expectedKind}`,
        ),
      }
    }
    try {
      verifiedChecks.push(cloneCheck(actual))
    } catch (error) {
      return {
        status: 'invalid',
        error: runError(
          'invalid_check_verification',
          `Check verifier returned invalid proof at index ${String(index)}: ${errorMessage(error)}`,
        ),
      }
    }
  }

  return {
    status: 'valid',
    checks: verifiedChecks,
  }
}

function checkLabel(check: WorkbookCheckResult): string {
  return check.target?.label ?? check.kind
}

function ownCheckValue<Key extends keyof WorkbookCheckResult>(check: WorkbookCheckResult, key: Key): WorkbookCheckResult[Key] | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(check, key)
  if (descriptor === undefined) {
    return undefined
  }
  if (!('value' in descriptor)) {
    throw new Error(`Workbook check result ${String(key)} must be a data property`)
  }
  return descriptor.value
}

function checkContract(check: WorkbookCheckResult): Omit<WorkbookCheckResult, 'status'> {
  const kind = ownCheckValue(check, 'kind')
  const target = ownCheckValue(check, 'target')
  const refs = ownCheckValue(check, 'refs')
  const message = ownCheckValue(check, 'message')
  const expectation = ownCheckValue(check, 'expectation')
  if (kind === undefined || message === undefined) {
    throw new Error('invalid check contract')
  }

  return {
    kind,
    ...(target !== undefined ? { target } : {}),
    ...(refs !== undefined ? { refs } : {}),
    message,
    ...(expectation !== undefined ? { expectation } : {}),
  }
}

function cloneCheck(check: WorkbookCheckResult): WorkbookCheckResult {
  const status = ownCheckValue(check, 'status')
  const kind = ownCheckValue(check, 'kind')
  const target = ownCheckValue(check, 'target')
  const refs = ownCheckValue(check, 'refs')
  const message = ownCheckValue(check, 'message')
  const expectation = ownCheckValue(check, 'expectation')
  const proof = ownCheckValue(check, 'proof')
  if (status === undefined || kind === undefined || message === undefined) {
    throw new Error('invalid check result')
  }

  return {
    status,
    kind,
    ...(target !== undefined ? { target } : {}),
    ...(refs !== undefined ? { refs } : {}),
    message,
    ...(expectation !== undefined ? { expectation } : {}),
    ...(proof !== undefined ? { proof: normalizeWorkbookActionInput(proof) } : {}),
  }
}

function checkContractMatches(expectedContract: string, actual: WorkbookCheckResult): boolean {
  try {
    return expectedContract === canonicalJson(checkContract(actual))
  } catch {
    return false
  }
}

function isCheckStatus(value: unknown): value is WorkbookCheckResult['status'] {
  return value === 'planned' || value === 'passed' || value === 'failed'
}

function isWorkbookCheckResult(value: unknown): value is WorkbookCheckResult {
  if (!isRecord(value)) {
    return false
  }
  return (
    isCheckStatus(ownValue(value, 'status')) &&
    typeof ownValue(value, 'kind') === 'string' &&
    typeof ownValue(value, 'message') === 'string'
  )
}
