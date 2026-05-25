import type { CellRangeRef } from '@bilig/protocol'
import { commandRangeBounds, commandRangeContains, commandRangeLabel, normalizeCommandRange } from './command-ranges.js'
import { normalizeCommandResultUndoRef } from './command-result-undo.js'
import {
  checkWorkbookCommandReceipt,
  normalizeWorkbookCommandReceipt,
  workbookCommandReceiptOpsMatch,
  workbookCommandReceiptStatuses,
  type WorkbookCommandReceipt,
  type WorkbookCommandReceiptStatus,
} from './features.js'
import type { WorkbookCommandBundle, WorkbookCommandBundleCommand } from './command-bundle.js'
import type { WorkbookUndoRef } from './result.js'

export type WorkbookCommandResultStatus = 'accepted' | WorkbookCommandReceiptStatus

export const workbookCommandResultStatuses = Object.freeze([
  'accepted',
  ...workbookCommandReceiptStatuses,
] as const satisfies readonly WorkbookCommandResultStatus[])

const WORKBOOK_COMMAND_RESULT_STATUS_SET = new Set<string>(workbookCommandResultStatuses)

const acceptedResultSettledFields = Object.freeze(['receipts', 'matched', 'changedRanges', 'revision', 'undo', 'errors'] as const)
const commandResultDataFields = Object.freeze([
  'status',
  'bundleId',
  'targetRevision',
  'idempotencyKey',
  'commandCount',
  'touchedRanges',
  'touchedCellCount',
  ...acceptedResultSettledFields,
] as const)

export interface WorkbookCommandResultBase {
  readonly bundleId?: string
  readonly targetRevision: number
  readonly idempotencyKey: string
  readonly commandCount: number
  readonly touchedRanges: readonly CellRangeRef[]
  readonly touchedCellCount: number
}

export interface WorkbookCommandAcceptedResult extends WorkbookCommandResultBase {
  readonly status: 'accepted'
}

export interface WorkbookCommandSettledResult extends WorkbookCommandResultBase {
  readonly status: WorkbookCommandReceiptStatus
  readonly receipts: readonly WorkbookCommandReceipt[]
  readonly matched: boolean | null
  readonly changedRanges: readonly CellRangeRef[]
  readonly revision?: number
  readonly undo?: WorkbookUndoRef
  readonly errors?: readonly string[]
}

export type WorkbookCommandResult = WorkbookCommandAcceptedResult | WorkbookCommandSettledResult

export interface WorkbookCommandResultForReceiptsOptions {
  readonly revision?: number
  readonly undo?: WorkbookUndoRef
}

export const workbookOpCommandFeatureId = 'workbook-op'

export interface WorkbookOpCommandReceiptIdentity {
  readonly featureId: typeof workbookOpCommandFeatureId
  readonly commandId: string
  readonly category: 'operation'
}

export type WorkbookOpCommandReceiptOptions = Omit<WorkbookCommandReceipt, 'featureId' | 'commandId' | 'category'>

export type WorkbookCommandResultIssueCode =
  | 'invalid_command_result'
  | 'invalid_receipt'
  | 'bundle_result_mismatch'
  | 'changed_range_out_of_scope'
  | 'receipt_result_mismatch'
  | 'receipt_count_mismatch'
  | 'receipt_command_mismatch'
  | 'revision_mismatch'
  | 'invalid_undo'

export interface WorkbookCommandResultIssue {
  readonly code: WorkbookCommandResultIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookCommandResultCheckResult =
  | {
      readonly status: 'valid'
      readonly result: WorkbookCommandResult
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookCommandResultIssue[]
    }

export function isWorkbookCommandResultStatus(value: unknown): value is WorkbookCommandResultStatus {
  return typeof value === 'string' && WORKBOOK_COMMAND_RESULT_STATUS_SET.has(value)
}

export function workbookCommandResultFor(bundle: WorkbookCommandBundle): WorkbookCommandResult {
  const touchedRanges: CellRangeRef[] = []
  let touchedCellCount = 0
  for (const command of bundle.commands) {
    for (const range of command.touchedRanges ?? []) {
      const normalized = normalizeCommandRange(range, 'touchedRanges')
      touchedRanges.push(normalized.range)
      touchedCellCount += normalized.cellCount
    }
  }

  return Object.freeze({
    status: 'accepted',
    ...(bundle.id !== undefined ? { bundleId: bundle.id } : {}),
    targetRevision: bundle.targetRevision,
    idempotencyKey: bundle.idempotencyKey,
    commandCount: bundle.commands.length,
    touchedRanges: Object.freeze(touchedRanges),
    touchedCellCount,
  })
}

export function workbookCommandResultForReceipts(
  bundle: WorkbookCommandBundle,
  receipts: readonly WorkbookCommandReceipt[],
  options: WorkbookCommandResultForReceiptsOptions = {},
): WorkbookCommandResult {
  const normalizedReceipts = normalizeReceiptArrayForResult(receipts)
  if (normalizedReceipts.length !== bundle.commands.length) {
    throw new Error(
      `Workbook command result is invalid: expected ${String(bundle.commands.length)} receipts for ${String(bundle.commands.length)} commands, got ${String(normalizedReceipts.length)}`,
    )
  }
  normalizedReceipts.forEach((receipt, index) => {
    const command = bundle.commands[index]!
    assertReceiptMatchesCommand(command, receipt, index)
    assertReceiptChangedRangesInScope(command, receipt, index)
  })

  const revision = options.revision
  if (revision !== undefined && !isSafeNonNegativeInteger(revision)) {
    throw new Error('Workbook command result is invalid: revision must be a safe non-negative integer')
  }
  const undo = options.undo === undefined ? undefined : normalizeCommandResultUndoRef(options.undo, 'undo')
  const changedRanges = normalizedReceipts.flatMap((receipt) => [...(receipt.changedRanges ?? [])])
  const errors = commandResultErrorsForReceipts(normalizedReceipts)
  return normalizeWorkbookCommandResult({
    ...workbookCommandResultFor(bundle),
    status: commandResultStatusForReceipts(normalizedReceipts),
    ...(revision !== undefined ? { revision } : {}),
    receipts: normalizedReceipts,
    matched: commandResultMatchedForReceipts(normalizedReceipts),
    changedRanges,
    ...(undo !== undefined ? { undo } : {}),
    ...(errors.length > 0 ? { errors } : {}),
  })
}

export function workbookOpCommandReceiptIdentity(command: WorkbookCommandBundleCommand, index: number): WorkbookOpCommandReceiptIdentity {
  if (command.kind !== 'op') {
    throw new Error(`Workbook command bundle commands[${String(index)}] is not an op command`)
  }
  return Object.freeze({
    featureId: workbookOpCommandFeatureId,
    commandId: command.id ?? `commands[${String(index)}].op`,
    category: 'operation',
  })
}

export function workbookOpCommandReceipt(
  command: WorkbookCommandBundleCommand,
  index: number,
  options: WorkbookOpCommandReceiptOptions,
): WorkbookCommandReceipt {
  return normalizeWorkbookCommandReceipt({
    ...options,
    ...workbookOpCommandReceiptIdentity(command, index),
  })
}

export function checkWorkbookCommandResult(value: unknown): WorkbookCommandResultCheckResult {
  if (!isRecord(value)) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze([commandResultIssue('invalid_command_result', 'result', 'Workbook command result must be an object')]),
    })
  }

  const issues: WorkbookCommandResultIssue[] = []
  const accessorKeys = new Set(ownAccessorKeys(value, commandResultDataFields))
  accessorKeys.forEach((key) => {
    issues.push(commandResultIssue('invalid_command_result', key, `Workbook command result ${key} must be a data property`))
  })
  const status = accessorKeys.has('status') ? undefined : ownValue(value, 'status')
  if (!isWorkbookCommandResultStatus(status)) {
    issues.push(commandResultIssue('invalid_command_result', 'status', 'Workbook command result status is invalid'))
  }
  if (!accessorKeys.has('bundleId')) {
    pushResultOptionalStringIssue(issues, ownValue(value, 'bundleId'), 'bundleId', 'bundle id')
  }
  if (!accessorKeys.has('targetRevision')) {
    pushResultSafeIntegerIssue(issues, ownValue(value, 'targetRevision'), 'targetRevision', 'target revision', true)
  }
  if (!accessorKeys.has('idempotencyKey')) {
    pushResultRequiredStringIssue(issues, ownValue(value, 'idempotencyKey'), 'idempotencyKey', 'idempotency key')
  }
  if (!accessorKeys.has('commandCount')) {
    pushResultSafeIntegerIssue(issues, ownValue(value, 'commandCount'), 'commandCount', 'command count', true)
  }
  if (!accessorKeys.has('touchedRanges')) {
    pushResultRangesIssues(issues, ownValue(value, 'touchedRanges'), 'touchedRanges', 'touched ranges')
  }
  if (!accessorKeys.has('touchedCellCount')) {
    pushResultSafeIntegerIssue(issues, ownValue(value, 'touchedCellCount'), 'touchedCellCount', 'touched cell count', true)
  }

  if (status === 'accepted') {
    pushAcceptedResultSettledFieldIssues(issues, value)
  } else if (isWorkbookCommandResultStatus(status)) {
    if (!accessorKeys.has('revision')) {
      pushResultSafeIntegerIssue(issues, ownValue(value, 'revision'), 'revision', 'revision', false)
    }
    if (!accessorKeys.has('receipts')) {
      pushResultReceiptsIssues(issues, ownValue(value, 'receipts'))
    }
    if (!accessorKeys.has('matched')) {
      const matched = ownValue(value, 'matched')
      if (matched !== null && typeof matched !== 'boolean') {
        issues.push(commandResultIssue('invalid_command_result', 'matched', 'Workbook command result matched must be boolean or null'))
      }
    }
    if (!accessorKeys.has('changedRanges')) {
      pushResultRangesIssues(issues, ownValue(value, 'changedRanges'), 'changedRanges', 'changed ranges')
    }
    const undo = accessorKeys.has('undo') ? undefined : ownValue(value, 'undo')
    if (!accessorKeys.has('undo') && undo !== undefined) {
      const undoAccessorKeys = isRecord(undo) ? ownAccessorKeys(undo, ['id', 'ops']) : []
      undoAccessorKeys.forEach((key) => {
        issues.push(commandResultIssue('invalid_undo', `undo.${key}`, `Workbook command result undo.${key} must be a data property`))
      })
      if (undoAccessorKeys.length === 0) {
        try {
          normalizeCommandResultUndoRef(undo, 'undo')
        } catch (error) {
          issues.push(commandResultIssue('invalid_undo', 'undo', errorMessage(error)))
        }
      }
    }
    if (!accessorKeys.has('errors')) {
      pushResultErrorsIssues(issues, ownValue(value, 'errors'))
    }
  }

  if (issues.length > 0) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze(issues),
    })
  }

  return Object.freeze({
    status: 'valid',
    result: normalizeWorkbookCommandResultData(value),
    issues: Object.freeze([] as const),
  })
}

export function normalizeWorkbookCommandResult(value: unknown): WorkbookCommandResult {
  const check = checkWorkbookCommandResult(value)
  if (check.status === 'invalid') {
    const [firstIssue] = check.issues
    throw new Error(
      firstIssue === undefined ? 'Workbook command result is invalid' : `Workbook command result is invalid: ${firstIssue.message}`,
    )
  }
  return check.result
}

export function isWorkbookCommandResult(value: unknown): value is WorkbookCommandResult {
  return checkWorkbookCommandResult(value).status === 'valid'
}

export function checkWorkbookCommandResultForBundle(bundle: WorkbookCommandBundle, value: unknown): WorkbookCommandResultCheckResult {
  const resultCheck = checkWorkbookCommandResult(value)
  if (resultCheck.status === 'invalid') {
    return resultCheck
  }

  let expected: WorkbookCommandResult
  try {
    expected = workbookCommandResultFor(bundle)
  } catch (error) {
    return invalidCommandResult([
      commandResultIssue('invalid_command_result', 'bundle', `Workbook command bundle is invalid: ${errorMessage(error)}`),
    ])
  }

  const result = resultCheck.result
  const issues: WorkbookCommandResultIssue[] = []
  pushBaseResultMismatchIssues(issues, expected, result)
  if (result.status !== 'accepted') {
    const receiptIssueStart = issues.length
    pushReceiptBundleIssues(issues, bundle, result)
    if (issues.length === receiptIssueStart) {
      pushReceiptDerivedResultIssues(issues, bundle, result)
    }
    if (result.status === 'applied' && result.revision === undefined) {
      issues.push(commandResultIssue('revision_mismatch', 'revision', 'Applied workbook command result must include a revision'))
    }
    if (result.revision !== undefined && result.revision < result.targetRevision) {
      issues.push(commandResultIssue('revision_mismatch', 'revision', 'Workbook command result revision must not be before targetRevision'))
    }
  }

  if (issues.length > 0) {
    return invalidCommandResult(issues)
  }
  return resultCheck
}

export function isWorkbookCommandResultForBundle(bundle: WorkbookCommandBundle, value: unknown): value is WorkbookCommandResult {
  return checkWorkbookCommandResultForBundle(bundle, value).status === 'valid'
}

function normalizeReceiptArrayForResult(receipts: readonly WorkbookCommandReceipt[]): readonly WorkbookCommandReceipt[] {
  if (!Array.isArray(receipts)) {
    throw new Error('Workbook command result is invalid: receipts must be an array')
  }
  const normalizedReceipts: WorkbookCommandReceipt[] = []
  for (let index = 0; index < receipts.length; index += 1) {
    const receipt = arrayDataValue(receipts, index)
    if (receipt === undefined) {
      throw new Error(`Workbook command result is invalid: receipts[${String(index)}] must contain only data properties`)
    }
    try {
      normalizedReceipts.push(normalizeWorkbookCommandReceipt(receipt))
    } catch (error) {
      throw new Error(`Workbook command result is invalid: receipts[${String(index)}] ${errorMessage(error)}`, { cause: error })
    }
  }
  return Object.freeze(normalizedReceipts)
}

function assertReceiptMatchesCommand(command: WorkbookCommandBundleCommand, receipt: WorkbookCommandReceipt, index: number): void {
  if (command.kind !== 'request') {
    const expected = workbookOpCommandReceiptIdentity(command, index)
    if (receipt.featureId !== expected.featureId || receipt.commandId !== expected.commandId || receipt.category !== expected.category) {
      throw new Error(`Workbook command result is invalid: receipts[${String(index)}] does not match commands[${String(index)}].op`)
    }
    const opProofIssue = opReceiptProofIssue(command, receipt, index)
    if (opProofIssue !== null) {
      throw new Error(`Workbook command result is invalid: ${opProofIssue.message}`)
    }
    return
  }
  if (receipt.featureId !== command.request.featureId || receipt.commandId !== command.request.commandId) {
    throw new Error(`Workbook command result is invalid: receipts[${String(index)}] does not match commands[${String(index)}].request`)
  }
  if (command.request.category !== undefined && receipt.category !== command.request.category) {
    throw new Error(
      `Workbook command result is invalid: receipts[${String(index)}].category does not match commands[${String(index)}].request.category`,
    )
  }
}

function pushBaseResultMismatchIssues(
  issues: WorkbookCommandResultIssue[],
  expected: WorkbookCommandResult,
  result: WorkbookCommandResult,
): void {
  if (result.bundleId !== expected.bundleId) {
    issues.push(commandResultIssue('bundle_result_mismatch', 'bundleId', 'Workbook command result bundleId does not match bundle'))
  }
  if (result.targetRevision !== expected.targetRevision) {
    issues.push(
      commandResultIssue('bundle_result_mismatch', 'targetRevision', 'Workbook command result targetRevision does not match bundle'),
    )
  }
  if (result.idempotencyKey !== expected.idempotencyKey) {
    issues.push(
      commandResultIssue('bundle_result_mismatch', 'idempotencyKey', 'Workbook command result idempotencyKey does not match bundle'),
    )
  }
  if (result.commandCount !== expected.commandCount) {
    issues.push(commandResultIssue('bundle_result_mismatch', 'commandCount', 'Workbook command result commandCount does not match bundle'))
  }
  if (result.touchedCellCount !== expected.touchedCellCount) {
    issues.push(
      commandResultIssue('bundle_result_mismatch', 'touchedCellCount', 'Workbook command result touchedCellCount does not match bundle'),
    )
  }
  if (!rangesMatch(result.touchedRanges, expected.touchedRanges)) {
    issues.push(commandResultIssue('bundle_result_mismatch', 'touchedRanges', 'Workbook command result touchedRanges do not match bundle'))
  }
}

function pushReceiptBundleIssues(
  issues: WorkbookCommandResultIssue[],
  bundle: WorkbookCommandBundle,
  result: WorkbookCommandSettledResult,
): void {
  if (result.receipts.length !== bundle.commands.length) {
    issues.push(
      commandResultIssue(
        'receipt_count_mismatch',
        'receipts',
        `Workbook command result has ${String(result.receipts.length)} receipts for ${String(bundle.commands.length)} commands`,
      ),
    )
    return
  }

  result.receipts.forEach((receipt, index) => {
    const command = bundle.commands[index]
    if (command === undefined) {
      return
    }
    if (command.kind === 'op') {
      const expected = workbookOpCommandReceiptIdentity(command, index)
      if (receipt.featureId !== expected.featureId || receipt.commandId !== expected.commandId || receipt.category !== expected.category) {
        issues.push(
          commandResultIssue(
            'receipt_command_mismatch',
            `receipts[${String(index)}]`,
            `Workbook command result receipt ${String(index)} does not match op command`,
          ),
        )
      }
      pushOpReceiptProofIssues(issues, command, receipt, index)
      pushReceiptChangedRangeScopeIssues(issues, command, receipt, index)
      return
    }
    if (receipt.featureId !== command.request.featureId || receipt.commandId !== command.request.commandId) {
      issues.push(
        commandResultIssue(
          'receipt_command_mismatch',
          `receipts[${String(index)}]`,
          `Workbook command result receipt ${String(index)} does not match command request`,
        ),
      )
      return
    }
    if (command.request.category !== undefined && receipt.category !== command.request.category) {
      issues.push(
        commandResultIssue(
          'receipt_command_mismatch',
          `receipts[${String(index)}].category`,
          `Workbook command result receipt ${String(index)} category does not match command request`,
        ),
      )
      return
    }
    pushReceiptChangedRangeScopeIssues(issues, command, receipt, index)
  })
}

function receiptOpsMatchCommandOp(
  ops: readonly unknown[] | undefined,
  command: Extract<WorkbookCommandBundleCommand, { readonly kind: 'op' }>,
): boolean {
  if (ops === undefined || ops.length !== 1) {
    return false
  }
  const [op] = ops
  try {
    return canonicalJson(op) === canonicalJson(command.op)
  } catch {
    return false
  }
}

function opReceiptProofIssue(
  command: Extract<WorkbookCommandBundleCommand, { readonly kind: 'op' }>,
  receipt: WorkbookCommandReceipt,
  index: number,
): WorkbookCommandResultIssue | null {
  if (receipt.previewOps !== undefined && !receiptOpsMatchCommandOp(receipt.previewOps, command)) {
    return commandResultIssue(
      'receipt_command_mismatch',
      `receipts[${String(index)}].previewOps`,
      `Workbook command result receipt ${String(index)} previewOps must equal commands[${String(index)}].op`,
    )
  }
  if (receipt.status === 'previewed' && !receiptOpsMatchCommandOp(receipt.previewOps, command)) {
    return commandResultIssue(
      'receipt_command_mismatch',
      `receipts[${String(index)}].previewOps`,
      `Workbook command result receipt ${String(index)} previewOps must equal commands[${String(index)}].op`,
    )
  }
  if (receipt.appliedOps !== undefined && !receiptOpsMatchCommandOp(receipt.appliedOps, command)) {
    return commandResultIssue(
      'receipt_command_mismatch',
      `receipts[${String(index)}].appliedOps`,
      `Workbook command result receipt ${String(index)} appliedOps must equal commands[${String(index)}].op`,
    )
  }
  if (receipt.status === 'applied' && !receiptOpsMatchCommandOp(receipt.appliedOps, command)) {
    return commandResultIssue(
      'receipt_command_mismatch',
      `receipts[${String(index)}].appliedOps`,
      `Workbook command result receipt ${String(index)} appliedOps must equal commands[${String(index)}].op`,
    )
  }
  return null
}

function pushOpReceiptProofIssues(
  issues: WorkbookCommandResultIssue[],
  command: Extract<WorkbookCommandBundleCommand, { readonly kind: 'op' }>,
  receipt: WorkbookCommandReceipt,
  index: number,
): void {
  const issue = opReceiptProofIssue(command, receipt, index)
  if (issue !== null) {
    issues.push(issue)
  }
}

function assertReceiptChangedRangesInScope(command: WorkbookCommandBundleCommand, receipt: WorkbookCommandReceipt, index: number): void {
  const issue = firstReceiptChangedRangeScopeIssue(command, receipt, index)
  if (issue !== null) {
    throw new Error(`Workbook command result is invalid: ${issue.message}`)
  }
}

function pushReceiptChangedRangeScopeIssues(
  issues: WorkbookCommandResultIssue[],
  command: WorkbookCommandBundleCommand,
  receipt: WorkbookCommandReceipt,
  index: number,
): void {
  const issue = firstReceiptChangedRangeScopeIssue(command, receipt, index)
  if (issue !== null) {
    issues.push(issue)
  }
}

function firstReceiptChangedRangeScopeIssue(
  command: WorkbookCommandBundleCommand,
  receipt: WorkbookCommandReceipt,
  index: number,
): WorkbookCommandResultIssue | null {
  const declaredRanges = command.touchedRanges ?? []
  const changedRanges = receipt.changedRanges ?? []
  if (declaredRanges.length === 0 || changedRanges.length === 0) {
    return null
  }

  const declared = declaredRanges.map((range, rangeIndex) =>
    commandRangeBounds(range, `commands[${String(index)}].touchedRanges[${String(rangeIndex)}]`, 'Workbook command result'),
  )
  for (let rangeIndex = 0; rangeIndex < changedRanges.length; rangeIndex += 1) {
    const changedRange = changedRanges[rangeIndex]
    if (changedRange === undefined) {
      continue
    }
    const changed = commandRangeBounds(
      changedRange,
      `receipts[${String(index)}].changedRanges[${String(rangeIndex)}]`,
      'Workbook command result',
    )
    if (!declared.some((candidate) => commandRangeContains(candidate, changed))) {
      return commandResultIssue(
        'changed_range_out_of_scope',
        `receipts[${String(index)}].changedRanges[${String(rangeIndex)}]`,
        `Workbook command result receipt ${String(index)} changed range ${commandRangeLabel(changed)} is outside commands[${String(index)}].touchedRanges`,
      )
    }
  }
  return null
}

function pushReceiptDerivedResultIssues(
  issues: WorkbookCommandResultIssue[],
  bundle: WorkbookCommandBundle,
  result: WorkbookCommandSettledResult,
): void {
  if (result.receipts.length !== bundle.commands.length) {
    return
  }

  let derived: WorkbookCommandResult
  try {
    derived = workbookCommandResultForReceipts(bundle, result.receipts, {
      ...(result.revision !== undefined ? { revision: result.revision } : {}),
      ...(result.undo !== undefined ? { undo: result.undo } : {}),
    })
  } catch (error) {
    issues.push(commandResultIssue('invalid_receipt', 'receipts', errorMessage(error)))
    return
  }
  if (derived.status === 'accepted') {
    return
  }

  if (result.status !== derived.status) {
    issues.push(commandResultIssue('receipt_result_mismatch', 'status', 'Workbook command result status does not match receipts'))
  }
  if (result.matched !== derived.matched) {
    issues.push(commandResultIssue('receipt_result_mismatch', 'matched', 'Workbook command result matched does not match receipts'))
  }
  if (!rangesMatch(result.changedRanges, derived.changedRanges)) {
    issues.push(
      commandResultIssue('receipt_result_mismatch', 'changedRanges', 'Workbook command result changedRanges do not match receipts'),
    )
  }
  if (!stringArraysMatch(result.errors ?? [], derived.errors ?? [])) {
    issues.push(commandResultIssue('receipt_result_mismatch', 'errors', 'Workbook command result errors do not match receipts'))
  }
}

function rangesMatch(left: readonly CellRangeRef[], right: readonly CellRangeRef[]): boolean {
  try {
    return canonicalJson(left) === canonicalJson(right)
  } catch {
    return false
  }
}

function stringArraysMatch(left: readonly string[], right: readonly string[]): boolean {
  try {
    return canonicalJson(left) === canonicalJson(right)
  } catch {
    return false
  }
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
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

function commandResultStatusForReceipts(receipts: readonly WorkbookCommandReceipt[]): WorkbookCommandReceiptStatus {
  if (receipts.some((receipt) => receipt.status === 'rejected')) {
    return 'rejected'
  }
  if (receipts.some((receipt) => receipt.status === 'applied')) {
    return 'applied'
  }
  if (receipts.some((receipt) => receipt.status === 'previewed')) {
    return 'previewed'
  }
  return 'noop'
}

function commandResultMatchedForReceipts(receipts: readonly WorkbookCommandReceipt[]): boolean | null {
  let matched: boolean | null = null
  for (const receipt of receipts) {
    const receiptMatched = workbookCommandReceiptOpsMatch(receipt)
    if (receiptMatched === false) {
      return false
    }
    if (receiptMatched === true) {
      matched = true
    }
  }
  return matched
}

function commandResultErrorsForReceipts(receipts: readonly WorkbookCommandReceipt[]): readonly string[] {
  const errors: string[] = []
  for (const receipt of receipts) {
    errors.push(...(receipt.errors ?? []))
    if (receipt.status === 'rejected' && receipt.errors === undefined) {
      errors.push(receipt.message ?? `${receipt.featureId}.${receipt.commandId} rejected`)
    }
  }
  return Object.freeze(errors)
}

function commandResultIssue(code: WorkbookCommandResultIssueCode, path: string, message: string): WorkbookCommandResultIssue {
  return Object.freeze({
    code,
    path,
    message,
  })
}

function invalidCommandResult(issues: readonly WorkbookCommandResultIssue[]): WorkbookCommandResultCheckResult {
  return Object.freeze({
    status: 'invalid',
    issues: Object.freeze([...issues]),
  })
}

function pushResultReceiptsIssues(issues: WorkbookCommandResultIssue[], value: unknown): void {
  if (!Array.isArray(value) || value.length === 0) {
    issues.push(commandResultIssue('invalid_command_result', 'receipts', 'Workbook command result receipts must be a non-empty array'))
    return
  }
  pushResultArrayDataIssues(issues, value, 'receipts', 'Workbook command result receipts')
  for (let index = 0; index < value.length; index += 1) {
    const receipt = arrayDataValue(value, index)
    if (receipt === undefined) {
      continue
    }
    const receiptCheck = checkWorkbookCommandReceipt(receipt)
    if (receiptCheck.status === 'invalid') {
      receiptCheck.issues.forEach((issue) => {
        issues.push(commandResultIssue('invalid_receipt', `receipts[${String(index)}].${issue.path}`, issue.message))
      })
    }
  }
}

function pushAcceptedResultSettledFieldIssues(issues: WorkbookCommandResultIssue[], value: Record<string, unknown>): void {
  for (const field of acceptedResultSettledFields) {
    if (Object.hasOwn(value, field)) {
      issues.push(commandResultIssue('invalid_command_result', field, `Accepted workbook command result must not include ${field}`))
    }
  }
}

function pushResultRangesIssues(issues: WorkbookCommandResultIssue[], value: unknown, path: string, label: string): void {
  if (!Array.isArray(value)) {
    issues.push(commandResultIssue('invalid_command_result', path, `Workbook command result ${label} must be an array`))
    return
  }
  pushResultArrayDataIssues(issues, value, path, `Workbook command result ${label}`)
  for (let index = 0; index < value.length; index += 1) {
    const range = arrayDataValue(value, index)
    if (range === undefined) {
      continue
    }
    try {
      normalizeCommandRange(range, `${path}[${String(index)}]`, 'Workbook command result')
    } catch (error) {
      issues.push(commandResultIssue('invalid_command_result', `${path}[${String(index)}]`, errorMessage(error)))
    }
  }
}

function pushResultErrorsIssues(issues: WorkbookCommandResultIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  if (!Array.isArray(value)) {
    issues.push(commandResultIssue('invalid_command_result', 'errors', 'Workbook command result errors must be an array'))
    return
  }
  pushResultArrayDataIssues(issues, value, 'errors', 'Workbook command result errors')
  for (let index = 0; index < value.length; index += 1) {
    const error = arrayDataValue(value, index)
    if (error === undefined) {
      continue
    }
    pushResultOptionalStringIssue(issues, error, `errors[${String(index)}]`, 'error')
  }
}

function pushResultArrayDataIssues(issues: WorkbookCommandResultIssue[], value: readonly unknown[], path: string, label: string): void {
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      issues.push(commandResultIssue('invalid_command_result', `${path}[${String(index)}]`, `${label} must contain only data properties`))
    }
  }
}

function pushResultSafeIntegerIssue(
  issues: WorkbookCommandResultIssue[],
  value: unknown,
  path: string,
  label: string,
  required: boolean,
): void {
  if (value === undefined) {
    if (required) {
      issues.push(commandResultIssue('invalid_command_result', path, `Workbook command result ${label} is required`))
    }
    return
  }
  if (!isSafeNonNegativeInteger(value)) {
    issues.push(commandResultIssue('invalid_command_result', path, `Workbook command result ${label} must be a safe non-negative integer`))
  }
}

function pushResultRequiredStringIssue(issues: WorkbookCommandResultIssue[], value: unknown, path: string, label: string): void {
  if (typeof value !== 'string') {
    issues.push(commandResultIssue('invalid_command_result', path, `Workbook command result ${label} must be a string`))
    return
  }
  pushResultOptionalStringIssue(issues, value, path, label)
}

function pushResultOptionalStringIssue(issues: WorkbookCommandResultIssue[], value: unknown, path: string, label: string): void {
  if (value === undefined) {
    return
  }
  if (typeof value !== 'string') {
    issues.push(commandResultIssue('invalid_command_result', path, `Workbook command result ${label} must be a string`))
    return
  }
  const normalized = value.trim()
  if (normalized === '') {
    issues.push(commandResultIssue('invalid_command_result', path, `Workbook command result ${label} cannot be empty`))
    return
  }
  if (normalized !== value) {
    issues.push(
      commandResultIssue('invalid_command_result', path, `Workbook command result ${label} must not have leading or trailing whitespace`),
    )
  }
}

function normalizeWorkbookCommandResultData(value: Record<string, unknown>): WorkbookCommandResult {
  const status = ownValue(value, 'status')
  if (!isWorkbookCommandResultStatus(status)) {
    throw new Error('Workbook command result status is invalid')
  }
  const bundleId = ownValue(value, 'bundleId')
  const targetRevision = ownValue(value, 'targetRevision')
  const idempotencyKey = ownValue(value, 'idempotencyKey')
  const commandCount = ownValue(value, 'commandCount')
  const touchedCellCount = ownValue(value, 'touchedCellCount')
  if (
    !isSafeNonNegativeInteger(targetRevision) ||
    typeof idempotencyKey !== 'string' ||
    !isSafeNonNegativeInteger(commandCount) ||
    !isSafeNonNegativeInteger(touchedCellCount)
  ) {
    throw new Error('Workbook command result base is invalid')
  }
  const base = {
    ...(typeof bundleId === 'string' ? { bundleId: normalizeExactString(bundleId, 'bundleId') } : {}),
    targetRevision,
    idempotencyKey: normalizeExactString(idempotencyKey, 'idempotencyKey'),
    commandCount,
    touchedRanges: Object.freeze(normalizeResultRanges(ownValue(value, 'touchedRanges'), 'touchedRanges')),
    touchedCellCount,
  }

  if (status === 'accepted') {
    return Object.freeze({
      status,
      ...base,
    })
  }

  const matched = ownValue(value, 'matched')
  const revision = ownValue(value, 'revision')
  const undo = ownValue(value, 'undo')
  const errors = ownValue(value, 'errors')
  return Object.freeze({
    status,
    ...base,
    ...(isSafeNonNegativeInteger(revision) ? { revision } : {}),
    receipts: Object.freeze(normalizeResultReceipts(ownValue(value, 'receipts'))),
    matched: matched === true ? true : matched === false ? false : null,
    changedRanges: Object.freeze(normalizeResultRanges(ownValue(value, 'changedRanges'), 'changedRanges')),
    ...(undo !== undefined ? { undo: normalizeCommandResultUndoRef(undo, 'undo') } : {}),
    ...(Array.isArray(errors) ? { errors: Object.freeze(normalizeResultErrors(errors)) } : {}),
  })
}

function normalizeResultReceipts(value: unknown): readonly WorkbookCommandReceipt[] {
  if (!Array.isArray(value)) {
    throw new Error('Workbook command result receipts are invalid')
  }
  return value.map((receipt) => normalizeWorkbookCommandReceipt(receipt))
}

function normalizeResultRanges(value: unknown, path: string): readonly CellRangeRef[] {
  if (!Array.isArray(value)) {
    throw new Error(`Workbook command result ${path} is invalid`)
  }
  return value.map((range, index) => normalizeCommandRange(range, `${path}[${String(index)}]`, 'Workbook command result').range)
}

function normalizeResultErrors(errors: readonly unknown[]): readonly string[] {
  return errors.map((entry, index) => {
    if (typeof entry !== 'string') {
      throw new Error(`Workbook command result errors[${String(index)}] must be a string`)
    }
    return normalizeExactString(entry, `errors[${String(index)}]`)
  })
}

function normalizeExactString(value: string, path: string): string {
  const normalized = value.trim()
  if (normalized === '') {
    throw new Error(`Workbook command result ${path} cannot be empty`)
  }
  if (normalized !== value) {
    throw new Error(`Workbook command result ${path} must not have leading or trailing whitespace`)
  }
  return normalized
}

function arrayDataValue(value: readonly unknown[], index: number): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
  return descriptor !== undefined && descriptor.enumerable && 'value' in descriptor ? descriptor.value : undefined
}

function ownValue(value: Record<string, unknown>, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function ownAccessorKeys(value: Record<string, unknown>, keys: readonly string[]): readonly string[] {
  const accessors: string[] = []
  for (const key of keys) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    if (descriptor !== undefined && !('value' in descriptor)) {
      accessors.push(key)
    }
  }
  return accessors
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
