import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import {
  checkWorkbookCommandReceipt,
  normalizeWorkbookCommandReceipt,
  workbookCommandReceiptOpsMatch,
  workbookCommandReceiptStatuses,
  type WorkbookCommandReceipt,
  type WorkbookCommandReceiptStatus,
} from './features.js'
import { isWorkbookOp } from './guards.js'
import type { EngineOp } from './ops.js'
import type { WorkbookCommandBundle, WorkbookCommandBundleCommand } from './command-bundle.js'
import type { WorkbookUndoRef } from './result.js'

export type WorkbookCommandResultStatus = 'accepted' | WorkbookCommandReceiptStatus

export const workbookCommandResultStatuses = Object.freeze([
  'accepted',
  ...workbookCommandReceiptStatuses,
] as const satisfies readonly WorkbookCommandResultStatus[])

const WORKBOOK_COMMAND_RESULT_STATUS_SET = new Set<string>(workbookCommandResultStatuses)

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

interface NormalizedRangeWithSize {
  readonly range: CellRangeRef
  readonly cellCount: number
}

export function isWorkbookCommandResultStatus(value: unknown): value is WorkbookCommandResultStatus {
  return typeof value === 'string' && WORKBOOK_COMMAND_RESULT_STATUS_SET.has(value)
}

export function workbookCommandResultFor(bundle: WorkbookCommandBundle): WorkbookCommandResult {
  const touchedRanges: CellRangeRef[] = []
  let touchedCellCount = 0
  for (const command of bundle.commands) {
    for (const range of command.touchedRanges ?? []) {
      const normalized = normalizeRange(range, 'touchedRanges')
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
    assertReceiptMatchesCommand(bundle.commands[index]!, receipt, index)
  })

  const revision = options.revision
  if (revision !== undefined && !isSafeNonNegativeInteger(revision)) {
    throw new Error('Workbook command result is invalid: revision must be a safe non-negative integer')
  }
  const undo = options.undo === undefined ? undefined : normalizeUndoRef(options.undo, 'undo')
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
    return {
      status: 'invalid',
      issues: Object.freeze([commandResultIssue('invalid_command_result', 'result', 'Workbook command result must be an object')]),
    }
  }

  const accessorPath = firstAccessorPath(value, 'result')
  if (accessorPath !== null) {
    return {
      status: 'invalid',
      issues: Object.freeze([
        commandResultIssue('invalid_command_result', accessorPath, 'Workbook command result must contain only data properties'),
      ]),
    }
  }

  const issues: WorkbookCommandResultIssue[] = []
  const status = ownValue(value, 'status')
  if (!isWorkbookCommandResultStatus(status)) {
    issues.push(commandResultIssue('invalid_command_result', 'status', 'Workbook command result status is invalid'))
  }
  pushResultOptionalStringIssue(issues, ownValue(value, 'bundleId'), 'bundleId', 'bundle id')
  pushResultSafeIntegerIssue(issues, ownValue(value, 'targetRevision'), 'targetRevision', 'target revision', true)
  pushResultRequiredStringIssue(issues, ownValue(value, 'idempotencyKey'), 'idempotencyKey', 'idempotency key')
  pushResultSafeIntegerIssue(issues, ownValue(value, 'commandCount'), 'commandCount', 'command count', true)
  pushResultRangesIssues(issues, ownValue(value, 'touchedRanges'), 'touchedRanges', 'touched ranges')
  pushResultSafeIntegerIssue(issues, ownValue(value, 'touchedCellCount'), 'touchedCellCount', 'touched cell count', true)

  if (isWorkbookCommandResultStatus(status) && status !== 'accepted') {
    pushResultSafeIntegerIssue(issues, ownValue(value, 'revision'), 'revision', 'revision', false)
    pushResultReceiptsIssues(issues, ownValue(value, 'receipts'))
    const matched = ownValue(value, 'matched')
    if (matched !== null && typeof matched !== 'boolean') {
      issues.push(commandResultIssue('invalid_command_result', 'matched', 'Workbook command result matched must be boolean or null'))
    }
    pushResultRangesIssues(issues, ownValue(value, 'changedRanges'), 'changedRanges', 'changed ranges')
    const undo = ownValue(value, 'undo')
    if (undo !== undefined) {
      try {
        normalizeUndoRef(undo, 'undo')
      } catch (error) {
        issues.push(commandResultIssue('invalid_undo', 'undo', errorMessage(error)))
      }
    }
    pushResultErrorsIssues(issues, ownValue(value, 'errors'))
  }

  if (issues.length > 0) {
    return {
      status: 'invalid',
      issues: Object.freeze(issues),
    }
  }

  return {
    status: 'valid',
    result: normalizeWorkbookCommandResultData(value),
    issues: Object.freeze([]),
  }
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
    pushReceiptBundleIssues(issues, bundle, result)
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
  const accessorPath = firstAccessorPath(receipts, 'receipts')
  if (accessorPath !== null) {
    throw new Error(`Workbook command result is invalid: ${accessorPath} must contain only data properties`)
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
    }
  })
}

function rangesMatch(left: readonly CellRangeRef[], right: readonly CellRangeRef[]): boolean {
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
  return {
    status: 'invalid',
    issues: Object.freeze([...issues]),
  }
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
      normalizeRange(range, `${path}[${String(index)}]`, 'Workbook command result')
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
  const accessorPath = firstAccessorPath(value, path)
  if (accessorPath !== null) {
    issues.push(commandResultIssue('invalid_command_result', accessorPath, `${label} must contain only data properties`))
    return
  }
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
    ...(undo !== undefined ? { undo: normalizeUndoRef(undo, 'undo') } : {}),
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
  return value.map((range, index) => normalizeRange(range, `${path}[${String(index)}]`, 'Workbook command result').range)
}

function normalizeResultErrors(errors: readonly unknown[]): readonly string[] {
  return errors.map((entry, index) => {
    if (typeof entry !== 'string') {
      throw new Error(`Workbook command result errors[${String(index)}] must be a string`)
    }
    return normalizeExactString(entry, `errors[${String(index)}]`)
  })
}

function normalizeUndoRef(value: unknown, path: string): WorkbookUndoRef {
  if (!isRecord(value)) {
    throw new Error(`Workbook command result ${path} must be an object`)
  }
  const id = ownValue(value, 'id')
  if (typeof id !== 'string') {
    throw new Error(`Workbook command result ${path}.id must be a string`)
  }
  const ops = ownValue(value, 'ops')
  if (ops !== undefined && !Array.isArray(ops)) {
    throw new Error(`Workbook command result ${path}.ops must be an array`)
  }
  if (Array.isArray(ops)) {
    const accessorPath = firstAccessorPath(ops, `${path}.ops`)
    if (accessorPath !== null) {
      throw new Error(`Workbook command result ${accessorPath} must contain only data properties`)
    }
    const normalizedOps: EngineOp[] = []
    for (let index = 0; index < ops.length; index += 1) {
      const op = arrayDataValue(ops, index)
      if (op === undefined) {
        throw new Error(`Workbook command result ${path}.ops[${String(index)}] must contain only data properties`)
      }
      normalizedOps.push(normalizeOp(op))
    }
    return Object.freeze({
      id: normalizeExactString(id, `${path}.id`),
      ops: Object.freeze(normalizedOps),
    })
  }
  return Object.freeze({
    id: normalizeExactString(id, `${path}.id`),
  })
}

function normalizeRange(value: unknown, path: string, label = 'Workbook command bundle'): NormalizedRangeWithSize {
  if (!isRecord(value)) {
    throw new Error(`${label} ${path} must be an object`)
  }
  const sheetName = ownValue(value, 'sheetName')
  const startAddress = ownValue(value, 'startAddress')
  const endAddress = ownValue(value, 'endAddress')
  if (typeof sheetName !== 'string' || typeof startAddress !== 'string' || typeof endAddress !== 'string') {
    throw new Error(`${label} ${path} must include sheetName, startAddress, and endAddress strings`)
  }
  const normalizedSheetName = normalizeExactString(sheetName, `${path}.sheetName`)
  const start = normalizeCellAddress(startAddress, `${path}.startAddress`, label)
  const end = normalizeCellAddress(endAddress, `${path}.endAddress`, label)
  if (end.row < start.row || end.col < start.col) {
    throw new Error(`${label} ${path} endAddress must not be before startAddress`)
  }
  return {
    range: Object.freeze({
      sheetName: normalizedSheetName,
      startAddress: start.text,
      endAddress: end.text,
    }),
    cellCount: (end.row - start.row + 1) * (end.col - start.col + 1),
  }
}

function normalizeCellAddress(
  value: string,
  path: string,
  label = 'Workbook command bundle',
): { readonly row: number; readonly col: number; readonly text: string } {
  try {
    const parsed = parseCellAddress(value)
    if (parsed.sheetName !== undefined) {
      throw new Error('qualified')
    }
    return {
      row: parsed.row,
      col: parsed.col,
      text: formatAddress(parsed.row, parsed.col),
    }
  } catch {
    throw new Error(`${label} ${path} is invalid: ${value}`)
  }
}

function normalizeOp(value: unknown): EngineOp {
  if (!isWorkbookOp(value)) {
    throw new Error('Workbook command result op is invalid')
  }
  const cloned = cloneData(value)
  if (!isWorkbookOp(cloned)) {
    throw new Error('Workbook command result op clone is invalid')
  }
  return freezeData(cloned)
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function firstAccessorPath(value: unknown, path: string, seen = new WeakSet<object>()): string | null {
  if (typeof value !== 'object' || value === null) {
    return null
  }
  if (seen.has(value)) {
    return null
  }
  seen.add(value)

  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    const childPath = Array.isArray(value) && /^\d+$/.test(key) ? `${path}[${key}]` : `${path}.${key}`
    if (!('value' in descriptor)) {
      return childPath
    }
    const nestedPath = firstAccessorPath(descriptor.value, childPath, seen)
    if (nestedPath !== null) {
      return nestedPath
    }
  }
  return null
}

function freezeData<T>(value: T, seen = new WeakSet<object>()): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeData(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}

function cloneData(value: unknown, seen = new WeakMap<object, unknown>()): unknown {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  const existing = seen.get(value)
  if (existing !== undefined) {
    return existing
  }
  if (Array.isArray(value)) {
    const cloned: unknown[] = []
    seen.set(value, cloned)
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
      if (descriptor !== undefined && descriptor.enumerable && 'value' in descriptor) {
        cloned[index] = cloneData(descriptor.value, seen)
      }
    }
    return cloned
  }
  const cloned: Record<string, unknown> = Object.create(Object.getPrototypeOf(value))
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (descriptor.enumerable && 'value' in descriptor) {
      Object.defineProperty(cloned, key, {
        configurable: true,
        enumerable: true,
        value: cloneData(descriptor.value, seen),
        writable: true,
      })
    }
  })
  return cloned
}
