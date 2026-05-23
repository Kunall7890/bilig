import type { CellRangeRef } from '@bilig/protocol'
import { isWorkbookOp } from './guards.js'
import { WorkbookActionInputError, isWorkbookActionInput, normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import type { EngineOp } from './ops.js'
import type { WorkbookUndoRef } from './result.js'

export type WorkbookFeatureId = string
export type WorkbookCommandCategory = 'command' | 'operation' | 'mutation'
export type WorkbookCommandExecutionMode = 'preview' | 'apply' | 'applyAndVerify'
export type WorkbookCommandReceiptStatus = 'previewed' | 'applied' | 'rejected' | 'noop'
export type WorkbookProjectionInterceptorPoint =
  | 'cellDisplay'
  | 'cellStyle'
  | 'rangeChrome'
  | 'rowVisibility'
  | 'beforeCommand'
  | 'commandMetadata'
export type WorkbookUiContributionSlot = 'toolbar' | 'sidePanel' | 'floatingOverlay' | 'status'

export const workbookCommandCategories = Object.freeze([
  'command',
  'operation',
  'mutation',
] as const satisfies readonly WorkbookCommandCategory[])
const WORKBOOK_COMMAND_CATEGORY_SET = new Set<string>(workbookCommandCategories)

export const workbookCommandExecutionModes = Object.freeze([
  'preview',
  'apply',
  'applyAndVerify',
] as const satisfies readonly WorkbookCommandExecutionMode[])
const WORKBOOK_COMMAND_EXECUTION_MODE_SET = new Set<string>(workbookCommandExecutionModes)

export const workbookCommandReceiptStatuses = Object.freeze([
  'previewed',
  'applied',
  'rejected',
  'noop',
] as const satisfies readonly WorkbookCommandReceiptStatus[])
const WORKBOOK_COMMAND_RECEIPT_STATUS_SET = new Set<string>(workbookCommandReceiptStatuses)

export const workbookProjectionInterceptorPoints = Object.freeze([
  'cellDisplay',
  'cellStyle',
  'rangeChrome',
  'rowVisibility',
  'beforeCommand',
  'commandMetadata',
] as const satisfies readonly WorkbookProjectionInterceptorPoint[])
const WORKBOOK_PROJECTION_INTERCEPTOR_POINT_SET = new Set<string>(workbookProjectionInterceptorPoints)

export const workbookUiContributionSlots = Object.freeze([
  'toolbar',
  'sidePanel',
  'floatingOverlay',
  'status',
] as const satisfies readonly WorkbookUiContributionSlot[])
const WORKBOOK_UI_CONTRIBUTION_SLOT_SET = new Set<string>(workbookUiContributionSlots)

export interface WorkbookCommandRequest {
  readonly featureId: WorkbookFeatureId
  readonly commandId: string
  readonly category?: WorkbookCommandCategory
  readonly mode?: WorkbookCommandExecutionMode
  readonly input?: WorkbookActionInput
}

export type WorkbookCommandRequestIssueCode = 'invalid_command_request'

export interface WorkbookCommandRequestIssue {
  readonly code: WorkbookCommandRequestIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookCommandRequestCheckResult =
  | {
      readonly status: 'valid'
      readonly request: WorkbookCommandRequest
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookCommandRequestIssue[]
    }

export interface WorkbookCommandReceipt {
  readonly status: WorkbookCommandReceiptStatus
  readonly featureId: WorkbookFeatureId
  readonly commandId: string
  readonly category: WorkbookCommandCategory
  readonly previewOps?: readonly EngineOp[]
  readonly appliedOps?: readonly EngineOp[]
  readonly undo?: WorkbookUndoRef
  readonly changedRanges?: readonly CellRangeRef[]
  readonly proof?: WorkbookActionInput
  readonly message?: string
  readonly metadata?: WorkbookActionInput
  readonly errors?: readonly string[]
}

export type WorkbookCommandReceiptIssueCode = 'invalid_command_receipt'

export interface WorkbookCommandReceiptIssue {
  readonly code: WorkbookCommandReceiptIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookCommandReceiptCheckResult =
  | {
      readonly status: 'valid'
      readonly receipt: WorkbookCommandReceipt
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookCommandReceiptIssue[]
    }

export function normalizeWorkbookFeatureId(value: string, label = 'Workbook feature id'): WorkbookFeatureId {
  const normalized = value.trim()
  if (normalized === '') {
    throw new Error(`${label} cannot be empty`)
  }
  if (normalized !== value) {
    throw new Error(`${label} must not have leading or trailing whitespace`)
  }
  return normalized
}

export function isWorkbookCommandCategory(value: unknown): value is WorkbookCommandCategory {
  return typeof value === 'string' && WORKBOOK_COMMAND_CATEGORY_SET.has(value)
}

export function isWorkbookCommandExecutionMode(value: unknown): value is WorkbookCommandExecutionMode {
  return typeof value === 'string' && WORKBOOK_COMMAND_EXECUTION_MODE_SET.has(value)
}

export function isWorkbookCommandReceiptStatus(value: unknown): value is WorkbookCommandReceiptStatus {
  return typeof value === 'string' && WORKBOOK_COMMAND_RECEIPT_STATUS_SET.has(value)
}

export function isWorkbookProjectionInterceptorPoint(value: unknown): value is WorkbookProjectionInterceptorPoint {
  return typeof value === 'string' && WORKBOOK_PROJECTION_INTERCEPTOR_POINT_SET.has(value)
}

export function isWorkbookUiContributionSlot(value: unknown): value is WorkbookUiContributionSlot {
  return typeof value === 'string' && WORKBOOK_UI_CONTRIBUTION_SLOT_SET.has(value)
}

function commandRequestIssue(path: string, message: string): WorkbookCommandRequestIssue {
  return Object.freeze({
    code: 'invalid_command_request',
    path,
    message,
  })
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function inputPath(basePath: string, error: unknown): string {
  if (!(error instanceof WorkbookActionInputError)) {
    return basePath
  }
  if (error.path === 'input') {
    return basePath
  }
  if (error.path.startsWith('input.')) {
    return `${basePath}${error.path.slice('input'.length)}`
  }
  if (error.path.startsWith('input[')) {
    return `${basePath}${error.path.slice('input'.length)}`
  }
  return basePath
}

function ownValue(value: object, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function pushCommandRequestInputIssue(issues: WorkbookCommandRequestIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  try {
    normalizeWorkbookActionInput(value)
  } catch (error) {
    issues.push(commandRequestIssue(inputPath('input', error), `Workbook command request input is invalid: ${errorMessage(error)}`))
  }
}

function pushRequiredCommandRequestStringIssue(
  issues: WorkbookCommandRequestIssue[],
  value: Record<string, unknown>,
  key: 'featureId' | 'commandId',
  label: string,
): void {
  const entry = ownValue(value, key)
  if (typeof entry !== 'string') {
    issues.push(commandRequestIssue(key, `Workbook command request ${label} must be a string`))
    return
  }
  const normalized = entry.trim()
  if (normalized === '') {
    issues.push(commandRequestIssue(key, `Workbook command request ${label} cannot be empty`))
    return
  }
  if (normalized !== entry) {
    issues.push(commandRequestIssue(key, `Workbook command request ${label} must not have leading or trailing whitespace`))
  }
}

function normalizedCommandRequest(value: unknown): WorkbookCommandRequest | null {
  if (!isRecord(value)) {
    return null
  }
  const featureId = ownValue(value, 'featureId')
  const commandId = ownValue(value, 'commandId')
  const category = ownValue(value, 'category')
  const mode = ownValue(value, 'mode')
  const input = ownValue(value, 'input')
  if (typeof featureId !== 'string' || typeof commandId !== 'string') {
    return null
  }
  if (
    (category !== undefined && !isWorkbookCommandCategory(category)) ||
    (mode !== undefined && !isWorkbookCommandExecutionMode(mode)) ||
    (input !== undefined && !isWorkbookActionInput(input))
  ) {
    return null
  }
  return Object.freeze({
    featureId: normalizeWorkbookFeatureId(featureId, 'Workbook command request feature id'),
    commandId: normalizeRequiredString(commandId, 'Workbook command request command id'),
    ...(category !== undefined ? { category } : {}),
    ...(mode !== undefined ? { mode } : {}),
    ...(input !== undefined ? { input: normalizeWorkbookActionInput(input) } : {}),
  })
}

export function checkWorkbookCommandRequest(value: unknown): WorkbookCommandRequestCheckResult {
  if (!isRecord(value)) {
    return {
      status: 'invalid',
      issues: Object.freeze([commandRequestIssue('request', 'Workbook command request must be an object')]),
    }
  }

  const issues: WorkbookCommandRequestIssue[] = []
  pushRequiredCommandRequestStringIssue(issues, value, 'featureId', 'feature id')
  pushRequiredCommandRequestStringIssue(issues, value, 'commandId', 'command id')
  const category = ownValue(value, 'category')
  const mode = ownValue(value, 'mode')
  const input = ownValue(value, 'input')
  if (category !== undefined && !isWorkbookCommandCategory(category)) {
    issues.push(commandRequestIssue('category', 'Workbook command request category is invalid'))
  }
  if (mode !== undefined && !isWorkbookCommandExecutionMode(mode)) {
    issues.push(commandRequestIssue('mode', 'Workbook command request mode is invalid'))
  }
  pushCommandRequestInputIssue(issues, input)

  if (issues.length > 0) {
    return {
      status: 'invalid',
      issues: Object.freeze(issues),
    }
  }
  const request = normalizedCommandRequest(value)
  if (request === null) {
    return {
      status: 'invalid',
      issues: Object.freeze([commandRequestIssue('request', 'Workbook command request is invalid')]),
    }
  }
  return {
    status: 'valid',
    request,
    issues: Object.freeze([]),
  }
}

export function normalizeWorkbookCommandRequest(value: unknown): WorkbookCommandRequest {
  const check = checkWorkbookCommandRequest(value)
  if (check.status === 'invalid') {
    const [firstIssue] = check.issues
    throw new Error(
      firstIssue === undefined ? 'Workbook command request is invalid' : `Workbook command request is invalid: ${firstIssue.message}`,
    )
  }
  return check.request
}

export function isWorkbookCommandRequest(value: unknown): value is WorkbookCommandRequest {
  return checkWorkbookCommandRequest(value).status === 'valid'
}

function commandReceiptIssue(path: string, message: string): WorkbookCommandReceiptIssue {
  return Object.freeze({
    code: 'invalid_command_receipt',
    path,
    message,
  })
}

function pushCommandReceiptInputIssue(
  issues: WorkbookCommandReceiptIssue[],
  value: unknown,
  path: 'proof' | 'metadata',
  label: 'proof' | 'metadata',
): void {
  if (value === undefined) {
    return
  }
  try {
    normalizeWorkbookActionInput(value)
  } catch (error) {
    issues.push(commandReceiptIssue(inputPath(path, error), `Workbook command receipt ${label} is invalid: ${errorMessage(error)}`))
  }
}

function pushRequiredCommandReceiptStringIssue(
  issues: WorkbookCommandReceiptIssue[],
  value: Record<string, unknown>,
  key: 'featureId' | 'commandId',
  label: string,
): void {
  const entry = ownValue(value, key)
  if (typeof entry !== 'string') {
    issues.push(commandReceiptIssue(key, `Workbook command receipt ${label} must be a string`))
    return
  }
  const normalized = entry.trim()
  if (normalized === '') {
    issues.push(commandReceiptIssue(key, `Workbook command receipt ${label} cannot be empty`))
    return
  }
  if (normalized !== entry) {
    issues.push(commandReceiptIssue(key, `Workbook command receipt ${label} must not have leading or trailing whitespace`))
  }
}

function pushOptionalCommandReceiptStringIssue(issues: WorkbookCommandReceiptIssue[], value: unknown, path: string, label: string): void {
  if (value === undefined) {
    return
  }
  if (typeof value !== 'string') {
    issues.push(commandReceiptIssue(path, `Workbook command receipt ${label} must be a string`))
    return
  }
  const normalized = value.trim()
  if (normalized === '') {
    issues.push(commandReceiptIssue(path, `Workbook command receipt ${label} cannot be empty`))
    return
  }
  if (normalized !== value) {
    issues.push(commandReceiptIssue(path, `Workbook command receipt ${label} must not have leading or trailing whitespace`))
  }
}

function pushCommandReceiptOpsIssues(issues: WorkbookCommandReceiptIssue[], value: unknown, path: string, label: string): void {
  if (value === undefined) {
    return
  }
  if (!Array.isArray(value)) {
    issues.push(commandReceiptIssue(path, `Workbook command receipt ${label} ops must be an array`))
    return
  }
  const accessorPath = firstAccessorPath(value, path)
  if (accessorPath !== null) {
    issues.push(commandReceiptIssue(accessorPath, `Workbook command receipt ${label} ops must contain only data properties`))
    return
  }
  value.forEach((op, index) => {
    if (!isWorkbookOp(op)) {
      issues.push(commandReceiptIssue(`${path}[${index}]`, `Workbook command receipt ${label} op is invalid`))
    }
  })
}

function pushCommandReceiptChangedRangesIssues(issues: WorkbookCommandReceiptIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  if (!Array.isArray(value)) {
    issues.push(commandReceiptIssue('changedRanges', 'Workbook command receipt changed ranges must be an array'))
    return
  }
  const accessorPath = firstAccessorPath(value, 'changedRanges')
  if (accessorPath !== null) {
    issues.push(commandReceiptIssue(accessorPath, 'Workbook command receipt changed ranges must contain only data properties'))
    return
  }
  value.forEach((range, index) => {
    if (!isCellRangeRefData(range)) {
      issues.push(commandReceiptIssue(`changedRanges[${index}]`, 'Workbook command receipt changed range is invalid'))
    }
  })
}

function pushCommandReceiptUndoIssues(issues: WorkbookCommandReceiptIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  if (!isRecord(value)) {
    issues.push(commandReceiptIssue('undo', 'Workbook command receipt undo must be an object'))
    return
  }
  const id = ownValue(value, 'id')
  pushOptionalCommandReceiptStringIssue(issues, id, 'undo.id', 'undo id')
  if (id === undefined) {
    issues.push(commandReceiptIssue('undo.id', 'Workbook command receipt undo id must be a string'))
  }
  pushCommandReceiptOpsIssues(issues, ownValue(value, 'ops'), 'undo.ops', 'undo')
}

function pushCommandReceiptErrorsIssues(issues: WorkbookCommandReceiptIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  if (!Array.isArray(value)) {
    issues.push(commandReceiptIssue('errors', 'Workbook command receipt errors must be an array'))
    return
  }
  const accessorPath = firstAccessorPath(value, 'errors')
  if (accessorPath !== null) {
    issues.push(commandReceiptIssue(accessorPath, 'Workbook command receipt errors must contain only data properties'))
    return
  }
  value.forEach((error, index) => {
    pushOptionalCommandReceiptStringIssue(issues, error, `errors[${index}]`, 'error')
  })
}

function normalizeReceiptUndo(value: unknown): WorkbookUndoRef | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const id = ownValue(value, 'id')
  if (typeof id !== 'string') {
    return undefined
  }
  const rawOps = ownValue(value, 'ops')
  const ops = Array.isArray(rawOps) ? rawOps.map((op) => normalizeReceiptOp(id, op, 'undo')) : undefined
  return Object.freeze({
    id: normalizeRequiredString(id, 'Workbook command receipt undo id'),
    ...(ops !== undefined ? { ops: Object.freeze(ops) } : {}),
  })
}

function normalizedCommandReceipt(value: unknown): WorkbookCommandReceipt | null {
  if (!isRecord(value)) {
    return null
  }
  const status = ownValue(value, 'status')
  const featureId = ownValue(value, 'featureId')
  const commandId = ownValue(value, 'commandId')
  const category = ownValue(value, 'category')
  if (
    typeof featureId !== 'string' ||
    typeof commandId !== 'string' ||
    !isWorkbookCommandReceiptStatus(status) ||
    !isWorkbookCommandCategory(category)
  ) {
    return null
  }
  const previewOps = ownValue(value, 'previewOps')
  const appliedOps = ownValue(value, 'appliedOps')
  const undo = ownValue(value, 'undo')
  const changedRanges = ownValue(value, 'changedRanges')
  const proof = ownValue(value, 'proof')
  const message = ownValue(value, 'message')
  const metadata = ownValue(value, 'metadata')
  const errors = ownValue(value, 'errors')
  if (
    (previewOps !== undefined && !isEngineOpArray(previewOps)) ||
    (appliedOps !== undefined && !isEngineOpArray(appliedOps)) ||
    (changedRanges !== undefined && (!Array.isArray(changedRanges) || !changedRanges.every((range) => isCellRangeRefData(range)))) ||
    (proof !== undefined && !isWorkbookActionInput(proof)) ||
    (metadata !== undefined && !isWorkbookActionInput(metadata)) ||
    (message !== undefined && typeof message !== 'string') ||
    (errors !== undefined && (!Array.isArray(errors) || !errors.every((error) => typeof error === 'string')))
  ) {
    return null
  }
  const normalizedUndo = normalizeReceiptUndo(undo)
  if (undo !== undefined && normalizedUndo === undefined) {
    return null
  }
  return Object.freeze({
    status,
    featureId: normalizeWorkbookFeatureId(featureId, 'Workbook command receipt feature id'),
    commandId: normalizeRequiredString(commandId, 'Workbook command receipt command id'),
    category,
    ...(previewOps !== undefined
      ? { previewOps: Object.freeze(previewOps.map((op) => normalizeReceiptOp(commandId, op, 'preview'))) }
      : {}),
    ...(appliedOps !== undefined
      ? { appliedOps: Object.freeze(appliedOps.map((op) => normalizeReceiptOp(commandId, op, 'applied'))) }
      : {}),
    ...(normalizedUndo !== undefined ? { undo: normalizedUndo } : {}),
    ...(changedRanges !== undefined
      ? { changedRanges: Object.freeze(changedRanges.map((range) => normalizeReceiptRange(commandId, range))) }
      : {}),
    ...(proof !== undefined ? { proof: normalizeWorkbookActionInput(proof) } : {}),
    ...(message !== undefined ? { message: normalizeRequiredString(message, `Workbook command receipt ${commandId} message`) } : {}),
    ...(metadata !== undefined ? { metadata: normalizeWorkbookActionInput(metadata) } : {}),
    ...(errors !== undefined
      ? { errors: Object.freeze(errors.map((error) => normalizeRequiredString(error, `Workbook command receipt ${commandId} error`))) }
      : {}),
  })
}

export function checkWorkbookCommandReceipt(value: unknown): WorkbookCommandReceiptCheckResult {
  if (!isRecord(value)) {
    return {
      status: 'invalid',
      issues: Object.freeze([commandReceiptIssue('receipt', 'Workbook command receipt must be an object')]),
    }
  }

  const issues: WorkbookCommandReceiptIssue[] = []
  if (!isWorkbookCommandReceiptStatus(ownValue(value, 'status'))) {
    issues.push(commandReceiptIssue('status', 'Workbook command receipt status is invalid'))
  }
  pushRequiredCommandReceiptStringIssue(issues, value, 'featureId', 'feature id')
  pushRequiredCommandReceiptStringIssue(issues, value, 'commandId', 'command id')
  if (!isWorkbookCommandCategory(ownValue(value, 'category'))) {
    issues.push(commandReceiptIssue('category', 'Workbook command receipt category is invalid'))
  }
  pushCommandReceiptOpsIssues(issues, ownValue(value, 'previewOps'), 'previewOps', 'preview')
  pushCommandReceiptOpsIssues(issues, ownValue(value, 'appliedOps'), 'appliedOps', 'applied')
  pushCommandReceiptUndoIssues(issues, ownValue(value, 'undo'))
  pushCommandReceiptChangedRangesIssues(issues, ownValue(value, 'changedRanges'))
  pushCommandReceiptInputIssue(issues, ownValue(value, 'proof'), 'proof', 'proof')
  pushOptionalCommandReceiptStringIssue(issues, ownValue(value, 'message'), 'message', 'message')
  pushCommandReceiptInputIssue(issues, ownValue(value, 'metadata'), 'metadata', 'metadata')
  pushCommandReceiptErrorsIssues(issues, ownValue(value, 'errors'))

  if (issues.length > 0) {
    return {
      status: 'invalid',
      issues: Object.freeze(issues),
    }
  }
  const receipt = normalizedCommandReceipt(value)
  if (receipt === null) {
    return {
      status: 'invalid',
      issues: Object.freeze([commandReceiptIssue('receipt', 'Workbook command receipt is invalid')]),
    }
  }
  return {
    status: 'valid',
    receipt,
    issues: Object.freeze([]),
  }
}

export function normalizeWorkbookCommandReceipt(receipt: unknown): WorkbookCommandReceipt {
  const check = checkWorkbookCommandReceipt(receipt)
  if (check.status === 'invalid') {
    const [firstIssue] = check.issues
    throw new Error(
      firstIssue === undefined ? 'Workbook command receipt is invalid' : `Workbook command receipt is invalid: ${firstIssue.message}`,
    )
  }
  return check.receipt
}

export function isWorkbookCommandReceipt(value: unknown): value is WorkbookCommandReceipt {
  return checkWorkbookCommandReceipt(value).status === 'valid'
}

export function workbookCommandReceiptOpsMatch(receipt: Pick<WorkbookCommandReceipt, 'previewOps' | 'appliedOps'>): boolean | null {
  if (receipt.previewOps === undefined || receipt.appliedOps === undefined) {
    return null
  }
  if (!isEngineOpArray(receipt.previewOps) || !isEngineOpArray(receipt.appliedOps)) {
    return false
  }
  try {
    return canonicalJson(receipt.previewOps) === canonicalJson(receipt.appliedOps)
  } catch {
    return false
  }
}

function normalizeReceiptOp(commandId: string, op: EngineOp, label: string): EngineOp {
  if (!isWorkbookOp(op)) {
    throw new Error(`Workbook command receipt ${commandId} ${label} op is invalid`)
  }
  return deepFreezeOpClone(op)
}

function normalizeReceiptRange(commandId: string, range: CellRangeRef): CellRangeRef {
  if (!isCellRangeRefData(range)) {
    throw new Error(`Workbook command receipt ${commandId} changed range is invalid`)
  }
  return deepFreezeRangeClone(commandId, range)
}

function isEngineOpArray(value: unknown): value is readonly EngineOp[] {
  return Array.isArray(value) && firstAccessorPath(value, 'ops') === null && value.every((op) => isWorkbookOp(op))
}

function isCellRangeRefData(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof ownValue(value, 'sheetName') === 'string' &&
    typeof ownValue(value, 'startAddress') === 'string' &&
    typeof ownValue(value, 'endAddress') === 'string'
  )
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    const entries = Object.getOwnPropertyDescriptors(value)
    return Array.from({ length: value.length }, (_, index) => {
      const descriptor = entries[String(index)]
      if (descriptor === undefined) {
        return undefined
      }
      if (!('value' in descriptor)) {
        throw new Error('Accessor values cannot be canonicalized')
      }
      return canonicalValue(descriptor.value)
    })
  }
  if (isRecord(value)) {
    return Object.fromEntries(
      Object.entries(Object.getOwnPropertyDescriptors(value))
        .filter(([, descriptor]) => descriptor.enumerable)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, descriptor]) => {
          if (!('value' in descriptor)) {
            throw new Error('Accessor values cannot be canonicalized')
          }
          return [key, canonicalValue(descriptor.value)]
        }),
    )
  }
  return value
}

function deepFreezeOpClone(value: EngineOp): EngineOp {
  const cloned = cloneData(value)
  if (!isWorkbookOp(cloned)) {
    throw new Error('Workbook command receipt op clone is invalid')
  }
  return deepFreeze(cloned, new WeakSet())
}

function deepFreezeRangeClone(commandId: string, value: CellRangeRef): CellRangeRef {
  const cloned = cloneData(value)
  if (!isCellRangeRefData(cloned)) {
    throw new Error(`Workbook command receipt ${commandId} changed range clone is invalid`)
  }
  return deepFreeze(cloned, new WeakSet())
}

function deepFreeze<T>(value: T, seen: WeakSet<object>): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      deepFreeze(descriptor.value, seen)
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
    const descriptors = Object.getOwnPropertyDescriptors(value)
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined) {
        continue
      }
      if (!('value' in descriptor)) {
        throw new Error('Transport data must not contain accessors')
      }
      cloned[index] = cloneData(descriptor.value, seen)
    }
    return cloned
  }
  const cloned: Record<string, unknown> = Object.create(Object.getPrototypeOf(value))
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (!descriptor.enumerable) {
      return
    }
    if (!('value' in descriptor)) {
      throw new Error('Transport data must not contain accessors')
    }
    Object.defineProperty(cloned, key, {
      configurable: true,
      enumerable: true,
      value: cloneData(descriptor.value, seen),
      writable: true,
    })
  })
  return cloned
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function normalizeRequiredString(value: string, label: string): string {
  const normalized = value.trim()
  if (normalized === '') {
    throw new Error(`${label} cannot be empty`)
  }
  if (normalized !== value) {
    throw new Error(`${label} must not have leading or trailing whitespace`)
  }
  return normalized
}
