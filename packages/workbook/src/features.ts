import type { CellRangeRef } from '@bilig/protocol'
import { isCellRangeRef } from '@bilig/protocol'
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
  const entry = value[key]
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
  if (!isRecord(value) || typeof value['featureId'] !== 'string' || typeof value['commandId'] !== 'string') {
    return null
  }
  const category = value['category']
  const mode = value['mode']
  const input = value['input']
  if (
    (category !== undefined && !isWorkbookCommandCategory(category)) ||
    (mode !== undefined && !isWorkbookCommandExecutionMode(mode)) ||
    (input !== undefined && !isWorkbookActionInput(input))
  ) {
    return null
  }
  return Object.freeze({
    featureId: normalizeWorkbookFeatureId(value['featureId'], 'Workbook command request feature id'),
    commandId: normalizeRequiredString(value['commandId'], 'Workbook command request command id'),
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
  if (value['category'] !== undefined && !isWorkbookCommandCategory(value['category'])) {
    issues.push(commandRequestIssue('category', 'Workbook command request category is invalid'))
  }
  if (value['mode'] !== undefined && !isWorkbookCommandExecutionMode(value['mode'])) {
    issues.push(commandRequestIssue('mode', 'Workbook command request mode is invalid'))
  }
  pushCommandRequestInputIssue(issues, value['input'])

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
  const entry = value[key]
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
  value.forEach((range, index) => {
    if (!isCellRangeRef(range)) {
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
  pushOptionalCommandReceiptStringIssue(issues, value['id'], 'undo.id', 'undo id')
  if (value['id'] === undefined) {
    issues.push(commandReceiptIssue('undo.id', 'Workbook command receipt undo id must be a string'))
  }
  pushCommandReceiptOpsIssues(issues, value['ops'], 'undo.ops', 'undo')
}

function pushCommandReceiptErrorsIssues(issues: WorkbookCommandReceiptIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  if (!Array.isArray(value)) {
    issues.push(commandReceiptIssue('errors', 'Workbook command receipt errors must be an array'))
    return
  }
  value.forEach((error, index) => {
    pushOptionalCommandReceiptStringIssue(issues, error, `errors[${index}]`, 'error')
  })
}

function normalizeReceiptUndo(value: unknown): WorkbookUndoRef | undefined {
  if (!isRecord(value) || typeof value['id'] !== 'string') {
    return undefined
  }
  const id = value['id']
  const ops = Array.isArray(value['ops']) ? value['ops'].map((op) => normalizeReceiptOp(id, op, 'undo')) : undefined
  return Object.freeze({
    id: normalizeRequiredString(id, 'Workbook command receipt undo id'),
    ...(ops !== undefined ? { ops: Object.freeze(ops) } : {}),
  })
}

function normalizedCommandReceipt(value: unknown): WorkbookCommandReceipt | null {
  if (
    !isRecord(value) ||
    typeof value['featureId'] !== 'string' ||
    typeof value['commandId'] !== 'string' ||
    !isWorkbookCommandReceiptStatus(value['status']) ||
    !isWorkbookCommandCategory(value['category'])
  ) {
    return null
  }
  const previewOps = value['previewOps']
  const appliedOps = value['appliedOps']
  const undo = value['undo']
  const changedRanges = value['changedRanges']
  const proof = value['proof']
  const message = value['message']
  const metadata = value['metadata']
  const errors = value['errors']
  if (
    (previewOps !== undefined && (!Array.isArray(previewOps) || !previewOps.every((op) => isWorkbookOp(op)))) ||
    (appliedOps !== undefined && (!Array.isArray(appliedOps) || !appliedOps.every((op) => isWorkbookOp(op)))) ||
    (changedRanges !== undefined && (!Array.isArray(changedRanges) || !changedRanges.every((range) => isCellRangeRef(range)))) ||
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
  const commandId = value['commandId']
  return Object.freeze({
    status: value['status'],
    featureId: normalizeWorkbookFeatureId(value['featureId'], 'Workbook command receipt feature id'),
    commandId: normalizeRequiredString(commandId, 'Workbook command receipt command id'),
    category: value['category'],
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
  if (!isWorkbookCommandReceiptStatus(value['status'])) {
    issues.push(commandReceiptIssue('status', 'Workbook command receipt status is invalid'))
  }
  pushRequiredCommandReceiptStringIssue(issues, value, 'featureId', 'feature id')
  pushRequiredCommandReceiptStringIssue(issues, value, 'commandId', 'command id')
  if (!isWorkbookCommandCategory(value['category'])) {
    issues.push(commandReceiptIssue('category', 'Workbook command receipt category is invalid'))
  }
  pushCommandReceiptOpsIssues(issues, value['previewOps'], 'previewOps', 'preview')
  pushCommandReceiptOpsIssues(issues, value['appliedOps'], 'appliedOps', 'applied')
  pushCommandReceiptUndoIssues(issues, value['undo'])
  pushCommandReceiptChangedRangesIssues(issues, value['changedRanges'])
  pushCommandReceiptInputIssue(issues, value['proof'], 'proof', 'proof')
  pushOptionalCommandReceiptStringIssue(issues, value['message'], 'message', 'message')
  pushCommandReceiptInputIssue(issues, value['metadata'], 'metadata', 'metadata')
  pushCommandReceiptErrorsIssues(issues, value['errors'])

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
  return JSON.stringify(receipt.previewOps) === JSON.stringify(receipt.appliedOps)
}

function normalizeReceiptOp(commandId: string, op: EngineOp, label: string): EngineOp {
  if (!isWorkbookOp(op)) {
    throw new Error(`Workbook command receipt ${commandId} ${label} op is invalid`)
  }
  return structuredClone(op)
}

function normalizeReceiptRange(commandId: string, range: CellRangeRef): CellRangeRef {
  if (!isCellRangeRef(range)) {
    throw new Error(`Workbook command receipt ${commandId} changed range is invalid`)
  }
  return structuredClone(range)
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
