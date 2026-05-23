import type { CellRangeRef, CellStylePatch, CellStyleRecord, LiteralInput } from '@bilig/protocol'
import { isCellRangeRef } from '@bilig/protocol'
import { isWorkbookOp } from './guards.js'
import {
  isWorkbookActionInput,
  isWorkbookActionInputDescription,
  normalizeWorkbookActionInput,
  type WorkbookActionInput,
  type WorkbookActionInputDescription,
} from './input.js'
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

export interface WorkbookFeatureLifecycleContext {
  readonly featureId: WorkbookFeatureId
  readonly activeFeatures: readonly WorkbookFeatureId[]
}

export interface WorkbookCommandDescriptor {
  readonly id: string
  readonly featureId: WorkbookFeatureId
  readonly category: WorkbookCommandCategory
  readonly label: string
  readonly description?: string
  readonly input?: WorkbookActionInputDescription
}

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

export interface WorkbookCellDisplayProjection {
  readonly value?: LiteralInput
  readonly text?: string
  readonly metadata?: WorkbookActionInput
}

export interface WorkbookCellStyleProjection {
  readonly style?: CellStylePatch | CellStyleRecord
  readonly metadata?: WorkbookActionInput
}

export interface WorkbookRangeChromeProjection {
  readonly id: string
  readonly featureId: WorkbookFeatureId
  readonly source: 'workbook-metadata' | 'command-preview' | 'runtime'
  readonly range: CellRangeRef
  readonly role: string
  readonly label?: string
  readonly metadata?: WorkbookActionInput
}

export interface WorkbookRowVisibilityProjection {
  readonly hidden?: boolean
  readonly metadata?: WorkbookActionInput
}

export interface WorkbookCommandMetadataProjection {
  readonly label?: string
  readonly changedRanges?: readonly CellRangeRef[]
  readonly semanticTargets?: readonly WorkbookActionInput[]
  readonly metadata?: WorkbookActionInput
}

export interface WorkbookProjectionContext {
  readonly featureId: WorkbookFeatureId
}

export interface WorkbookProjectionInterceptorRegistration {
  readonly id: string
  readonly featureId: WorkbookFeatureId
  readonly point: WorkbookProjectionInterceptorPoint
  readonly priority?: number
  readonly label?: string
}

export interface WorkbookUiContribution {
  readonly id: string
  readonly featureId: WorkbookFeatureId
  readonly slot: WorkbookUiContributionSlot
  readonly label: string
  readonly order?: number
  readonly metadata?: WorkbookActionInput
}

export interface WorkbookFeatureRegistration {
  readonly commands: readonly WorkbookCommandDescriptor[]
  readonly projectionInterceptors: readonly WorkbookProjectionInterceptorRegistration[]
  readonly uiContributions: readonly WorkbookUiContribution[]
}

export interface WorkbookFeaturePlugin extends WorkbookFeatureRegistration {
  readonly id: WorkbookFeatureId
  readonly version: string
  readonly dependsOn?: readonly WorkbookFeatureId[]
  readonly register?: (context: WorkbookFeatureLifecycleContext) => void
  readonly activate?: (context: WorkbookFeatureLifecycleContext) => void
  readonly dispose?: (context: WorkbookFeatureLifecycleContext) => void
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
  if (value['input'] !== undefined && !isWorkbookActionInput(value['input'])) {
    issues.push(commandRequestIssue('input', 'Workbook command request input must be JSON-safe'))
  }

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

export function defineWorkbookFeaturePlugin(plugin: WorkbookFeaturePlugin): WorkbookFeaturePlugin {
  const id = normalizeWorkbookFeatureId(plugin.id)
  const version = normalizeRequiredString(plugin.version, `Workbook feature ${id} version`)
  const dependsOn = plugin.dependsOn?.map((dependency) => normalizeWorkbookFeatureId(dependency, `Workbook feature ${id} dependency`))
  const commands = plugin.commands.map((command) => normalizeWorkbookCommandDescriptor(command, id))
  const projectionInterceptors = plugin.projectionInterceptors.map((interceptor) => normalizeProjectionInterceptor(interceptor, id))
  const uiContributions = plugin.uiContributions.map((contribution) => normalizeUiContribution(contribution, id))
  return Object.freeze({
    id,
    version,
    ...(dependsOn !== undefined ? { dependsOn: Object.freeze([...dependsOn]) } : {}),
    commands: Object.freeze(commands),
    projectionInterceptors: Object.freeze(projectionInterceptors),
    uiContributions: Object.freeze(uiContributions),
    ...(plugin.register !== undefined ? { register: plugin.register } : {}),
    ...(plugin.activate !== undefined ? { activate: plugin.activate } : {}),
    ...(plugin.dispose !== undefined ? { dispose: plugin.dispose } : {}),
  })
}

export function normalizeWorkbookCommandDescriptor(
  descriptor: WorkbookCommandDescriptor,
  expectedFeatureId?: WorkbookFeatureId,
): WorkbookCommandDescriptor {
  const featureId = normalizeWorkbookFeatureId(descriptor.featureId, 'Workbook command feature id')
  if (expectedFeatureId !== undefined && featureId !== expectedFeatureId) {
    throw new Error(`Workbook command ${descriptor.id} feature id ${featureId} does not match plugin ${expectedFeatureId}`)
  }
  const id = normalizeRequiredString(descriptor.id, 'Workbook command id')
  const label = normalizeRequiredString(descriptor.label, `Workbook command ${id} label`)
  const description =
    descriptor.description === undefined ? undefined : normalizeRequiredString(descriptor.description, `Workbook command ${id} description`)
  if (!isWorkbookCommandCategory(descriptor.category)) {
    throw new Error(`Workbook command ${id} category is invalid`)
  }
  if (descriptor.input !== undefined && !isWorkbookActionInputDescription(descriptor.input)) {
    throw new Error(`Workbook command ${id} input description is invalid`)
  }
  return Object.freeze({
    id,
    featureId,
    category: descriptor.category,
    label,
    ...(description !== undefined ? { description } : {}),
    ...(descriptor.input !== undefined ? { input: descriptor.input } : {}),
  })
}

export function normalizeWorkbookCommandReceipt(receipt: WorkbookCommandReceipt): WorkbookCommandReceipt {
  const featureId = normalizeWorkbookFeatureId(receipt.featureId, 'Workbook command receipt feature id')
  const commandId = normalizeRequiredString(receipt.commandId, 'Workbook command receipt command id')
  if (!isWorkbookCommandCategory(receipt.category)) {
    throw new Error(`Workbook command receipt ${commandId} category is invalid`)
  }
  if (!isWorkbookCommandReceiptStatus(receipt.status)) {
    throw new Error(`Workbook command receipt ${commandId} status is invalid`)
  }
  const previewOps = receipt.previewOps?.map((op) => normalizeReceiptOp(commandId, op, 'preview'))
  const appliedOps = receipt.appliedOps?.map((op) => normalizeReceiptOp(commandId, op, 'applied'))
  const changedRanges = receipt.changedRanges?.map((range) => normalizeReceiptRange(commandId, range))
  const proof = receipt.proof === undefined ? undefined : normalizeWorkbookActionInput(receipt.proof)
  const metadata = receipt.metadata === undefined ? undefined : normalizeWorkbookActionInput(receipt.metadata)
  const message =
    receipt.message === undefined ? undefined : normalizeRequiredString(receipt.message, `Workbook command receipt ${commandId} message`)
  const errors = receipt.errors?.map((error) => normalizeRequiredString(error, `Workbook command receipt ${commandId} error`))
  return Object.freeze({
    status: receipt.status,
    featureId,
    commandId,
    category: receipt.category,
    ...(previewOps !== undefined ? { previewOps: Object.freeze(previewOps) } : {}),
    ...(appliedOps !== undefined ? { appliedOps: Object.freeze(appliedOps) } : {}),
    ...(receipt.undo !== undefined ? { undo: receipt.undo } : {}),
    ...(changedRanges !== undefined ? { changedRanges: Object.freeze(changedRanges) } : {}),
    ...(proof !== undefined ? { proof } : {}),
    ...(message !== undefined ? { message } : {}),
    ...(metadata !== undefined ? { metadata } : {}),
    ...(errors !== undefined ? { errors: Object.freeze(errors) } : {}),
  })
}

export function isWorkbookCommandReceipt(value: unknown): value is WorkbookCommandReceipt {
  if (!isWorkbookCommandReceiptCandidate(value)) {
    return false
  }
  try {
    normalizeWorkbookCommandReceipt(value)
    return true
  } catch {
    return false
  }
}

export function workbookCommandReceiptOpsMatch(receipt: Pick<WorkbookCommandReceipt, 'previewOps' | 'appliedOps'>): boolean | null {
  if (receipt.previewOps === undefined || receipt.appliedOps === undefined) {
    return null
  }
  return JSON.stringify(receipt.previewOps) === JSON.stringify(receipt.appliedOps)
}

function normalizeProjectionInterceptor(
  interceptor: WorkbookProjectionInterceptorRegistration,
  expectedFeatureId: WorkbookFeatureId,
): WorkbookProjectionInterceptorRegistration {
  const featureId = normalizeWorkbookFeatureId(interceptor.featureId, 'Workbook projection interceptor feature id')
  if (featureId !== expectedFeatureId) {
    throw new Error(`Workbook projection interceptor ${interceptor.id} feature id ${featureId} does not match plugin ${expectedFeatureId}`)
  }
  const id = normalizeRequiredString(interceptor.id, 'Workbook projection interceptor id')
  if (!isWorkbookProjectionInterceptorPoint(interceptor.point)) {
    throw new Error(`Workbook projection interceptor ${id} point is invalid`)
  }
  const label =
    interceptor.label === undefined ? undefined : normalizeRequiredString(interceptor.label, `Workbook projection interceptor ${id} label`)
  if (interceptor.priority !== undefined && (!Number.isSafeInteger(interceptor.priority) || !Number.isFinite(interceptor.priority))) {
    throw new Error(`Workbook projection interceptor ${id} priority is invalid`)
  }
  return Object.freeze({
    id,
    featureId,
    point: interceptor.point,
    ...(interceptor.priority !== undefined ? { priority: interceptor.priority } : {}),
    ...(label !== undefined ? { label } : {}),
  })
}

function normalizeUiContribution(contribution: WorkbookUiContribution, expectedFeatureId: WorkbookFeatureId): WorkbookUiContribution {
  const featureId = normalizeWorkbookFeatureId(contribution.featureId, 'Workbook UI contribution feature id')
  if (featureId !== expectedFeatureId) {
    throw new Error(`Workbook UI contribution ${contribution.id} feature id ${featureId} does not match plugin ${expectedFeatureId}`)
  }
  const id = normalizeRequiredString(contribution.id, 'Workbook UI contribution id')
  const label = normalizeRequiredString(contribution.label, `Workbook UI contribution ${id} label`)
  if (!isWorkbookUiContributionSlot(contribution.slot)) {
    throw new Error(`Workbook UI contribution ${id} slot is invalid`)
  }
  if (contribution.order !== undefined && (!Number.isSafeInteger(contribution.order) || !Number.isFinite(contribution.order))) {
    throw new Error(`Workbook UI contribution ${id} order is invalid`)
  }
  if (contribution.metadata !== undefined && !isWorkbookActionInput(contribution.metadata)) {
    throw new Error(`Workbook UI contribution ${id} metadata is invalid`)
  }
  return Object.freeze({
    id,
    featureId,
    slot: contribution.slot,
    label,
    ...(contribution.order !== undefined ? { order: contribution.order } : {}),
    ...(contribution.metadata !== undefined ? { metadata: normalizeWorkbookActionInput(contribution.metadata) } : {}),
  })
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

function isWorkbookCommandReceiptCandidate(value: unknown): value is WorkbookCommandReceipt {
  return (
    isRecord(value) &&
    typeof value['featureId'] === 'string' &&
    typeof value['commandId'] === 'string' &&
    isWorkbookCommandCategory(value['category']) &&
    isWorkbookCommandReceiptStatus(value['status']) &&
    (value['previewOps'] === undefined || (Array.isArray(value['previewOps']) && value['previewOps'].every((op) => isWorkbookOp(op)))) &&
    (value['appliedOps'] === undefined || (Array.isArray(value['appliedOps']) && value['appliedOps'].every((op) => isWorkbookOp(op)))) &&
    (value['changedRanges'] === undefined ||
      (Array.isArray(value['changedRanges']) && value['changedRanges'].every((range) => isCellRangeRef(range)))) &&
    (value['proof'] === undefined || isWorkbookActionInput(value['proof'])) &&
    (value['metadata'] === undefined || isWorkbookActionInput(value['metadata'])) &&
    (value['message'] === undefined || typeof value['message'] === 'string') &&
    (value['errors'] === undefined || (Array.isArray(value['errors']) && value['errors'].every((error) => typeof error === 'string')))
  )
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
