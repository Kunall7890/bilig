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

function isWorkbookCommandCategory(value: unknown): value is WorkbookCommandCategory {
  return value === 'command' || value === 'operation' || value === 'mutation'
}

function isWorkbookCommandReceiptStatus(value: unknown): value is WorkbookCommandReceiptStatus {
  return value === 'previewed' || value === 'applied' || value === 'rejected' || value === 'noop'
}

function isWorkbookProjectionInterceptorPoint(value: unknown): value is WorkbookProjectionInterceptorPoint {
  return (
    value === 'cellDisplay' ||
    value === 'cellStyle' ||
    value === 'rangeChrome' ||
    value === 'rowVisibility' ||
    value === 'beforeCommand' ||
    value === 'commandMetadata'
  )
}

function isWorkbookUiContributionSlot(value: unknown): value is WorkbookUiContributionSlot {
  return value === 'toolbar' || value === 'sidePanel' || value === 'floatingOverlay' || value === 'status'
}
