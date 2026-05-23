import type { CellRangeRef, CellStylePatch, CellStyleRecord, LiteralInput } from '@bilig/protocol'
import {
  isWorkbookCommandCategory,
  isWorkbookProjectionInterceptorPoint,
  isWorkbookUiContributionSlot,
  normalizeWorkbookFeatureId,
  type WorkbookCommandCategory,
  type WorkbookFeatureId,
  type WorkbookProjectionInterceptorPoint,
  type WorkbookUiContributionSlot,
} from './features.js'
import {
  WorkbookActionInputError,
  isWorkbookActionInput,
  isWorkbookActionInputDescription,
  normalizeWorkbookActionInput,
  normalizeWorkbookActionInputDescription,
  type WorkbookActionInput,
  type WorkbookActionInputDescription,
} from './input.js'

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

export type WorkbookFeaturePluginIssueCode = 'invalid_feature_plugin'

export interface WorkbookFeaturePluginIssue {
  readonly code: WorkbookFeaturePluginIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookFeaturePluginCheckResult =
  | {
      readonly status: 'valid'
      readonly plugin: WorkbookFeaturePlugin
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookFeaturePluginIssue[]
    }

export function checkWorkbookFeaturePlugin(value: unknown): WorkbookFeaturePluginCheckResult {
  if (!isRecord(value)) {
    return {
      status: 'invalid',
      issues: Object.freeze([featurePluginIssue('plugin', 'Workbook feature plugin must be an object')]),
    }
  }

  const issues: WorkbookFeaturePluginIssue[] = []
  const pluginId = pushRequiredFeaturePluginStringIssue(issues, value['id'], 'id', 'Workbook feature id')
  pushRequiredFeaturePluginStringIssue(issues, value['version'], 'version', 'Workbook feature version')
  pushFeaturePluginDependencyIssues(issues, value['dependsOn'])
  readFeaturePluginArray(issues, value['commands'], 'commands', 'Workbook feature commands').forEach((command, index) => {
    pushFeaturePluginCommandIssues(issues, command, index, pluginId)
  })
  readFeaturePluginArray(
    issues,
    value['projectionInterceptors'],
    'projectionInterceptors',
    'Workbook feature projection interceptors',
  ).forEach((interceptor, index) => {
    pushFeaturePluginProjectionIssues(issues, interceptor, index, pluginId)
  })
  readFeaturePluginArray(issues, value['uiContributions'], 'uiContributions', 'Workbook feature UI contributions').forEach(
    (contribution, index) => {
      pushFeaturePluginUiContributionIssues(issues, contribution, index, pluginId)
    },
  )
  pushFeaturePluginLifecycleIssue(issues, value['register'], 'register')
  pushFeaturePluginLifecycleIssue(issues, value['activate'], 'activate')
  pushFeaturePluginLifecycleIssue(issues, value['dispose'], 'dispose')

  if (issues.length > 0) {
    return {
      status: 'invalid',
      issues: Object.freeze(issues),
    }
  }
  if (!isWorkbookFeaturePluginRecord(value)) {
    return {
      status: 'invalid',
      issues: Object.freeze([featurePluginIssue('plugin', 'Workbook feature plugin is invalid')]),
    }
  }
  return {
    status: 'valid',
    plugin: normalizedFeaturePlugin(value),
    issues: Object.freeze([]),
  }
}

export function defineWorkbookFeaturePlugin(plugin: WorkbookFeaturePlugin): WorkbookFeaturePlugin {
  const check = checkWorkbookFeaturePlugin(plugin)
  if (check.status === 'invalid') {
    const [firstIssue] = check.issues
    throw new Error(
      firstIssue === undefined ? 'Workbook feature plugin is invalid' : `Workbook feature plugin is invalid: ${firstIssue.message}`,
    )
  }
  return check.plugin
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
  const input = descriptor.input === undefined ? undefined : normalizeWorkbookActionInputDescription(descriptor.input)
  if (!isWorkbookCommandCategory(descriptor.category)) {
    throw new Error(`Workbook command ${id} category is invalid`)
  }
  return Object.freeze({
    id,
    featureId,
    category: descriptor.category,
    label,
    ...(description !== undefined ? { description } : {}),
    ...(input !== undefined ? { input } : {}),
  })
}

function featurePluginIssue(path: string, message: string): WorkbookFeaturePluginIssue {
  return Object.freeze({
    code: 'invalid_feature_plugin',
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

function pushFeaturePluginInputDescriptionIssue(issues: WorkbookFeaturePluginIssue[], value: unknown, path: string, label: string): void {
  if (value === undefined) {
    return
  }
  try {
    normalizeWorkbookActionInputDescription(value)
  } catch (error) {
    issues.push(featurePluginIssue(inputPath(path, error), `${label} input description is invalid: ${errorMessage(error)}`))
  }
}

function pushFeaturePluginInputIssue(issues: WorkbookFeaturePluginIssue[], value: unknown, path: string, label: string): void {
  if (value === undefined) {
    return
  }
  try {
    normalizeWorkbookActionInput(value)
  } catch (error) {
    issues.push(featurePluginIssue(inputPath(path, error), `${label} metadata is invalid: ${errorMessage(error)}`))
  }
}

function pushRequiredFeaturePluginStringIssue(
  issues: WorkbookFeaturePluginIssue[],
  value: unknown,
  path: string,
  label: string,
): string | undefined {
  if (typeof value !== 'string') {
    issues.push(featurePluginIssue(path, `${label} must be a string`))
    return undefined
  }
  const normalized = value.trim()
  if (normalized === '') {
    issues.push(featurePluginIssue(path, `${label} cannot be empty`))
    return undefined
  }
  if (normalized !== value) {
    issues.push(featurePluginIssue(path, `${label} must not have leading or trailing whitespace`))
    return undefined
  }
  return normalized
}

function pushOptionalFeaturePluginStringIssue(issues: WorkbookFeaturePluginIssue[], value: unknown, path: string, label: string): void {
  if (value === undefined) {
    return
  }
  pushRequiredFeaturePluginStringIssue(issues, value, path, label)
}

function readFeaturePluginArray(issues: WorkbookFeaturePluginIssue[], value: unknown, path: string, label: string): readonly unknown[] {
  if (!Array.isArray(value)) {
    issues.push(featurePluginIssue(path, `${label} must be an array`))
    return []
  }
  return value
}

function pushFeaturePluginDependencyIssues(issues: WorkbookFeaturePluginIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  if (!Array.isArray(value)) {
    issues.push(featurePluginIssue('dependsOn', 'Workbook feature dependencies must be an array'))
    return
  }
  value.forEach((dependency, index) => {
    pushRequiredFeaturePluginStringIssue(issues, dependency, `dependsOn[${index}]`, 'Workbook feature dependency')
  })
}

function pushFeaturePluginCommandIssues(
  issues: WorkbookFeaturePluginIssue[],
  command: unknown,
  index: number,
  pluginId: string | undefined,
): void {
  const path = `commands[${index}]`
  if (!isRecord(command)) {
    issues.push(featurePluginIssue(path, 'Workbook command must be an object'))
    return
  }

  const commandId = pushRequiredFeaturePluginStringIssue(issues, command['id'], `${path}.id`, 'Workbook command id')
  const featureId = pushRequiredFeaturePluginStringIssue(issues, command['featureId'], `${path}.featureId`, 'Workbook command feature id')
  if (pluginId !== undefined && featureId !== undefined && featureId !== pluginId) {
    issues.push(
      featurePluginIssue(
        `${path}.featureId`,
        `Workbook command ${commandId ?? path} feature id ${featureId} does not match plugin ${pluginId}`,
      ),
    )
  }
  if (!isWorkbookCommandCategory(command['category'])) {
    issues.push(featurePluginIssue(`${path}.category`, `Workbook command ${commandId ?? path} category is invalid`))
  }
  pushRequiredFeaturePluginStringIssue(issues, command['label'], `${path}.label`, `Workbook command ${commandId ?? path} label`)
  pushOptionalFeaturePluginStringIssue(
    issues,
    command['description'],
    `${path}.description`,
    `Workbook command ${commandId ?? path} description`,
  )
  pushFeaturePluginInputDescriptionIssue(issues, command['input'], `${path}.input`, `Workbook command ${commandId ?? path}`)
}

function pushFeaturePluginProjectionIssues(
  issues: WorkbookFeaturePluginIssue[],
  interceptor: unknown,
  index: number,
  pluginId: string | undefined,
): void {
  const path = `projectionInterceptors[${index}]`
  if (!isRecord(interceptor)) {
    issues.push(featurePluginIssue(path, 'Workbook projection interceptor must be an object'))
    return
  }

  const id = pushRequiredFeaturePluginStringIssue(issues, interceptor['id'], `${path}.id`, 'Workbook projection interceptor id')
  const featureId = pushRequiredFeaturePluginStringIssue(
    issues,
    interceptor['featureId'],
    `${path}.featureId`,
    'Workbook projection interceptor feature id',
  )
  if (pluginId !== undefined && featureId !== undefined && featureId !== pluginId) {
    issues.push(
      featurePluginIssue(
        `${path}.featureId`,
        `Workbook projection interceptor ${id ?? path} feature id ${featureId} does not match plugin ${pluginId}`,
      ),
    )
  }
  if (!isWorkbookProjectionInterceptorPoint(interceptor['point'])) {
    issues.push(featurePluginIssue(`${path}.point`, `Workbook projection interceptor ${id ?? path} point is invalid`))
  }
  if (interceptor['priority'] !== undefined && !isSafeInteger(interceptor['priority'])) {
    issues.push(featurePluginIssue(`${path}.priority`, `Workbook projection interceptor ${id ?? path} priority is invalid`))
  }
  pushOptionalFeaturePluginStringIssue(issues, interceptor['label'], `${path}.label`, `Workbook projection interceptor ${id ?? path} label`)
}

function pushFeaturePluginUiContributionIssues(
  issues: WorkbookFeaturePluginIssue[],
  contribution: unknown,
  index: number,
  pluginId: string | undefined,
): void {
  const path = `uiContributions[${index}]`
  if (!isRecord(contribution)) {
    issues.push(featurePluginIssue(path, 'Workbook UI contribution must be an object'))
    return
  }

  const id = pushRequiredFeaturePluginStringIssue(issues, contribution['id'], `${path}.id`, 'Workbook UI contribution id')
  const featureId = pushRequiredFeaturePluginStringIssue(
    issues,
    contribution['featureId'],
    `${path}.featureId`,
    'Workbook UI contribution feature id',
  )
  if (pluginId !== undefined && featureId !== undefined && featureId !== pluginId) {
    issues.push(
      featurePluginIssue(
        `${path}.featureId`,
        `Workbook UI contribution ${id ?? path} feature id ${featureId} does not match plugin ${pluginId}`,
      ),
    )
  }
  if (!isWorkbookUiContributionSlot(contribution['slot'])) {
    issues.push(featurePluginIssue(`${path}.slot`, `Workbook UI contribution ${id ?? path} slot is invalid`))
  }
  pushRequiredFeaturePluginStringIssue(issues, contribution['label'], `${path}.label`, `Workbook UI contribution ${id ?? path} label`)
  if (contribution['order'] !== undefined && !isSafeInteger(contribution['order'])) {
    issues.push(featurePluginIssue(`${path}.order`, `Workbook UI contribution ${id ?? path} order is invalid`))
  }
  pushFeaturePluginInputIssue(issues, contribution['metadata'], `${path}.metadata`, `Workbook UI contribution ${id ?? path}`)
}

function pushFeaturePluginLifecycleIssue(
  issues: WorkbookFeaturePluginIssue[],
  value: unknown,
  path: 'register' | 'activate' | 'dispose',
): void {
  if (value !== undefined && typeof value !== 'function') {
    issues.push(featurePluginIssue(path, `Workbook feature ${path} must be a function`))
  }
}

function normalizedFeaturePlugin(plugin: WorkbookFeaturePlugin): WorkbookFeaturePlugin {
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
  if (interceptor.priority !== undefined && !isSafeInteger(interceptor.priority)) {
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
  if (contribution.order !== undefined && !isSafeInteger(contribution.order)) {
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

function isWorkbookFeaturePluginRecord(value: unknown): value is WorkbookFeaturePlugin {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['version'] === 'string' &&
    (value['dependsOn'] === undefined || (Array.isArray(value['dependsOn']) && value['dependsOn'].every(isString))) &&
    Array.isArray(value['commands']) &&
    value['commands'].every(isWorkbookCommandDescriptorRecord) &&
    Array.isArray(value['projectionInterceptors']) &&
    value['projectionInterceptors'].every(isWorkbookProjectionInterceptorRecord) &&
    Array.isArray(value['uiContributions']) &&
    value['uiContributions'].every(isWorkbookUiContributionRecord) &&
    isOptionalLifecycleHook(value['register']) &&
    isOptionalLifecycleHook(value['activate']) &&
    isOptionalLifecycleHook(value['dispose'])
  )
}

function isWorkbookCommandDescriptorRecord(value: unknown): value is WorkbookCommandDescriptor {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['featureId'] === 'string' &&
    isWorkbookCommandCategory(value['category']) &&
    typeof value['label'] === 'string' &&
    (value['description'] === undefined || typeof value['description'] === 'string') &&
    (value['input'] === undefined || isWorkbookActionInputDescription(value['input']))
  )
}

function isWorkbookProjectionInterceptorRecord(value: unknown): value is WorkbookProjectionInterceptorRegistration {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['featureId'] === 'string' &&
    isWorkbookProjectionInterceptorPoint(value['point']) &&
    (value['priority'] === undefined || isSafeInteger(value['priority'])) &&
    (value['label'] === undefined || typeof value['label'] === 'string')
  )
}

function isWorkbookUiContributionRecord(value: unknown): value is WorkbookUiContribution {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['featureId'] === 'string' &&
    isWorkbookUiContributionSlot(value['slot']) &&
    typeof value['label'] === 'string' &&
    (value['order'] === undefined || isSafeInteger(value['order'])) &&
    (value['metadata'] === undefined || isWorkbookActionInput(value['metadata']))
  )
}

function isString(value: unknown): value is string {
  return typeof value === 'string'
}

function isOptionalLifecycleHook(value: unknown): value is undefined | ((context: WorkbookFeatureLifecycleContext) => void) {
  return value === undefined || typeof value === 'function'
}

function isSafeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && Number.isFinite(value)
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
