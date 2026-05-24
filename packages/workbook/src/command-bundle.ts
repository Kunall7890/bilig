import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import { workbookCommandResultFor, type WorkbookCommandResult } from './command-result.js'
import { isWorkbookOp } from './guards.js'
import type { EngineOp } from './ops.js'
import { checkWorkbookCommandRequest, normalizeWorkbookCommandRequest, type WorkbookCommandRequest } from './features.js'

export type WorkbookCommandBundleCommandKind = 'request' | 'op'

export const workbookCommandBundleCommandKinds = Object.freeze([
  'request',
  'op',
] as const satisfies readonly WorkbookCommandBundleCommandKind[])

const WORKBOOK_COMMAND_BUNDLE_COMMAND_KIND_SET = new Set<string>(workbookCommandBundleCommandKinds)

export interface WorkbookCommandBundleScope {
  readonly maxTouchedCells?: number
}

export interface WorkbookCommandBundleCommandBase {
  readonly id?: string
  readonly touchedRanges?: readonly CellRangeRef[]
  readonly destructive?: boolean
}

export interface WorkbookCommandBundleRequestCommand extends WorkbookCommandBundleCommandBase {
  readonly kind: 'request'
  readonly request: WorkbookCommandRequest
}

export interface WorkbookCommandBundleOpCommand extends WorkbookCommandBundleCommandBase {
  readonly kind: 'op'
  readonly op: EngineOp
}

export type WorkbookCommandBundleCommand = WorkbookCommandBundleRequestCommand | WorkbookCommandBundleOpCommand

export interface WorkbookCommandBundle {
  readonly id?: string
  readonly targetRevision: number
  readonly idempotencyKey: string
  readonly scope?: WorkbookCommandBundleScope
  readonly commands: readonly WorkbookCommandBundleCommand[]
}

export type WorkbookCommandBundleIssueCode =
  | 'invalid_bundle'
  | 'missing_target_revision'
  | 'invalid_target_revision'
  | 'missing_idempotency_key'
  | 'invalid_idempotency_key'
  | 'missing_commands'
  | 'unknown_command_kind'
  | 'invalid_command'
  | 'duplicate_command_id'
  | 'invalid_range'
  | 'missing_touched_ranges'
  | 'destructive_not_confirmed'
  | 'too_many_touched_cells'

export interface WorkbookCommandBundleIssue {
  readonly code: WorkbookCommandBundleIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookCommandBundleCheckResult =
  | {
      readonly status: 'valid'
      readonly bundle: WorkbookCommandBundle
      readonly result: WorkbookCommandResult
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookCommandBundleIssue[]
    }

interface NormalizedRangeWithSize {
  readonly range: CellRangeRef
  readonly cellCount: number
}

export function isWorkbookCommandBundleCommandKind(value: unknown): value is WorkbookCommandBundleCommandKind {
  return typeof value === 'string' && WORKBOOK_COMMAND_BUNDLE_COMMAND_KIND_SET.has(value)
}

export function checkWorkbookCommandBundle(value: unknown): WorkbookCommandBundleCheckResult {
  if (!isRecord(value)) {
    return {
      status: 'invalid',
      issues: Object.freeze([commandBundleIssue('invalid_bundle', 'bundle', 'Workbook command bundle must be an object')]),
    }
  }

  const accessorPath = firstAccessorPath(value, 'bundle')
  if (accessorPath !== null) {
    return {
      status: 'invalid',
      issues: Object.freeze([
        commandBundleIssue('invalid_bundle', accessorPath, 'Workbook command bundle must contain only data properties'),
      ]),
    }
  }

  const issues: WorkbookCommandBundleIssue[] = []
  pushOptionalStringIssue(issues, ownValue(value, 'id'), 'id', 'bundle id')
  pushTargetRevisionIssue(issues, ownValue(value, 'targetRevision'))
  pushIdempotencyKeyIssue(issues, ownValue(value, 'idempotencyKey'))
  const scope = ownValue(value, 'scope')
  pushScopeIssues(issues, scope)
  const scopedTouchedRangesRequired = scopeRequiresTouchedRanges(scope)

  const commands = ownValue(value, 'commands')
  if (!Array.isArray(commands) || commands.length === 0) {
    issues.push(commandBundleIssue('missing_commands', 'commands', 'Workbook command bundle commands must be a non-empty array'))
  } else {
    pushArrayDataIssues(issues, commands, 'commands', 'Workbook command bundle commands')
    for (let index = 0; index < commands.length; index += 1) {
      const command = arrayDataValue(commands, index)
      if (command === undefined) {
        continue
      }
      pushCommandIssues(issues, command, `commands[${index}]`, scopedTouchedRangesRequired)
    }
    pushDuplicateCommandIdIssues(issues, commands)
  }

  if (issues.length > 0) {
    return {
      status: 'invalid',
      issues: Object.freeze(issues),
    }
  }

  const normalized = normalizeWorkbookCommandBundleData(value)
  const result = workbookCommandResultFor(normalized)
  const maxTouchedCells = normalized.scope?.maxTouchedCells
  if (maxTouchedCells !== undefined && result.touchedCellCount > maxTouchedCells) {
    return {
      status: 'invalid',
      issues: Object.freeze([
        commandBundleIssue(
          'too_many_touched_cells',
          'scope.maxTouchedCells',
          `Workbook command bundle touches ${result.touchedCellCount} cells, exceeding scope.maxTouchedCells ${maxTouchedCells}`,
        ),
      ]),
    }
  }

  return {
    status: 'valid',
    bundle: normalized,
    result,
    issues: Object.freeze([]),
  }
}

export function normalizeWorkbookCommandBundle(value: unknown): WorkbookCommandBundle {
  const check = checkWorkbookCommandBundle(value)
  if (check.status === 'invalid') {
    const [firstIssue] = check.issues
    throw new Error(
      firstIssue === undefined ? 'Workbook command bundle is invalid' : `Workbook command bundle is invalid: ${firstIssue.message}`,
    )
  }
  return check.bundle
}

export function isWorkbookCommandBundle(value: unknown): value is WorkbookCommandBundle {
  return checkWorkbookCommandBundle(value).status === 'valid'
}

function pushCommandIssues(issues: WorkbookCommandBundleIssue[], value: unknown, path: string, scopedTouchedRangesRequired: boolean): void {
  if (!isRecord(value)) {
    issues.push(commandBundleIssue('invalid_command', path, 'Workbook command bundle command must be an object'))
    return
  }

  pushOptionalStringIssue(issues, ownValue(value, 'id'), `${path}.id`, 'command id')
  const kind = ownValue(value, 'kind')
  if (!isWorkbookCommandBundleCommandKind(kind)) {
    issues.push(commandBundleIssue('unknown_command_kind', `${path}.kind`, 'Workbook command bundle command kind is unknown'))
    return
  }

  const touchedRanges = ownValue(value, 'touchedRanges')
  pushTouchedRangesIssues(issues, touchedRanges, `${path}.touchedRanges`)

  const destructive = ownValue(value, 'destructive')
  if (destructive !== undefined && typeof destructive !== 'boolean') {
    issues.push(
      commandBundleIssue('invalid_command', `${path}.destructive`, 'Workbook command bundle command destructive flag must be boolean'),
    )
  }

  if (kind === 'request') {
    pushRequestCommandIssues(issues, value, path, touchedRanges, scopedTouchedRangesRequired)
    return
  }
  pushOpCommandIssues(issues, value, path, touchedRanges, scopedTouchedRangesRequired)
}

function pushRequestCommandIssues(
  issues: WorkbookCommandBundleIssue[],
  value: Record<string, unknown>,
  path: string,
  touchedRanges: unknown,
  scopedTouchedRangesRequired: boolean,
): void {
  const request = ownValue(value, 'request')
  const requestCheck = checkWorkbookCommandRequest(request)
  if (requestCheck.status === 'invalid') {
    requestCheck.issues.forEach((issue) => {
      issues.push(commandBundleIssue('invalid_command', `${path}.request.${issue.path}`, issue.message))
    })
    return
  }

  const destructive = commandNeedsDestructiveConfirmation(requestCheck.request)
  if (destructive && ownValue(value, 'destructive') !== true) {
    issues.push(
      commandBundleIssue(
        'destructive_not_confirmed',
        `${path}.destructive`,
        'Workbook command bundle request mutates the workbook and must set destructive: true',
      ),
    )
  }
  if (destructive && scopedTouchedRangesRequired) {
    pushScopedTouchedRangesIssue(issues, touchedRanges, `${path}.touchedRanges`)
  }
}

function pushOpCommandIssues(
  issues: WorkbookCommandBundleIssue[],
  value: Record<string, unknown>,
  path: string,
  touchedRanges: unknown,
  scopedTouchedRangesRequired: boolean,
): void {
  if (!isWorkbookOp(ownValue(value, 'op'))) {
    issues.push(commandBundleIssue('invalid_command', `${path}.op`, 'Workbook command bundle op is invalid'))
  }
  if (ownValue(value, 'destructive') !== true) {
    issues.push(
      commandBundleIssue(
        'destructive_not_confirmed',
        `${path}.destructive`,
        'Workbook command bundle op mutates the workbook and must set destructive: true',
      ),
    )
  }
  if (scopedTouchedRangesRequired) {
    pushScopedTouchedRangesIssue(issues, touchedRanges, `${path}.touchedRanges`)
  }
}

function commandNeedsDestructiveConfirmation(request: WorkbookCommandRequest): boolean {
  return request.category === 'mutation' || request.mode === 'apply' || request.mode === 'applyAndVerify'
}

function pushDuplicateCommandIdIssues(issues: WorkbookCommandBundleIssue[], commands: readonly unknown[]): void {
  const seen = new Map<string, number>()
  for (let index = 0; index < commands.length; index += 1) {
    const command = arrayDataValue(commands, index)
    if (!isRecord(command)) {
      continue
    }
    const id = ownValue(command, 'id')
    if (typeof id !== 'string' || id.trim() === '' || id.trim() !== id) {
      continue
    }
    const previousIndex = seen.get(id)
    if (previousIndex !== undefined) {
      issues.push(
        commandBundleIssue(
          'duplicate_command_id',
          `commands[${index}].id`,
          `Workbook command bundle command id ${id} already used by commands[${previousIndex}].id`,
        ),
      )
      continue
    }
    seen.set(id, index)
  }
}

function scopeRequiresTouchedRanges(value: unknown): boolean {
  return isRecord(value) && isSafeNonNegativeInteger(ownValue(value, 'maxTouchedCells'))
}

function pushScopedTouchedRangesIssue(issues: WorkbookCommandBundleIssue[], value: unknown, path: string): void {
  if (!Array.isArray(value) || value.length === 0) {
    issues.push(
      commandBundleIssue(
        'missing_touched_ranges',
        path,
        'Scoped destructive workbook command must declare touchedRanges so scope.maxTouchedCells is enforceable',
      ),
    )
  }
}

function pushTouchedRangesIssues(issues: WorkbookCommandBundleIssue[], value: unknown, path: string): void {
  if (value === undefined) {
    return
  }
  if (!Array.isArray(value)) {
    issues.push(commandBundleIssue('invalid_range', path, 'Workbook command bundle touched ranges must be an array'))
    return
  }
  pushArrayDataIssues(issues, value, path, 'Workbook command bundle touched ranges')
  for (let index = 0; index < value.length; index += 1) {
    const range = arrayDataValue(value, index)
    if (range === undefined) {
      continue
    }
    try {
      normalizeRange(range, `${path}[${index}]`)
    } catch (error) {
      issues.push(commandBundleIssue('invalid_range', `${path}[${index}]`, errorMessage(error)))
    }
  }
}

function pushScopeIssues(issues: WorkbookCommandBundleIssue[], value: unknown): void {
  if (value === undefined) {
    return
  }
  if (!isRecord(value)) {
    issues.push(commandBundleIssue('invalid_bundle', 'scope', 'Workbook command bundle scope must be an object'))
    return
  }
  const maxTouchedCells = ownValue(value, 'maxTouchedCells')
  if (maxTouchedCells !== undefined && !isSafeNonNegativeInteger(maxTouchedCells)) {
    issues.push(
      commandBundleIssue(
        'invalid_bundle',
        'scope.maxTouchedCells',
        'Workbook command bundle maxTouchedCells must be a safe non-negative integer',
      ),
    )
  }
}

function pushTargetRevisionIssue(issues: WorkbookCommandBundleIssue[], value: unknown): void {
  if (value === undefined) {
    issues.push(commandBundleIssue('missing_target_revision', 'targetRevision', 'Workbook command bundle targetRevision is required'))
    return
  }
  if (!isSafeNonNegativeInteger(value)) {
    issues.push(
      commandBundleIssue(
        'invalid_target_revision',
        'targetRevision',
        'Workbook command bundle targetRevision must be a safe non-negative integer',
      ),
    )
  }
}

function pushIdempotencyKeyIssue(issues: WorkbookCommandBundleIssue[], value: unknown): void {
  if (value === undefined) {
    issues.push(commandBundleIssue('missing_idempotency_key', 'idempotencyKey', 'Workbook command bundle idempotencyKey is required'))
    return
  }
  if (typeof value !== 'string') {
    issues.push(commandBundleIssue('invalid_idempotency_key', 'idempotencyKey', 'Workbook command bundle idempotencyKey must be a string'))
    return
  }
  pushNonEmptyExactStringIssue(issues, value, 'idempotencyKey', 'idempotencyKey', 'invalid_idempotency_key')
}

function pushOptionalStringIssue(issues: WorkbookCommandBundleIssue[], value: unknown, path: string, label: string): void {
  if (value === undefined) {
    return
  }
  if (typeof value !== 'string') {
    issues.push(commandBundleIssue('invalid_bundle', path, `Workbook command bundle ${label} must be a string`))
    return
  }
  pushNonEmptyExactStringIssue(issues, value, path, label, 'invalid_bundle')
}

function pushNonEmptyExactStringIssue(
  issues: WorkbookCommandBundleIssue[],
  value: string,
  path: string,
  label: string,
  code: WorkbookCommandBundleIssueCode,
): void {
  const normalized = value.trim()
  if (normalized === '') {
    issues.push(commandBundleIssue(code, path, `Workbook command bundle ${label} cannot be empty`))
    return
  }
  if (normalized !== value) {
    issues.push(commandBundleIssue(code, path, `Workbook command bundle ${label} must not have leading or trailing whitespace`))
  }
}

function pushArrayDataIssues(issues: WorkbookCommandBundleIssue[], value: readonly unknown[], path: string, label: string): void {
  const accessorPath = firstAccessorPath(value, path)
  if (accessorPath !== null) {
    issues.push(commandBundleIssue('invalid_bundle', accessorPath, `${label} must contain only data properties`))
    return
  }
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      issues.push(commandBundleIssue('invalid_bundle', `${path}[${index}]`, `${label} must contain only data properties`))
    }
  }
}

function normalizeWorkbookCommandBundleData(value: Record<string, unknown>): WorkbookCommandBundle {
  const id = ownValue(value, 'id')
  const targetRevision = ownValue(value, 'targetRevision')
  const idempotencyKey = ownValue(value, 'idempotencyKey')
  const scope = ownValue(value, 'scope')
  const maxTouchedCells = isRecord(scope) ? ownValue(scope, 'maxTouchedCells') : undefined
  const commands = ownValue(value, 'commands')
  if (!isSafeNonNegativeInteger(targetRevision)) {
    throw new Error('Workbook command bundle targetRevision is invalid')
  }
  if (typeof idempotencyKey !== 'string') {
    throw new Error('Workbook command bundle idempotencyKey is invalid')
  }
  if (!Array.isArray(commands)) {
    throw new Error('Workbook command bundle commands are invalid')
  }

  return Object.freeze({
    ...(typeof id === 'string' ? { id } : {}),
    targetRevision,
    idempotencyKey,
    ...(isSafeNonNegativeInteger(maxTouchedCells) ? { scope: Object.freeze({ maxTouchedCells }) } : {}),
    commands: Object.freeze(commands.map((command, index) => normalizeCommand(command, index))),
  })
}

function normalizeCommand(value: unknown, index: number): WorkbookCommandBundleCommand {
  if (!isRecord(value)) {
    throw new Error(`Workbook command bundle command ${index} is invalid`)
  }
  const id = ownValue(value, 'id')
  const kind = ownValue(value, 'kind')
  const touchedRanges = normalizeTouchedRanges(ownValue(value, 'touchedRanges'))
  const destructive = ownValue(value, 'destructive')
  const common = Object.freeze({
    ...(typeof id === 'string' ? { id } : {}),
    ...(touchedRanges.length > 0 ? { touchedRanges } : {}),
    ...(destructive !== undefined ? { destructive: destructive === true } : {}),
  })

  if (kind === 'request') {
    return Object.freeze({
      ...common,
      kind,
      request: normalizeWorkbookCommandRequest(ownValue(value, 'request')),
    })
  }

  return Object.freeze({
    ...common,
    kind: 'op',
    op: normalizeOp(ownValue(value, 'op')),
  })
}

function normalizeOp(value: unknown): EngineOp {
  if (!isWorkbookOp(value)) {
    throw new Error('Workbook command bundle op is invalid')
  }
  const cloned = cloneData(value)
  if (!isWorkbookOp(cloned)) {
    throw new Error('Workbook command bundle op clone is invalid')
  }
  return freezeData(cloned)
}

function normalizeTouchedRanges(value: unknown): readonly CellRangeRef[] {
  if (!Array.isArray(value)) {
    return Object.freeze([])
  }
  return Object.freeze(value.map((range, index) => normalizeRange(range, `touchedRanges[${index}]`).range))
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

function normalizeExactString(value: string, path: string): string {
  const normalized = value.trim()
  if (normalized === '') {
    throw new Error(`Workbook command bundle ${path} cannot be empty`)
  }
  if (normalized !== value) {
    throw new Error(`Workbook command bundle ${path} must not have leading or trailing whitespace`)
  }
  return normalized
}

function commandBundleIssue(code: WorkbookCommandBundleIssueCode, path: string, message: string): WorkbookCommandBundleIssue {
  return Object.freeze({
    code,
    path,
    message,
  })
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
