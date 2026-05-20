import type { ReplicaState } from '../replica-state.js'
import type { SpreadsheetEngine } from '../engine.js'
import type { EngineOperationService } from '../engine/services/operation-service.js'

interface FormulaBindingNowService {
  bindFormulaNow: (...args: unknown[]) => boolean
  bindInitialFormulaNow: (...args: unknown[]) => void
  bindPreparedFormulaNow: (...args: unknown[]) => boolean
}

function isEngineOperationService(value: unknown): value is EngineOperationService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'applyBatch') === 'function' && typeof Reflect.get(value, 'applyDerivedOp') === 'function'
}

function isFormulaBindingNowService(value: unknown): value is FormulaBindingNowService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'bindFormulaNow') === 'function' &&
    typeof Reflect.get(value, 'bindInitialFormulaNow') === 'function' &&
    typeof Reflect.get(value, 'bindPreparedFormulaNow') === 'function'
  )
}

function isReplicaState(value: unknown): value is ReplicaState {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'replicaId') === 'string' &&
    typeof Reflect.get(value, 'clock') === 'object' &&
    Reflect.get(value, 'appliedBatchIds') instanceof Set
  )
}

export function getOperationService(engine: SpreadsheetEngine): EngineOperationService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (!isEngineOperationService(operations)) {
    throw new TypeError('Expected engine operation service')
  }
  return operations
}

export function getFormulaBindingNowService(engine: SpreadsheetEngine): FormulaBindingNowService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const binding = Reflect.get(runtime, 'binding')
  if (!isFormulaBindingNowService(binding)) {
    throw new TypeError('Expected formula binding service')
  }
  return binding
}

export function getReplicaState(engine: SpreadsheetEngine): ReplicaState {
  const replicaState = Reflect.get(engine, 'replicaState')
  if (!isReplicaState(replicaState)) {
    throw new TypeError('Expected engine replica state')
  }
  return replicaState
}

export function findErrorByName(error: unknown, name: string): Error | undefined {
  let current: unknown = error
  let depth = 0
  while (typeof current === 'object' && current !== null && depth < 16) {
    if (current instanceof Error && current.name === name) {
      return current
    }
    current = Reflect.get(current, 'cause')
    depth += 1
  }
  return undefined
}
