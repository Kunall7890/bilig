import { isWorkbookOp } from './guards.js'
import { commandOpsMatchExpected, workbookOpsMatch } from './command-ops.js'
import { isWorkbookRefData, toWorkbookRefData } from './find.js'
import { materializeFormulaLabels, type WorkbookFormulaLabelReplacement } from './formula-usage.js'
import { normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import type { WorkbookActionCommand, WorkbookActionPlan } from './model.js'
import type { EngineOp } from './ops.js'
import type {
  WorkbookCommandResolvedRefs,
  WorkbookResolvedRefData,
  WorkbookResolvedRefValue,
  WorkbookRunApplyCommandReceipt,
  WorkbookRunNoopProof,
} from './result.js'

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function keyName(key: string | symbol): string {
  return typeof key === 'symbol' ? key.toString() : key
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return Number.isSafeInteger(value) && typeof value === 'number' && value >= 0
}

function ownValue(value: object, key: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return undefined
  }
  if (!('value' in descriptor)) {
    throw new Error(`${key} must be a data property`)
  }
  return descriptor.value
}

function arrayDataValues<T>(value: readonly T[], guard: (entry: unknown) => entry is T): readonly T[] | null {
  if (!Array.isArray(value)) {
    return null
  }
  const entries: T[] = []
  const descriptors = Object.getOwnPropertyDescriptors(value)
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = descriptors[String(index)]
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor) || !guard(descriptor.value)) {
      return null
    }
    entries.push(descriptor.value)
  }
  return entries
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    const descriptors = Object.getOwnPropertyDescriptors(value)
    return Array.from({ length: value.length }, (_, index) => {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined) {
        return undefined
      }
      if (!('value' in descriptor)) {
        throw new Error('Accessor values cannot be canonicalized')
      }
      return canonicalValue(descriptor.value)
    })
  }
  if (typeof value === 'object' && value !== null) {
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

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

function fnv1a64(input: string, seed: bigint): string {
  let hash = seed
  for (let index = 0; index < input.length; index += 1) {
    hash ^= BigInt(input.charCodeAt(index))
    hash = BigInt.asUintN(64, hash * 0x100000001b3n)
  }
  return hash.toString(16).padStart(16, '0')
}

export function workbookActionCommandDigest(command: WorkbookActionCommand): string {
  const json = canonicalJson(command)
  return `bilig-command-v1:${fnv1a64(json, 0xcbf29ce484222325n)}${fnv1a64(json, 0x84222325cbf29cen)}`
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
  const cloned: Record<string, unknown> = {}
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

function cloneOp(op: EngineOp): EngineOp {
  const cloned = cloneData(op)
  if (!isWorkbookOp(cloned)) {
    throw new Error('invalid workbook op clone')
  }
  return cloned
}

function cloneOps(ops: readonly EngineOp[]): readonly EngineOp[] {
  const entries = arrayDataValues(ops, isWorkbookOp)
  if (entries === null) {
    throw new Error('invalid workbook op array')
  }
  return Object.freeze(entries.map((op) => cloneOp(op)))
}

function cloneReceiptOps(receipt: object, receiptIndex: number, key: 'previewOps' | 'appliedOps'): readonly EngineOp[] {
  const ops = ownValue(receipt, key)
  if (ops === undefined) {
    throw new Error(`commandReceipts[${String(receiptIndex)}].${key} is required`)
  }
  if (!Array.isArray(ops) || arrayDataValues(ops, isWorkbookOp) === null) {
    throw new Error(`commandReceipts[${String(receiptIndex)}].${key} is invalid`)
  }
  return cloneOps(ops)
}

function cloneFormulaLabelReplacements(receipt: object, receiptIndex: number): readonly WorkbookFormulaLabelReplacement[] | undefined {
  const value = ownValue(receipt, 'formulaLabels')
  if (value === undefined) {
    return undefined
  }
  if (!Array.isArray(value)) {
    throw new Error(`commandReceipts[${String(receiptIndex)}].formulaLabels is invalid`)
  }
  const entries = arrayDataValues(value, isRecord)
  if (entries === null) {
    throw new Error(`commandReceipts[${String(receiptIndex)}].formulaLabels is invalid`)
  }
  return Object.freeze(
    entries.map((entry, labelIndex) => {
      const name = ownValue(entry, 'name')
      const source = ownValue(entry, 'source')
      if (typeof name !== 'string' || typeof source !== 'string') {
        throw new Error(`commandReceipts[${String(receiptIndex)}].formulaLabels[${String(labelIndex)}] is invalid`)
      }
      let normalizedSource: string
      try {
        normalizedSource = materializeFormulaLabels(name, [{ name, source }])
      } catch (error) {
        throw new Error(
          `commandReceipts[${String(receiptIndex)}].formulaLabels[${String(labelIndex)}] is invalid: ${errorMessage(error)}`,
          { cause: error },
        )
      }
      return Object.freeze({
        name,
        source: normalizedSource,
      })
    }),
  )
}

function cloneResolvedRangeRef(value: unknown, path: string): WorkbookResolvedRefData {
  if (!isWorkbookRefData(value)) {
    throw new Error(`${path} must be workbook ref data`)
  }
  const ref = toWorkbookRefData(value)
  if (ref.kind !== 'range') {
    throw new Error(`${path} must be concrete range ref data`)
  }
  return ref
}

function cloneResolvedRefValue(value: unknown, path: string): WorkbookResolvedRefValue {
  if (Array.isArray(value)) {
    const entries = arrayDataValues(value, isRecord)
    if (entries === null) {
      throw new Error(`${path} must contain only concrete range ref data objects`)
    }
    return Object.freeze(entries.map((entry, index) => cloneResolvedRangeRef(entry, `${path}[${String(index)}]`)))
  }
  return cloneResolvedRangeRef(value, path)
}

function cloneResolvedInputs(value: unknown, path: string): readonly WorkbookResolvedRefValue[] {
  if (!Array.isArray(value)) {
    throw new Error(`${path} must be an array`)
  }
  const entries: WorkbookResolvedRefValue[] = []
  const descriptors = Object.getOwnPropertyDescriptors(value)
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = descriptors[String(index)]
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      throw new Error(`${path}[${String(index)}] must be a data property`)
    }
    entries.push(cloneResolvedRefValue(descriptor.value, `${path}[${String(index)}]`))
  }
  return Object.freeze(entries)
}

function cloneCommandResolvedRefs(receipt: object, receiptIndex: number): WorkbookCommandResolvedRefs | undefined {
  const value = ownValue(receipt, 'resolvedRefs')
  if (value === undefined) {
    return undefined
  }
  if (!isRecord(value)) {
    throw new Error(`commandReceipts[${String(receiptIndex)}].resolvedRefs must be an object`)
  }
  const target = ownValue(value, 'target')
  const inputs = ownValue(value, 'inputs')
  return Object.freeze({
    ...(target !== undefined
      ? { target: cloneResolvedRefValue(target, `commandReceipts[${String(receiptIndex)}].resolvedRefs.target`) }
      : {}),
    ...(inputs !== undefined
      ? { inputs: cloneResolvedInputs(inputs, `commandReceipts[${String(receiptIndex)}].resolvedRefs.inputs`) }
      : {}),
  })
}

function inputMatches(left: unknown, right: unknown): boolean {
  return JSON.stringify(normalizeWorkbookActionInput(left)) === JSON.stringify(normalizeWorkbookActionInput(right))
}

function assertNoopEffectMatchesCommand(command: WorkbookActionCommand, proof: object, path: string): void {
  const effect = ownValue(proof, 'effect')
  if (!isRecord(effect)) {
    throw new Error(`${path}.proof.effect must be an object`)
  }
  if (ownValue(effect, 'kind') !== command.kind) {
    throw new Error(`${path}.proof.effect.kind must match commandKind`)
  }
  if (command.kind === 'writeValue' && !inputMatches(ownValue(effect, 'value'), command.value)) {
    throw new Error(`${path}.proof.effect.value must match command value`)
  }
  if (command.kind === 'writeFormula' && ownValue(effect, 'formula') !== command.formula) {
    throw new Error(`${path}.proof.effect.formula must match command formula`)
  }
  if (command.kind === 'clear' && ownValue(effect, 'cleared') !== true) {
    throw new Error(`${path}.proof.effect.cleared must be true`)
  }
  if (command.kind === 'format') {
    if (command.style !== undefined && !inputMatches(ownValue(effect, 'style'), command.style)) {
      throw new Error(`${path}.proof.effect.style must match command style`)
    }
    if (command.numberFormat !== undefined && ownValue(effect, 'numberFormat') !== command.numberFormat) {
      throw new Error(`${path}.proof.effect.numberFormat must match command numberFormat`)
    }
  }
  if (command.kind === 'op' && ownValue(effect, 'opKind') !== command.op.kind) {
    throw new Error(`${path}.proof.effect.opKind must match command op kind`)
  }
}

function cloneNoopProof(
  receipt: object,
  receiptIndex: number,
  expected: { readonly command?: WorkbookActionCommand; readonly commandKind: string; readonly commandDigest: string },
): WorkbookRunNoopProof | undefined {
  const value = ownValue(receipt, 'noop')
  if (value === undefined) {
    return undefined
  }
  const path = `commandReceipts[${String(receiptIndex)}].noop`
  if (!isRecord(value)) {
    throw new Error(`${path} must be an object`)
  }
  for (const key of Reflect.ownKeys(value)) {
    const name = keyName(key)
    if (typeof key === 'symbol' || (key !== 'reason' && key !== 'message' && key !== 'proof')) {
      throw new Error(`${path}.${name} is unknown`)
    }
  }
  const reason = ownValue(value, 'reason')
  if (reason !== 'already_satisfied') {
    throw new Error(`${path}.reason is invalid`)
  }
  const message = ownValue(value, 'message')
  if (message !== undefined && (typeof message !== 'string' || message.trim() === '')) {
    throw new Error(`${path}.message is invalid`)
  }
  const rawProof = ownValue(value, 'proof')
  if (!isRecord(rawProof)) {
    throw new Error(`${path}.proof must be an object`)
  }
  for (const key of Reflect.ownKeys(rawProof)) {
    if (typeof key === 'symbol') {
      throw new Error(`${path}.proof.${keyName(key)} is unknown`)
    }
  }
  const source = ownValue(rawProof, 'source')
  if (typeof source !== 'string' || source.trim() === '') {
    throw new Error(`${path}.proof.source is invalid`)
  }
  const evidence = ownValue(rawProof, 'evidence')
  if (typeof evidence !== 'string' || evidence.trim() === '') {
    throw new Error(`${path}.proof.evidence is invalid`)
  }
  if (ownValue(rawProof, 'opCount') !== 0) {
    throw new Error(`${path}.proof.opCount must be 0`)
  }
  if (ownValue(rawProof, 'commandKind') !== expected.commandKind) {
    throw new Error(`${path}.proof.commandKind must match commandKind`)
  }
  if (ownValue(rawProof, 'commandDigest') !== expected.commandDigest) {
    throw new Error(`${path}.proof.commandDigest must match commandDigest`)
  }
  if (expected.command !== undefined) {
    assertNoopEffectMatchesCommand(expected.command, rawProof, path)
  }
  let proof: WorkbookActionInput
  try {
    proof = normalizeWorkbookActionInput(rawProof)
  } catch (error) {
    throw new Error(`${path}.proof is invalid: ${errorMessage(error)}`, { cause: error })
  }
  return Object.freeze({
    reason,
    ...(message !== undefined ? { message } : {}),
    proof,
  })
}

function flattenedReceiptOps(receipts: readonly WorkbookRunApplyCommandReceipt[], key: 'previewOps' | 'appliedOps'): readonly EngineOp[] {
  const ops: EngineOp[] = []
  for (const receipt of receipts) {
    for (const op of receipt[key]) {
      ops.push(op)
    }
  }
  return ops
}

export function cloneWorkbookRunApplyCommandReceipts(
  plan: WorkbookActionPlan,
  value: unknown,
  previewOps: readonly EngineOp[] | undefined,
  appliedOps: readonly EngineOp[] | undefined,
): readonly WorkbookRunApplyCommandReceipt[] {
  if (!Array.isArray(value)) {
    throw new Error('commandReceipts must be an array')
  }
  const entries = arrayDataValues(value, isRecord)
  if (entries === null) {
    throw new Error('commandReceipts must contain only data objects')
  }
  if (entries.length !== plan.commands.length) {
    throw new Error(`expected ${String(plan.commands.length)} command receipts, got ${String(entries.length)}`)
  }

  const seenCommandIndexes = new Set<number>()
  const receipts = entries.map((receipt, receiptIndex) => {
    const commandIndex = ownValue(receipt, 'commandIndex')
    if (!isSafeNonNegativeInteger(commandIndex) || commandIndex >= plan.commands.length) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].commandIndex is invalid`)
    }
    if (seenCommandIndexes.has(commandIndex)) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].commandIndex is duplicated`)
    }
    seenCommandIndexes.add(commandIndex)

    const command = plan.commands[commandIndex]
    if (command === undefined) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].commandIndex does not exist`)
    }

    const commandKind = ownValue(receipt, 'commandKind')
    if (typeof commandKind !== 'string') {
      throw new Error(`commandReceipts[${String(receiptIndex)}].commandKind is invalid`)
    }
    if (commandKind !== command.kind) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].commandKind does not match the planned command`)
    }

    const commandDigest = ownValue(receipt, 'commandDigest')
    if (commandDigest !== workbookActionCommandDigest(command)) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].commandDigest does not match the planned command`)
    }

    const receiptPreviewOps = cloneReceiptOps(receipt, receiptIndex, 'previewOps')
    const receiptAppliedOps = cloneReceiptOps(receipt, receiptIndex, 'appliedOps')
    const formulaLabels = command.kind === 'writeFormula' ? cloneFormulaLabelReplacements(receipt, receiptIndex) : undefined
    if (!workbookOpsMatch(receiptPreviewOps, receiptAppliedOps)) {
      throw new Error(`commandReceipts[${String(receiptIndex)}] previewOps do not match appliedOps`)
    }
    const resolvedRefs = cloneCommandResolvedRefs(receipt, receiptIndex)
    if (!commandOpsMatchExpected(command, receiptPreviewOps, formulaLabels, resolvedRefs)) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].previewOps do not match the planned command`)
    }
    if (!commandOpsMatchExpected(command, receiptAppliedOps, formulaLabels, resolvedRefs)) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].appliedOps do not match the planned command`)
    }
    const noop = cloneNoopProof(receipt, receiptIndex, { command, commandKind, commandDigest })
    if (noop !== undefined && receiptAppliedOps.length > 0) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].noop is only valid for commands with no applied ops`)
    }

    let proof: WorkbookActionInput | undefined
    const rawProof = ownValue(receipt, 'proof')
    try {
      proof = rawProof === undefined ? undefined : normalizeWorkbookActionInput(rawProof)
    } catch (error) {
      throw new Error(`commandReceipts[${String(receiptIndex)}].proof is invalid: ${errorMessage(error)}`, { cause: error })
    }

    return Object.freeze({
      commandIndex,
      commandKind,
      commandDigest,
      previewOps: receiptPreviewOps,
      appliedOps: receiptAppliedOps,
      ...(noop !== undefined ? { noop } : {}),
      ...(resolvedRefs !== undefined ? { resolvedRefs } : {}),
      ...(formulaLabels !== undefined ? { formulaLabels } : {}),
      ...(proof !== undefined ? { proof } : {}),
    })
  })

  const sortedReceipts = Object.freeze(receipts.toSorted((left, right) => left.commandIndex - right.commandIndex))
  if (previewOps !== undefined && !workbookOpsMatch(flattenedReceiptOps(sortedReceipts, 'previewOps'), previewOps)) {
    throw new Error('commandReceipts previewOps do not match apply previewOps')
  }
  if (appliedOps !== undefined && !workbookOpsMatch(flattenedReceiptOps(sortedReceipts, 'appliedOps'), appliedOps)) {
    throw new Error('commandReceipts appliedOps do not match apply appliedOps')
  }

  return sortedReceipts
}

export function cloneWorkbookRunApplyCommandReceiptsForSummary(
  receipts: readonly WorkbookRunApplyCommandReceipt[],
): readonly WorkbookRunApplyCommandReceipt[] {
  return Object.freeze(
    receipts.map((receipt, index) => {
      const formulaLabels = receipt.formulaLabels !== undefined ? cloneFormulaLabelReplacements(receipt, index) : undefined
      const resolvedRefs =
        receipt.resolvedRefs === undefined ? undefined : cloneCommandResolvedRefs({ resolvedRefs: receipt.resolvedRefs }, index)
      const noop =
        receipt.noop === undefined
          ? undefined
          : cloneNoopProof({ noop: receipt.noop }, index, { commandKind: receipt.commandKind, commandDigest: receipt.commandDigest })
      return Object.freeze({
        commandIndex: receipt.commandIndex,
        commandKind: receipt.commandKind,
        commandDigest: receipt.commandDigest,
        previewOps: cloneOps(receipt.previewOps),
        appliedOps: cloneOps(receipt.appliedOps),
        ...(noop !== undefined ? { noop } : {}),
        ...(resolvedRefs !== undefined ? { resolvedRefs } : {}),
        ...(formulaLabels !== undefined ? { formulaLabels } : {}),
        ...(receipt.proof !== undefined ? { proof: normalizeWorkbookActionInput(receipt.proof) } : {}),
      })
    }),
  )
}
