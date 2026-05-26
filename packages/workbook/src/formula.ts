import { normalizeFormulaFunctionName, parseFormula } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import { isObjectRecord } from './data-properties.js'
import { isWorkbookRef, type WorkbookRef } from './find.js'

export interface WorkbookFormulaExpression {
  readonly kind: 'formula'
  readonly source: string
  readonly inputs: readonly WorkbookRef[]
  readonly labels: readonly WorkbookFormulaLabel[]
}

export interface WorkbookFormulaLabel {
  readonly name: string
  readonly ref: WorkbookRef
}

export interface WorkbookRawFormulaOptions {
  readonly inputs?: readonly WorkbookRef[]
  readonly labels?: readonly WorkbookFormulaLabel[]
}

export type WorkbookFormulaOperand = WorkbookFormulaExpression | WorkbookRef | number | boolean

const FORMULA_OPERAND_ERROR =
  'Formula operands must be formula expressions, workbook refs, finite numbers, booleans, or formula.text/raw wrappers'
const EMPTY_FORMULA_INPUTS = Object.freeze([] as readonly WorkbookRef[])
const EMPTY_FORMULA_LABELS = Object.freeze([] as readonly WorkbookFormulaLabel[])

function dataValue(value: object, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function assertOnlyDataProperties(value: unknown, path: string, seen = new WeakSet<object>()): void {
  if (typeof value !== 'object' || value === null) {
    return
  }
  if (seen.has(value)) {
    return
  }
  seen.add(value)
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key === 'symbol') {
      throw new Error(`${path}[${String(key)}] must be a data property`)
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    const child = Array.isArray(value) && /^\d+$/.test(key) ? `${path}[${key}]` : `${path}.${key}`
    if (descriptor === undefined || !('value' in descriptor)) {
      throw new Error(`${child} must be a data property`)
    }
    assertOnlyDataProperties(descriptor.value, child, seen)
  }
}

function dataArrayValues(value: unknown, path: string): readonly unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`${path} must be an array`)
  }
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new Error(`${path} must be a plain array`)
  }
  assertOnlyDataProperties(value, path)
  const entries: unknown[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      throw new Error(`${path}[${index}] must be a data property`)
    }
    entries.push(descriptor.value)
  }
  return entries
}

function formulaText(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`${label} must be a string`)
  }
  return value
}

function normalizeFormulaSource(source: unknown): string {
  const sourceText = formulaText(source, 'Formula source')
  const trimmed = sourceText.trim()
  if (trimmed === '') {
    throw new Error('Formula source cannot be empty')
  }
  const normalized = trimmed.startsWith('=') ? trimmed.slice(1) : trimmed
  parseFormula(normalized)
  return normalized
}

function formulaFunctionName(value: unknown): string {
  return normalizeFormulaFunctionName(formulaText(value, 'Formula function name'))
}

function formulaRef(value: unknown, path: string): WorkbookRef {
  if (!isWorkbookRef(value)) {
    throw new Error(`${path} must be a workbook ref`)
  }
  return value
}

function formulaRefArray(value: unknown, path: string): readonly WorkbookRef[] {
  return Object.freeze(dataArrayValues(value, path).map((entry, index) => formulaRef(entry, `${path}[${index}]`)))
}

function formulaLabel(value: unknown, path: string): WorkbookFormulaLabel {
  if (!isObjectRecord(value)) {
    throw new Error(`${path} must be a formula label`)
  }
  assertOnlyDataProperties(value, path)
  return Object.freeze({
    name: formulaText(dataValue(value, 'name'), `${path}.name`),
    ref: formulaRef(dataValue(value, 'ref'), `${path}.ref`),
  })
}

function formulaLabelArray(value: unknown, path: string): readonly WorkbookFormulaLabel[] {
  return Object.freeze(dataArrayValues(value, path).map((entry, index) => formulaLabel(entry, `${path}[${index}]`)))
}

function isFormulaOperand(value: unknown): value is WorkbookFormulaOperand {
  return isFormulaExpression(value) || isWorkbookRef(value) || typeof value === 'number' || typeof value === 'boolean'
}

function formulaOperands(value: unknown, path: string): readonly WorkbookFormulaOperand[] {
  return Object.freeze(
    dataArrayValues(value, path).map((entry) => {
      if (!isFormulaOperand(entry)) {
        throw new Error(FORMULA_OPERAND_ERROR)
      }
      return entry
    }),
  )
}

function rawFormulaOptions(options: unknown): {
  readonly inputs: readonly WorkbookRef[]
  readonly labels?: readonly WorkbookFormulaLabel[]
} {
  if (!isObjectRecord(options)) {
    throw new Error('Formula raw options must be an object')
  }
  assertOnlyDataProperties(options, 'Formula raw options')
  const inputs = dataValue(options, 'inputs')
  const labels = dataValue(options, 'labels')
  return Object.freeze({
    inputs: inputs === undefined ? Object.freeze([]) : formulaRefArray(inputs, 'Formula raw options.inputs'),
    ...(labels === undefined ? {} : { labels: formulaLabelArray(labels, 'Formula raw options.labels') }),
  })
}

function uniqueRefs(refs: readonly WorkbookRef[]): readonly WorkbookRef[] {
  const seen = new Set<string>()
  const unique: WorkbookRef[] = []
  for (const ref of refs) {
    const key = `${ref.kind}:${ref.id}`
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    unique.push(ref)
  }
  return Object.freeze(unique)
}

function uniqueLabels(labels: readonly WorkbookFormulaLabel[]): readonly WorkbookFormulaLabel[] {
  const seen = new Set<string>()
  const unique: WorkbookFormulaLabel[] = []
  for (const label of labels) {
    if (label.name.trim() === '') {
      throw new Error('Formula label name cannot be empty')
    }
    const key = `${label.name}:${label.ref.kind}:${label.ref.id}`
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    unique.push(
      Object.freeze({
        name: label.name,
        ref: label.ref,
      }),
    )
  }
  return Object.freeze(unique)
}

function labelForRef(ref: WorkbookRef): WorkbookFormulaLabel {
  return Object.freeze({
    name: refSource(ref),
    ref,
  })
}

function labelsForRefs(refs: readonly WorkbookRef[]): readonly WorkbookFormulaLabel[] {
  return Object.freeze(refs.map(labelForRef))
}

function createFormulaExpression(
  source: string,
  inputs: readonly WorkbookRef[] = [],
  labels: readonly WorkbookFormulaLabel[] = labelsForRefs(inputs),
): WorkbookFormulaExpression {
  const normalizedLabels = uniqueLabels(labels)
  const normalizedInputs = uniqueRefs([...inputs, ...normalizedLabels.map((label) => label.ref)])
  return Object.freeze({
    kind: 'formula',
    source: normalizeFormulaSource(source),
    inputs: normalizedInputs,
    labels: normalizedLabels,
  })
}

function isFormulaExpression(value: unknown): value is WorkbookFormulaExpression {
  return (
    isObjectRecord(value) &&
    dataValue(value, 'kind') === 'formula' &&
    typeof dataValue(value, 'source') === 'string' &&
    Array.isArray(dataValue(value, 'inputs')) &&
    Array.isArray(dataValue(value, 'labels'))
  )
}

function expressionSource(expression: WorkbookFormulaExpression): string {
  return normalizeFormulaSource(dataValue(expression, 'source'))
}

function formulaExpressionInputs(expression: WorkbookFormulaExpression): readonly WorkbookRef[] {
  return formulaRefArray(dataValue(expression, 'inputs'), 'Formula expression inputs')
}

function formulaExpressionLabels(expression: WorkbookFormulaExpression): readonly WorkbookFormulaLabel[] {
  return formulaLabelArray(dataValue(expression, 'labels'), 'Formula expression labels')
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function formatRangeRef(range: CellRangeRef): string {
  const sheet = quoteSheetName(range.sheetName)
  return range.startAddress === range.endAddress ? `${sheet}!${range.startAddress}` : `${sheet}!${range.startAddress}:${range.endAddress}`
}

function placeholderRefSource(ref: WorkbookRef): string {
  return `__bilig_ref_${ref.id.replaceAll(/[^A-Za-z0-9_]+/g, '_')}`
}

function refSource(ref: WorkbookRef): string {
  if (ref.kind === 'range') {
    return formatRangeRef(ref.range)
  }
  if (ref.kind === 'name') {
    return ref.name
  }
  if (ref.kind === 'column' && ref.rows === undefined && ref.table.name) {
    return `${ref.table.name}[${ref.name}]`
  }
  return placeholderRefSource(ref)
}

function operandSource(operand: WorkbookFormulaOperand): string {
  if (isFormulaExpression(operand)) {
    return expressionSource(operand)
  }
  if (isWorkbookRef(operand)) {
    return refSource(operand)
  }
  if (typeof operand === 'number') {
    if (!Number.isFinite(operand)) {
      throw new Error('Formula numbers must be finite')
    }
    return String(operand)
  }
  if (typeof operand === 'boolean') {
    return operand ? 'TRUE()' : 'FALSE()'
  }
  throw new Error(FORMULA_OPERAND_ERROR)
}

function operandInputs(operand: WorkbookFormulaOperand): readonly WorkbookRef[] {
  if (isFormulaExpression(operand)) {
    return formulaExpressionInputs(operand)
  }
  if (isWorkbookRef(operand)) {
    return Object.freeze([operand])
  }
  return EMPTY_FORMULA_INPUTS
}

function operandLabels(operand: WorkbookFormulaOperand): readonly WorkbookFormulaLabel[] {
  if (isFormulaExpression(operand)) {
    return formulaExpressionLabels(operand)
  }
  if (isWorkbookRef(operand)) {
    return Object.freeze([labelForRef(operand)])
  }
  return EMPTY_FORMULA_LABELS
}

function collectInputs(args: readonly WorkbookFormulaOperand[]): readonly WorkbookRef[] {
  return uniqueRefs(args.flatMap((arg) => operandInputs(arg)))
}

function collectLabels(args: readonly WorkbookFormulaOperand[]): readonly WorkbookFormulaLabel[] {
  return uniqueLabels(args.flatMap((arg) => operandLabels(arg)))
}

function binary(left: WorkbookFormulaOperand, operator: '+' | '-' | '*' | '/', right: WorkbookFormulaOperand): WorkbookFormulaExpression {
  return createFormulaExpression(
    `(${operandSource(left)})${operator}(${operandSource(right)})`,
    collectInputs([left, right]),
    collectLabels([left, right]),
  )
}

function call(name: string, args: readonly WorkbookFormulaOperand[]): WorkbookFormulaExpression {
  const callee = formulaFunctionName(name)
  const operands = formulaOperands(args, 'Formula arguments')
  return createFormulaExpression(`${callee}(${operands.map(operandSource).join(',')})`, collectInputs(operands), collectLabels(operands))
}

export const formula = Object.freeze({
  raw(source: string, options: WorkbookRawFormulaOptions = {}): WorkbookFormulaExpression {
    const normalized = rawFormulaOptions(options)
    return createFormulaExpression(source, normalized.inputs, normalized.labels ?? labelsForRefs(normalized.inputs))
  },
  source(expression: WorkbookFormulaOperand): string {
    return normalizeFormulaSource(operandSource(expression))
  },
  inputs(expression: WorkbookFormulaOperand): readonly WorkbookRef[] {
    return operandInputs(expression)
  },
  labels(expression: WorkbookFormulaOperand): readonly WorkbookFormulaLabel[] {
    return operandLabels(expression)
  },
  ref(ref: WorkbookRef): WorkbookFormulaExpression {
    return createFormulaExpression(refSource(ref), [ref], [labelForRef(ref)])
  },
  text(value: string): WorkbookFormulaExpression {
    return createFormulaExpression(`"${formulaText(value, 'Formula text').replaceAll('"', '""')}"`)
  },
  call,
  add(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '+', right)
  },
  subtract(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '-', right)
  },
  multiply(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '*', right)
  },
  divide(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '/', right)
  },
  sum(...args: readonly WorkbookFormulaOperand[]): WorkbookFormulaExpression {
    return call('SUM', args)
  },
} as const)
