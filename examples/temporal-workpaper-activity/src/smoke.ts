import { MockActivityEnvironment } from '@temporalio/testing'
import { calculateWorkPaperQuoteActivity } from './activities'
import type { TemporalWorkPaperQuoteInput } from './types'

const cliInput = readCliInput(process.argv.slice(2))
const activityEnvironment = new MockActivityEnvironment()
const result = await activityEnvironment.run(calculateWorkPaperQuoteActivity, cliInput)

console.log(JSON.stringify(result, null, 2))

if (!isExpectedProof(result)) {
  throw new Error(`Unexpected Temporal WorkPaper proof: ${JSON.stringify(result)}`)
}

function readCliInput(args: readonly string[]): TemporalWorkPaperQuoteInput {
  const values = new Map<string, string>()

  for (let index = 0; index < args.length; index += 1) {
    const key = args[index]

    if (!key?.startsWith('--')) {
      continue
    }

    const next = args[index + 1]
    if (next === undefined) {
      throw new Error(`Missing value for ${key}`)
    }

    values.set(key.slice(2), next)
    index += 1
  }

  return {
    previousQuantity: readNumber(values.get('previous-quantity') ?? '12', 'previous quantity'),
    quantity: readNumber(values.get('quantity') ?? '18', 'quantity'),
    unitPrice: readNumber(values.get('unit-price') ?? '125', 'unit price'),
    discountRate: readNumber(values.get('discount-rate') ?? '0.1', 'discount rate'),
    taxRate: readNumber(values.get('tax-rate') ?? '0.08', 'tax rate'),
    unitCost: readNumber(values.get('unit-cost') ?? '52', 'unit cost'),
    output: values.get('output') ?? 'workpaper-proof.json',
  }
}

function readNumber(value: string, label: string): number {
  const numeric = Number(value)

  if (!Number.isFinite(numeric)) {
    throw new Error(`${label} must be a finite number`)
  }

  return numeric
}

function isExpectedProof(value: unknown): boolean {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }

  const patch = Reflect.get(value, 'patch')
  const proof = Reflect.get(value, 'proof')
  const temporalBoundary = Reflect.get(value, 'temporalBoundary')

  if (typeof patch !== 'object' || patch === null || typeof proof !== 'object' || proof === null) {
    return false
  }

  if (typeof temporalBoundary !== 'object' || temporalBoundary === null) {
    return false
  }

  return (
    Reflect.get(patch, 'total') === 2187 &&
    readNestedNumber(proof, 'before', 'total') === 1458 &&
    readNestedNumber(proof, 'after', 'total') === 2187 &&
    readNestedNumber(proof, 'afterRestore', 'total') === 2187 &&
    Reflect.get(proof, 'verified') === true &&
    Reflect.get(temporalBoundary, 'workflowImportsWorkPaper') === false &&
    Reflect.get(temporalBoundary, 'activityOwnsWorkPaper') === true
  )
}

function readNestedNumber(value: object, property: string, nestedProperty: string): unknown {
  const nested = Reflect.get(value, property)

  if (typeof nested !== 'object' || nested === null) {
    return undefined
  }

  return Reflect.get(nested, nestedProperty)
}
