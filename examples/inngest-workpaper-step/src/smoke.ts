import { calculateWorkPaperQuote } from './workpaper-quote.js'

const result = calculateWorkPaperQuote({
  previousQuantity: 12,
  quantity: 18,
  unitPrice: 125,
  discountRate: 0.1,
  taxRate: 0.08,
  unitCost: 52,
})

console.log(JSON.stringify(result, null, 2))

if (!isExpectedProof(result)) {
  throw new Error(`Unexpected Inngest WorkPaper proof: ${JSON.stringify(result)}`)
}

function isExpectedProof(value: unknown): boolean {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }

  const patch = Reflect.get(value, 'patch')
  const proof = Reflect.get(value, 'proof')

  if (typeof patch !== 'object' || patch === null || typeof proof !== 'object' || proof === null) {
    return false
  }

  return (
    Reflect.get(patch, 'total') === 2187 &&
    readNestedNumber(proof, 'before', 'total') === 1458 &&
    readNestedNumber(proof, 'after', 'total') === 2187 &&
    readNestedNumber(proof, 'afterRestore', 'total') === 2187 &&
    Reflect.get(proof, 'verified') === true
  )
}

function readNestedNumber(value: object, property: string, nestedProperty: string): unknown {
  const nested = Reflect.get(value, property)

  if (typeof nested !== 'object' || nested === null) {
    return undefined
  }

  return Reflect.get(nested, nestedProperty)
}
