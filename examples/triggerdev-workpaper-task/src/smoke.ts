import { calculateWorkPaperQuote } from './workpaper-quote.js'

const result = calculateWorkPaperQuote({
  previousQuantity: 12,
  quantity: 18,
  unitPrice: 125,
  discountRate: 0.1,
  taxRate: 0.08,
  unitCost: 52,
})

const expectedPatch = {
  subtotal: 2250,
  discount_amount: 225,
  taxable_amount: 2025,
  tax_amount: 162,
  total: 2187,
  margin_amount: 1089,
}

if (JSON.stringify(result.patch) !== JSON.stringify(expectedPatch)) {
  throw new Error(`unexpected patch: ${JSON.stringify(result.patch)}`)
}

if (
  result.proof.editedCell !== 'Inputs!B2' ||
  result.proof.before.total !== 1458 ||
  result.proof.after.total !== 2187 ||
  result.proof.afterRestore.total !== 2187 ||
  result.proof.persistedDocumentBytes <= 0 ||
  !result.proof.verified
) {
  throw new Error(`unexpected proof: ${JSON.stringify(result.proof)}`)
}

console.log(JSON.stringify(result, null, 2))
