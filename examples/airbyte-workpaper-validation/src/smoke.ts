import { readFile } from 'node:fs/promises'

import { readAirbyteMessagesFromJsonl, validateAirbyteOrdersWithWorkPaper } from './airbyte-workpaper-validation.js'

const fixture = await readFile(new URL('../fixtures/orders-airbyte-messages.jsonl', import.meta.url), 'utf8')
const messages = readAirbyteMessagesFromJsonl(fixture)
const result = validateAirbyteOrdersWithWorkPaper({
  initialStateCursor: '2026-05-27T10:05:00Z',
  expectedPaidAmount: 301.75,
  expectedRecordCount: 4,
  messages,
})

console.log(JSON.stringify(result, null, 2))

if (!isExpectedProof(result)) {
  throw new Error(`Unexpected Airbyte WorkPaper validation proof: ${JSON.stringify(result)}`)
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
    Reflect.get(patch, 'committed_state_cursor') === '2026-05-27T10:10:00Z' &&
    Reflect.get(patch, 'record_count') === 4 &&
    Reflect.get(patch, 'gross_amount') === 315 &&
    Reflect.get(patch, 'paid_amount') === 301.75 &&
    Reflect.get(patch, 'rejected_records') === 1 &&
    Reflect.get(patch, 'validation_passed') === true &&
    readNestedBoolean(proof, 'before', 'stateCursorMatchesRecords') === false &&
    readNestedBoolean(proof, 'after', 'stateCursorMatchesRecords') === true &&
    readNestedBoolean(proof, 'afterRestore', 'stateCursorMatchesRecords') === true &&
    Reflect.get(proof, 'verified') === true
  )
}

function readNestedBoolean(value: object, property: string, nestedProperty: string): unknown {
  const nested = Reflect.get(value, property)

  if (typeof nested !== 'object' || nested === null) {
    return undefined
  }

  return Reflect.get(nested, nestedProperty)
}
