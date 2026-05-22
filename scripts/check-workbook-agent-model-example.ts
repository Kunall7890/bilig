import assert from 'node:assert/strict'
import { runWorkbookAgentModelExample } from '../examples/workbook-agent-model/agent-model.ts'

const output = await runWorkbookAgentModelExample()

assert.equal(output.model.name, 'consumer-table-calculation')
assert.deepEqual(output.model.actions, ['calculate'])
assert.equal(output.planning.verification.status, 'valid')
assert.equal(output.run.status, 'done')
assert.deepEqual(output.workbook.formulas, {
  D2: '(Sheet1!B2)*(Sheet1!C2)',
  D3: null,
  D4: '(Sheet1!B4)*(Sheet1!C4)',
})
assert.deepEqual(output.workbook.values, {
  D2: 6,
  D3: null,
  D4: 20,
})

const capabilities = new Set(output.planning.requirements.requirements.map((requirement) => requirement.capability))
assert(capabilities.has('writeFormula'))
assert(capabilities.has('read'))
assert(capabilities.has('verifyCheck'))

const checkStatuses = output.run.checks.map((check) => [check.kind, check.status])
assert.deepEqual(checkStatuses, [
  ['exists', 'passed'],
  ['exists', 'passed'],
  ['noFormulaErrors', 'passed'],
  ['valuesEqual', 'passed'],
])

console.log('workbook agent model example passed')
