import assert from 'node:assert/strict'
import { runWorkbookAgentModelExample } from '../examples/workbook-agent-model/agent-model.ts'

const output = await runWorkbookAgentModelExample()

assert.equal(output.model.name, 'consumer-table-calculation')
assert.deepEqual(output.model.actions, ['calculate'])
assert.equal(output.planning.verification.status, 'valid')
assert.match(output.planning.command.commandId, /^cmd_[0-9a-f]{16}$/u)
assert.equal(output.planning.command.baseRevision, 'example-rev-1')
assert.equal(output.planning.command.idempotencyKey, 'example-calculate')
assert.equal(output.preview.status, 'previewed')
if (output.preview.status !== 'previewed') {
  throw new Error('expected previewed output')
}
assert.equal(output.preview.preview.materializedOps.length, 2)
assert.deepEqual(output.runtime.verification, {
  status: 'supported',
  missing: [],
})
assert.equal(output.run.status, 'done')
if (output.run.status !== 'done') {
  throw new Error('expected done output')
}
assert.equal(output.run.receipt?.commandId, output.planning.command.commandId)
assert.equal(output.run.receipt?.baseRevision, 'example-rev-1')
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
