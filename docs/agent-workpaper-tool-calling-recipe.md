# WorkPaper Tool-Calling Recipe For Agents

This recipe shows how to wrap `@bilig/headless` WorkPaper operations as
agent-callable functions without binding the workflow to one agent SDK.

Use this pattern when an agent needs to inspect, edit, verify, and persist a
formula-backed workbook from Node. Do not screen scrape a spreadsheet UI when
the WorkPaper API is available. Screenshots are useful for final human review,
but they hide formulas, typed addresses, recalculation state, and persistence
contracts.

Start with the package README for the public API contract:
[`packages/headless/README.md`](../packages/headless/README.md).

For a runnable external example, use
[`examples/headless-workpaper`](../examples/headless-workpaper) and run
`npm run agent:tool-call`. For a smaller writeback-only proof, run
`npm run agent:verify`.

## Tool Contract

Expose a small, boring tool surface first:

- `readSummary(range)` returns computed values and serialized inputs for a
  summary range.
- `setInputCell(sheetName, address, value)` validates the target sheet and A1
  address, writes one value, and returns before/after computed verification.
- `serializeWorkbook()` exports a persisted WorkPaper document only after the
  edit succeeds.

Keep each tool deterministic. Let the agent choose the next action, but make the
tool result carry enough evidence for verification.

## Complete Node Example

```js
import {
  WorkPaper,
  exportWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Assumptions: [
    ['Metric', 'Value'],
    ['Growth rate', 0.1],
  ],
  Revenue: [
    ['Segment', 'Customers', 'ARPA', 'MRR'],
    ['Self serve', 200, 30, '=B2*C2'],
    ['Sales', 15, 300, '=B3*C3'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Current MRR', '=SUM(Revenue!D2:D3)'],
    ['Next month MRR', '=B2*(1+Assumptions!B2)'],
  ],
})

const summarySheet = requireSheet('Summary')
const currentMrrAddress = requireCellAddress('Summary', 'B2')
const nextMonthMrrAddress = requireCellAddress('Summary', 'B3')

const tools = {
  readSummary(range = 'Summary!A1:B3') {
    const parsedRange = workbook.simpleCellRangeFromString(range, summarySheet)
    if (parsedRange === undefined) {
      throw new Error(`invalid summary range: ${range}`)
    }

    return {
      range,
      values: workbook.getRangeValues(parsedRange),
      serialized: workbook.getRangeSerialized(parsedRange),
    }
  },

  setInputCell({ sheetName, address, value }) {
    const target = requireCellAddress(sheetName, address)
    const before = readComputedSummary()

    workbook.setCellContents(target, value)

    const after = readComputedSummary()
    const serializedWorkbook = serializeWorkbook()

    return {
      editedCell: workbook.simpleCellAddressToString(target, {
        includeSheetName: true,
      }),
      before,
      after,
      checks: {
        currentMrrChanged: before.currentMrr !== after.currentMrr,
        nextMonthMrrChanged: before.nextMonthMrr !== after.nextMonthMrr,
        serializedBytes: Buffer.byteLength(serializedWorkbook, 'utf8'),
      },
    }
  },

  serializeWorkbook,
}

console.log(tools.readSummary())
console.log(
  tools.setInputCell({
    sheetName: 'Revenue',
    address: 'B3',
    value: 25,
  }),
)

function requireSheet(sheetName) {
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`unknown sheet: ${sheetName}`)
  }
  return sheetId
}

function requireCellAddress(sheetName, a1Address) {
  const sheetId = requireSheet(sheetName)
  const parsed = workbook.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined) {
    throw new Error(`invalid cell address: ${sheetName}!${a1Address}`)
  }

  if (parsed.sheet !== sheetId) {
    throw new Error(`address ${a1Address} does not belong to ${sheetName}`)
  }

  return parsed
}

function readComputedSummary() {
  return {
    currentMrr: readNumber(currentMrrAddress, 'Current MRR'),
    nextMonthMrr: readNumber(nextMonthMrrAddress, 'Next month MRR'),
  }
}

function readNumber(address, label) {
  const value = workbook.getCellValue(address)
  if (typeof value !== 'object' || value === null || typeof value.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(value)}`)
  }
  return Math.round(value.value * 100) / 100
}

function serializeWorkbook() {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workbook, {
      includeConfig: true,
    }),
  )
}
```

The important check is not that the write call returned. It is that the computed
summary changed as expected:

```json
{
  "editedCell": "Revenue!B3",
  "before": {
    "currentMrr": 10500,
    "nextMonthMrr": 11550
  },
  "after": {
    "currentMrr": 13500,
    "nextMonthMrr": 14850
  },
  "checks": {
    "currentMrrChanged": true,
    "nextMonthMrrChanged": true,
    "serializedBytes": 1155
  }
}
```

`serializedBytes` will vary as the document schema evolves. Treat it as a
positive persistence check, not a stable snapshot value.

## Agent Guardrails

- Validate sheet names with `getSheetId()` before parsing a target address.
- Parse user-facing addresses through `simpleCellAddressFromString()` or
  `simpleCellRangeFromString()` instead of building `{ row, col }` objects from
  ad hoc string splits.
- Return computed values after every write; do not ask the agent to infer
  success from a rendered grid.
- Serialize only after a successful write and verification readback.
- Keep tool results small. Return the range, changed cell, before/after values,
  and persistence check; do not dump the whole workbook unless the agent asks
  for it.
- Use public `@bilig/headless` exports and WorkPaper methods only. Do not import
  from internal `src/`, `dist/`, or monorepo package internals in an external
  agent workflow.

## When To Add More Tools

Add tools only after the agent has a repeated need for them:

- `readRange(range)` for broader model inspection
- `setFormula(sheetName, address, formula)` when formulas are first-class agent
  outputs
- `validateFormula(address)` when the workflow needs structured diagnostics
- `persistAndRestore()` when the workflow must prove round-trip safety before
  committing output

The same rule holds: every mutating tool should return computed verification
and enough context for the caller to explain what changed.
