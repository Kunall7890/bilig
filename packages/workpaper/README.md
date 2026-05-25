# @bilig/workpaper

Scoped Bilig WorkPaper runtime for Node.js services, agent tools, and server-side spreadsheet formulas.

Use this when business logic is easiest to review as workbook cells and
formulas, but the calculation needs to run in a backend service, queue worker,
serverless route, test, or coding-agent tool.

`@bilig/workpaper` is the canonical scoped npm entrypoint. The unscoped
`bilig-workpaper` package remains published as a compatibility and search alias.

## Install

```sh
npm install @bilig/workpaper
```

## Start Here

Pick the path that matches the workflow you are trying to unblock:

| You need...                                                              | Run this first                                                                                                | Proof you should get                                                                                                |
| ------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| Formula workbook logic inside a Node service, route, queue, or test      | `npm create @bilig/workpaper@latest pricing-workpaper`                                                        | Inputs are written, formulas recalculate, JSON persists, restore matches readback, and `verified: true` is printed. |
| A coding agent or MCP client that needs spreadsheet operations           | `npm create @bilig/workpaper@latest pricing-agent -- --agent`                                                 | The generated project includes an agent contract, MCP config, and `npm run agent:verify`.                           |
| Windmill TypeScript workflow fields                                      | `cd examples/windmill-workpaper-script && pnpm install --ignore-workspace --lockfile=false && pnpm run smoke` | The script returns a calculated field patch plus before/after/restore WorkPaper proof with `verified: true`.        |
| Trigger.dev durable task fields                                          | `cd examples/triggerdev-workpaper-task && pnpm install --ignore-workspace --lockfile=false && pnpm run smoke` | The task helper returns a calculated field patch plus before/after/restore WorkPaper proof with `verified: true`.   |
| Directus Flow operation for persisted calculated fields                  | `cd examples/directus-workpaper-flow-operation && npm install && npm run smoke`                               | The operation returns a Directus `patch` plus before/after/restore WorkPaper proof with `verified: true`.           |
| n8n, Dify, or Flowise formula readback without spreadsheet UI automation | `npm exec --package @bilig/workpaper@latest -- bilig-n8n-formula-server --port 4321`                          | The workflow writes one input cell, reads dependent formula output, and returns a compact JSON proof.               |
| Open WebUI needs MCP spreadsheet tools                                   | `npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json`                                    | Open WebUI can call the hosted Streamable HTTP endpoint or a local stdio server bridged through `mcpo`.             |
| An existing `.xlsx` file with stale formula results after Node edits     | `npx --package @bilig/xlsx-formula-recalc xlsx-recalc --demo --json`                                          | The file-level path updates inputs and returns fresh formula values without Excel, LibreOffice, or a browser.       |

## Use A WorkPaper In Node

```ts
import { WorkPaper } from '@bilig/workpaper'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Units', 40],
    ['Price', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

function cell(address: string) {
  const parsed = workbook.simpleCellAddressFromString(address)

  if (parsed === undefined) {
    throw new Error(`Unknown cell: ${address}`)
  }

  return parsed
}

function setCell(address: string, value: string | number | boolean | null) {
  workbook.setCellContents(cell(address), value)
}

function displayAt(address: string) {
  return workbook.getCellDisplayValue(cell(address))
}

const before = displayAt('Summary!B2')

setCell('Inputs!B2', 48)
setCell('Inputs!B3', 1500)

const after = displayAt('Summary!B2')
const document = workbook.exportSnapshot()

console.log({
  editedCells: ['Inputs!B2', 'Inputs!B3'],
  readCell: 'Summary!B2',
  before,
  after,
  persistedDocumentBytes: JSON.stringify(document).length,
  verified: after === '72000',
})

workbook.dispose()
```

## Prove The Agent Loop Without Cloning

The package ships proof commands for coding agents and service evaluators:

```sh
npm exec --package @bilig/workpaper -- bilig-agent-challenge
npm exec --package @bilig/workpaper -- bilig-mcp-challenge
npm exec --package @bilig/workpaper -- bilig-n8n-formula-server --port 4321
npm exec --package @bilig/workpaper -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

The challenge commands edit one input, recalculate dependent formulas, export
WorkPaper JSON, restore it, and print a `verified: true` proof object.

## Agent Workflow Builders

Use the local formula-readback server when an agent workflow platform should
orchestrate the task but Bilig should own workbook state:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-n8n-formula-server --port 4321
```

Start from the repo examples:

- `examples/directus-workpaper-flow-operation`
- `examples/windmill-workpaper-script`
- `examples/triggerdev-workpaper-task`
- `examples/n8n-workpaper-formula-readback/bilig-workpaper-formula-readback.n8n.json`
- `examples/dify-workpaper-formula-readback`
- `examples/flowise-workpaper-formula-readback/bilig-workpaper-formula-readback.flowise-tool.json`

Docs:

- <https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html>
- <https://proompteng.github.io/bilig/windmill-workpaper-script.html>
- <https://proompteng.github.io/bilig/triggerdev-workpaper-task.html>
- <https://proompteng.github.io/bilig/open-webui-workpaper-mcp.html>
- <https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html>
- <https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html>
- <https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html>

## XLSX Import And Export

```ts
import { WorkPaper } from '@bilig/workpaper'
import { exportXlsx, importXlsx } from '@bilig/workpaper/xlsx'
```

Use `@bilig/xlsx-formula-recalc` when you only need to edit and recalculate
XLSX files. Use `@bilig/exceljs-formula-recalc` when you already use ExcelJS
and need recalculated formula results after changing inputs.

## Agent Commands And Optional MCP

The npm tarball exposes the same CLI entrypoints as `@bilig/headless`, so agents
can install one focused package and still get the MCP stdio server:

```ts
import { createWorkPaperMcpServer } from '@bilig/workpaper/mcp'
```

For a runnable starter project with `AGENTS.md`, MCP client config, and an
`agent:verify` script:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
```

## Scope

Bilig is not a desktop Excel clone. It is a formula workbook runtime for
service-owned calculations, JSON persistence, XLSX import/export, and
agent-readable readback. Unsupported Excel functions, external workbook links,
macros, and volatile functions may need review.

Full docs: <https://proompteng.github.io/bilig/>
