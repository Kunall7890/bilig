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

Pick the door that matches the state you own:

| Door                       | Run first                                                                                            | What it proves                                                                                     |
| -------------------------- | ---------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| Node service or test       | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json` | edit input, recalculate output, persist JSON, restore, and return `verified: true`.                |
| Coding agent or MCP client | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json`         | tool discovery, cell mutation, formula readback, JSON export, restart proof, and `verified: true`. |
| New project                | `npm create @bilig/workpaper@latest pricing-agent -- --agent`                                        | a starter with `AGENTS.md`, MCP config, and `npm run agent:verify` wired to the same proof loop.   |

If one of those matches your workflow, keep Bilig nearby:
<https://github.com/proompteng/bilig>. Watch releases if this is close to a
production dependency: <https://github.com/proompteng/bilig/subscription>.

Framework examples live in the repo instead of this first screen. Start with
the closest runtime:

- Agent tools: Vercel AI SDK, LangGraph, LangChain MCP adapters, Open WebUI,
  FastMCP, smolagents.
- Workflow engines: Trigger.dev, Inngest, Temporal, Airflow, Dagster, Kestra,
  Prefect, Windmill.
- Low-code and data tools: n8n, Pipedream, Directus, Airbyte, Meltano.
- Existing `.xlsx` files: use `@bilig/xlsx-formula-recalc` for stale formula
  caches and exported workbook readback.

## Searchable Example Guides

These are integration guide names that users search for on npm. They are links,
not the first-run path:

| Guide need                                                                          | Start here                                                                   |
| ----------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Windmill TypeScript workflow fields                                                 | <https://proompteng.github.io/bilig/windmill-workpaper-script.html>          |
| Trigger.dev durable task fields                                                     | <https://proompteng.github.io/bilig/triggerdev-workpaper-task.html>          |
| Inngest durable step fields                                                         | <https://proompteng.github.io/bilig/inngest-workpaper-step.html>             |
| Temporal TypeScript Activity decisions                                              | <https://proompteng.github.io/bilig/temporal-workpaper-activity.html>        |
| Apache Airflow DAG task outputs                                                     | <https://proompteng.github.io/bilig/airflow-workpaper-dag.html>              |
| Dagster asset materialization metadata                                              | <https://proompteng.github.io/bilig/dagster-workpaper-asset.html>            |
| Kestra Node Commands flow fields                                                    | <https://proompteng.github.io/bilig/kestra-workpaper-flow.html>              |
| Prefect flow fields                                                                 | <https://proompteng.github.io/bilig/prefect-workpaper-flow.html>             |
| Directus Flow operation for persisted calculated fields                             | <https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html>  |
| n8n formula readback in Cloud or self-hosted workflows                             | <https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html>     |
| Dify formula readback                                                              | <https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html>    |
| Flowise formula readback                                                            | <https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html> |
| Pipedream formula readback                                                          | <https://proompteng.github.io/bilig/pipedream-workpaper-formula-readback.html> |
| FastMCP Python client for hosted MCP smoke tests                                    | <https://proompteng.github.io/bilig/fastmcp-workpaper-client.html>           |
| Hugging Face smolagents tool                                                        | <https://proompteng.github.io/bilig/smolagents-workpaper-tool.html>          |

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

## Use WorkPaper Tools With The Vercel AI SDK

Install the AI SDK and Zod in the application that owns the agent loop:

```sh
npm install @bilig/workpaper ai zod
```

Then expose a WorkPaper as normal AI SDK tools:

```ts
import { generateText, stepCountIs } from 'ai'
import { WorkPaper } from '@bilig/workpaper'
import { createAiSdkWorkPaperTools } from '@bilig/workpaper/ai-sdk'

const workpaper = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Qualified opportunities', 20],
    ['Win rate', 0.25],
    ['Average ARR', 12000],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Expected customers', '=Inputs!B2*Inputs!B3'],
    ['Expected ARR', '=B2*Inputs!B4'],
  ],
})

const tools = createAiSdkWorkPaperTools({
  workpaper,
  defaultReadRange: 'Summary!A1:B3',
  proofRange: 'Summary!A1:B3',
  writableSheets: ['Inputs'],
})

const result = await generateText({
  model,
  tools,
  stopWhen: stepCountIs(2),
  prompt: 'Read the summary, set Inputs!B3 to 0.4, then report the computed ARR change.',
})

console.log(result.text)
```

The mutating tool returns `editedCell`, `before`, `after`, `restored`, and
`checks`. Keep `writableSheets` narrow so the model can edit inputs without
rewriting formula sheets.

## Verify Without Cloning

The package ships evaluator and direct proof commands:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
npm exec --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-n8n-formula-server --port 4321
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

`bilig-evaluate` prints a `bilig-evaluator.v1` object with `door`, `evidence`,
`verified`, `limitations`, and the source command output.

The challenge commands edit one input, recalculate dependent formulas, export
WorkPaper JSON, restore it, and print a `verified: true` proof object.

## Agent Adoption Kit

When another coding agent is about to solve a spreadsheet task by opening Excel,
LibreOffice, Google Sheets, or a screenshot grid, hand it the adoption kit
instead:

```sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

The kit gives the agent one path: install the instructions, run the no-key MCP
proof, paste a workbook edit task, and require computed readback plus persisted
state before reporting success.

Docs: <https://proompteng.github.io/bilig/agent-adoption-kit.html>

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
- `examples/airbyte-workpaper-validation`
- `examples/meltano-workpaper-utility`
- `examples/temporal-workpaper-activity`
- `examples/airflow-workpaper-dag`
- `examples/dagster-workpaper-asset`
- `examples/kestra-workpaper-flow`
- `examples/prefect-workpaper-flow`
- `examples/n8n-workpaper-formula-readback/bilig-workpaper-formula-readback.n8n.json`
- `examples/dify-workpaper-formula-readback` mirrors the Dify plugin source;
  the package was merged upstream in
  <https://github.com/langgenius/dify-plugins/pull/2451>
- `examples/flowise-workpaper-formula-readback/bilig-workpaper-formula-readback.flowise-tool.json`
- `integrations/pipedream-bilig-workpaper` mirrors the Pipedream action shape;
  the public review is <https://github.com/PipedreamHQ/pipedream/pull/20972>
- `examples/fastmcp-workpaper-client`
- `examples/langchain-mcp-workpaper-toolnode`
- `examples/smolagents-workpaper-tool`

Docs:

- <https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html>
- <https://proompteng.github.io/bilig/windmill-workpaper-script.html>
- <https://proompteng.github.io/bilig/triggerdev-workpaper-task.html>
- <https://proompteng.github.io/bilig/airbyte-workpaper-validation.html>
- <https://proompteng.github.io/bilig/meltano-workpaper-utility.html>
- <https://proompteng.github.io/bilig/temporal-workpaper-activity.html>
- <https://proompteng.github.io/bilig/airflow-workpaper-dag.html>
- <https://proompteng.github.io/bilig/dagster-workpaper-asset.html>
- <https://proompteng.github.io/bilig/kestra-workpaper-flow.html>
- <https://proompteng.github.io/bilig/prefect-workpaper-flow.html>
- <https://proompteng.github.io/bilig/open-webui-workpaper-mcp.html>
- <https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html>
- <https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html>
- <https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html>
- <https://proompteng.github.io/bilig/fastmcp-workpaper-client.html>
- <https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html>
- <https://proompteng.github.io/bilig/smolagents-workpaper-tool.html>

## XLSX Import And Export

```ts
import { WorkPaper } from '@bilig/workpaper'
import { exportXlsx, importXlsx } from '@bilig/workpaper/xlsx'
```

Use `@bilig/xlsx-formula-recalc` when you only need to edit and recalculate
XLSX files. Use `@bilig/exceljs-formula-recalc` when you already use ExcelJS
and need recalculated formula results after changing inputs.

## Agent Commands And Optional MCP

The npm tarball exposes the same CLI entrypoints through the canonical scoped package, so agents
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

## After The Proof

If the starter or challenge output gives you `verified: true` for the service
or agent workflow you need, star or bookmark Bilig so the WorkPaper runtime is
easy to find again: <https://github.com/proompteng/bilig/stargazers>.

Watch releases if this is close to a production path:
<https://github.com/proompteng/bilig/subscription>.

If the model is close but blocked by a formula, import/export, persistence,
framework, MCP, or package-boundary gap, open the smallest adoption blocker:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

Full docs: <https://proompteng.github.io/bilig/>
