# @bilig/workpaper

Bilig WorkPaper is an API, CLI evaluator, and optional MCP server for
workbook-shaped business logic in Node.js.

Use this when business logic is easiest to review as workbook cells and
formulas, but the calculation needs to run in a backend service, queue worker,
serverless route, test, or tool.

`@bilig/workpaper` is the canonical scoped npm entrypoint. The unscoped
`bilig-workpaper` package remains published as a compatibility and search alias.

## Install

```sh
npm install @bilig/workpaper
```

## Start Here

Pick the door that matches the state you own:

| Door                    | Run first                                                                                            | What it proves                                                                                     |
| ----------------------- | ---------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| Node service or test    | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json` | edit input, recalculate output, persist JSON, restore, and return `verified: true`.                |
| New project             | `npm create @bilig/workpaper@latest pricing-workpaper`                                               | a starter wired to the same WorkPaper proof loop.                                                  |
| Tool host or MCP client | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json`         | tool discovery, cell mutation, formula readback, JSON export, restart proof, and `verified: true`. |
| Unsure which proof fits | `npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json`                       | compact routing card with proof commands, evidence fields, and public links.                       |

`bilig-agent-start --json` is intentionally small. It prints first proof
commands, required evidence fields, expected MCP tools, and public discovery
links without asking a tool host to read the whole site.

## What Success Looks Like

Run the service proof without cloning the repo:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

The useful output is not a write-call status. It is readback proof:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "workpaper-service",
  "verified": true,
  "packageVersions": {
    "@bilig/workpaper": "0.163.0"
  },
  "evidence": {
    "editedCell": "Inputs!B2",
    "dependentCell": "Summary!B2",
    "before": 24000,
    "after": 38400,
    "afterRestore": 38400,
    "persistedDocumentBytes": 999
  }
}
```

If this is close to production, watch releases and review the public limits:
<https://github.com/proompteng/bilig/subscription>.

For a richer tool check, add `--scenario revenue-plan` to the `agent-mcp`
evaluator. It proves `SUM`, `SUMIF`, `XLOOKUP`, `FILTER`, a named expression,
JSON persistence, and restart readback.

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

If the workbook has provider-backed formulas such as `IMPORTRANGE`, run
`npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json`.
That proves the formula fails closed with an adapter diagnostic, then verifies a
local synthetic adapter readback. It does not call Google Sheets.

If one of those matches your workflow, keep Bilig nearby:
<https://github.com/proompteng/bilig>.

Framework examples live in the repo instead of this first screen. Start with
the closest runtime:

- Tool runtimes: Vercel AI SDK, LangGraph, LangChain MCP adapters, Open WebUI,
  FastMCP, Agno, Pydantic AI, smolagents.
- Workflow engines: Trigger.dev, Inngest, Temporal, Airflow, Dagster, Kestra,
  Prefect, Windmill.
- Low-code and data tools: n8n, Pipedream, Directus, Airbyte, Meltano.
- Existing `.xlsx` files: use `@bilig/xlsx-formula-recalc` for saved-file
  recalculation and exported workbook readback.

## Searchable Example Guides

These are integration guide names that users search for on npm. They are links,
not the first-run path:

| Guide need                                              | Start here                                                                     |
| ------------------------------------------------------- | ------------------------------------------------------------------------------ |
| Windmill TypeScript workflow fields                     | <https://proompteng.github.io/bilig/windmill-workpaper-script.html>            |
| Trigger.dev durable task fields                         | <https://proompteng.github.io/bilig/triggerdev-workpaper-task.html>            |
| Inngest durable step fields                             | <https://proompteng.github.io/bilig/inngest-workpaper-step.html>               |
| Temporal TypeScript Activity decisions                  | <https://proompteng.github.io/bilig/temporal-workpaper-activity.html>          |
| Apache Airflow DAG task outputs                         | <https://proompteng.github.io/bilig/airflow-workpaper-dag.html>                |
| Dagster asset materialization metadata                  | <https://proompteng.github.io/bilig/dagster-workpaper-asset.html>              |
| Kestra Node Commands flow fields                        | <https://proompteng.github.io/bilig/kestra-workpaper-flow.html>                |
| Prefect flow fields                                     | <https://proompteng.github.io/bilig/prefect-workpaper-flow.html>               |
| Directus Flow operation for persisted calculated fields | <https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html>    |
| n8n formula readback for self-hosted workflows          | <https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html>       |
| Dify formula readback                                   | <https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html>      |
| Flowise formula readback                                | <https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html>   |
| Pipedream formula readback                              | <https://proompteng.github.io/bilig/pipedream-workpaper-formula-readback.html> |
| FastMCP Python client for hosted and local stdio checks | <https://proompteng.github.io/bilig/fastmcp-workpaper-client.html>             |
| Agno MCPTools workbook readback                         | <https://proompteng.github.io/bilig/agno-workpaper-mcp.html>                   |
| Pydantic AI MCPToolset typed readback                   | <https://proompteng.github.io/bilig/pydantic-ai-workpaper-mcp.html>            |
| Google ADK McpToolset workbook readback                 | <https://proompteng.github.io/bilig/google-adk-workpaper-mcp.html>             |
| Hugging Face smolagents tool                            | <https://proompteng.github.io/bilig/smolagents-workpaper-tool.html>            |
| Hugging Face Gradio MCP Space                           | <https://proompteng.github.io/bilig/huggingface-workpaper-space.html>          |

## Use A WorkPaper In Node

```ts
import { buildA1WorkPaper } from '@bilig/workpaper'

const book = buildA1WorkPaper({
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

const proof = book.editAndReadback('Inputs!B2', 48, {
  readbackRange: 'Summary!B2',
})

console.log({
  editedCell: proof.editedCell,
  before: proof.beforeReadback.displayValues,
  after: proof.afterReadback.displayValues,
  afterRestore: proof.restoredReadback.displayValues,
  persistedDocumentBytes: proof.persistedDocumentBytes,
  verified: proof.verified,
})

book.dispose()
```

Use `book.set('Inputs!B2', 48)`, `book.setMany({ 'Inputs!B3': 1500 })`,
`book.readMany(['Inputs!B2', 'Summary!B2'])`, `book.display('Summary!B2')`,
and `book.saveJson()` when you do not need the full proof object. Use
`book.editManyAndReadback()` when several inputs should commit as one atomic
proof with typed readback comparison, formula diagnostics, persistence, and
restore checks.

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
npm exec --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-n8n-formula-server --port 4321
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

`bilig-evaluate` prints a `bilig-evaluator.v1` object with `door`, `evidence`,
`verified`, `limitations`, and the source command output.

The challenge commands edit one input, recalculate dependent formulas, export
WorkPaper JSON, restore it, and print a `verified: true` proof object.
Use `--from-xlsx` when the tool host already has an XLSX file: Bilig imports it once
into an in-memory MCP server by default, or into persisted WorkPaper JSON when
`--workpaper --writable` is also supplied. That XLSX-backed MCP path also lists
`analyze_workbook_risk`, a read-only tool fixed to the source workbook passed at
startup. It reports workbook risk indicators before a workflow trusts the imported
WorkPaper and does not certify Excel compatibility.

## Tool Host WorkPaper Handoff

When a tool host is about to solve a spreadsheet task by opening Excel,
LibreOffice, Google Sheets, or a screenshot grid, hand it the WorkPaper checklist
instead:

```sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The kit gives the host one path: install the instructions, run the no-key MCP
evaluator, paste a workbook edit task, and require computed readback plus
persisted state before reporting success. Use `bilig-mcp-challenge --json` only
when debugging the lower-level MCP transcript.

Docs: <https://proompteng.github.io/bilig/agent-adoption-kit.html>

## Workflow Builders

Use the local formula-readback server when a workflow platform should
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

## Tool Commands And Optional MCP

The npm tarball exposes the same CLI entrypoints through the canonical scoped package, so tool
hosts can install one focused package and still get the MCP stdio server:

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
service-owned calculations, JSON persistence, XLSX import/export, and verified
readback. Unsupported Excel functions, external workbook links,
macros, and volatile functions may need review.

## After The Proof

If the starter or challenge output matches your service or tool workflow,
keep the repository nearby for release notes and public limits:
<https://github.com/proompteng/bilig>.

Watch releases if this is close to a production path:
<https://github.com/proompteng/bilig/subscription>.

If the model is close but blocked by a formula, import/export, persistence,
framework, MCP, or package-boundary gap, open the smallest implementation gap:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

Full docs: <https://proompteng.github.io/bilig/>
