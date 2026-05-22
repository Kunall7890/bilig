# Bilig

[![CI](https://github.com/proompteng/bilig/actions/workflows/ci.yml/badge.svg)](https://github.com/proompteng/bilig/actions/workflows/ci.yml)
[![npm: @bilig/workbook](https://img.shields.io/npm/v/@bilig/workbook?label=%40bilig%2Fworkbook)](https://www.npmjs.com/package/@bilig/workbook)
[![npm: @bilig/workpaper](https://img.shields.io/npm/v/@bilig/workpaper?label=%40bilig%2Fworkpaper)](https://www.npmjs.com/package/@bilig/workpaper)
[![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/proompteng/bilig/badge)](https://scorecard.dev/viewer/?uri=github.com/proompteng/bilig)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

Bilig is a TypeScript workbook platform for agent-first workbook models,
formula execution, and XLSX recalculation.

Use it when workbook logic needs to be run by code, verified by tests, or handed
to an AI agent without relying on a human spreadsheet UI. The core idea is
simple: describe workbook intent in a stable API, run it in a runtime, and
return proof.

Project site: <https://proompteng.github.io/bilig/>

## Start Here

| What you need | Start with | Why |
| --- | --- | --- |
| Define a generic workbook model for agents | [`@bilig/workbook`](packages/workbook/README.md) | Consumer-defined `find`, `checks`, and `actions` that plan portable workbook intent. |
| Execute workbook formulas in a Node service | [`@bilig/workpaper`](packages/workpaper/README.md) | Headless workbook state, formula calculation, JSON save and restore. |
| Recalculate XLSX bytes after input edits | [`@bilig/xlsx-formula-recalc`](packages/xlsx-formula-recalc/README.md) | File-level recalculation without opening Excel, LibreOffice, or Google Sheets. |
| Recalculate a SheetJS pipeline | [`@bilig/sheetjs-formula-recalc`](packages/sheetjs-formula-recalc/README.md) | Keep `xlsx` / SheetJS as the file boundary and refresh cached formula values. |
| Recalculate an ExcelJS pipeline | [`@bilig/exceljs-formula-recalc`](packages/exceljs-formula-recalc/README.md) | Keep ExcelJS as the file boundary and refresh formula results there. |
| Build or run the full app | [`apps/bilig`](apps/bilig) and [`apps/web`](apps/web) | Fullstack runtime plus browser workbook shell. |

## Choose An Evaluation Path

If you are evaluating...

| Path | Start here |
| --- | --- |
| Basic package fit | [Why use Bilig?](docs/why-use-bilig.md) |
| Published npm package | [90-second Node quickstart](docs/try-bilig-headless-in-node.md) |
| Backend service shape | [Quote approval WorkPaper API](docs/quote-approval-workpaper-api.md) |
| XLSX workflow | [XLSX formula recalculation example](docs/xlsx-formula-recalculation-node.md) |
| MCP tools | [MCP spreadsheet tool server](docs/mcp-workpaper-tool-server.md) |
| Package trust | [npm provenance](docs/npm-provenance-package-trust.md) |
| Almost a fit | [adoption blocker form](https://github.com/proompteng/bilig/discussions/new?category=general) |
| Real workbook blocked | [submit a workbook fixture](docs/submit-workbook-fixture.md) |

## Agent-First Workbook Models

`@bilig/workbook` is the public package for generic agent workbook contracts.
The bar is simple: an agent must love this library.

```ts
import { defineModel, formula } from "@bilig/workbook";

export const model = defineModel({
  name: "custom-calculation",

  find(workbook) {
    const table = workbook.findTable({
      headers: ["Base", "Rate", "Result"],
    });

    return {
      table,
      base: table.column("Base"),
      rate: table.column("Rate"),
      result: table.column("Result"),
    };
  },

  checks({ refs, workbook }) {
    return [
      workbook.check.exists(refs.table),
      workbook.check.noFormulaErrors(refs.result),
    ];
  },

  actions: {
    calculate({ refs, workbook }) {
      workbook.writeFormula(refs.result, formula.multiply(refs.base, refs.rate));
      workbook.check.noFormulaErrors(refs.result);
    },
  },
});
```

That model is intentionally not a revenue model, prepaid model, reporting model,
or any other built-in business template. Consumers define their own models.
Bilig provides refs, formula helpers, checks, operation types, plan validation,
and runtime handoff.

Agent flow:

1. Inspect with `describeModel(model)`.
2. Plan with `planWorkbookAction(model, actionName, input)`.
3. Verify with `verifyPlan(plan)`.
4. Run through a runtime adapter with `runWorkbookAction`.
5. Accept success only when `WorkbookRunResult.status === "done"` and checks
   come back as proof.

Read the package README for the full API:
[`packages/workbook/README.md`](packages/workbook/README.md).

## Formula Runtime

For services that own workbook state and need calculated values immediately,
use `@bilig/workpaper`.

```sh
mkdir bilig-workpaper-eval
cd bilig-workpaper-eval
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install -D tsx typescript @types/node
curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx quickstart.ts
```

Expected shape:

```json
{
  "verified": true,
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400
}
```

Use `@bilig/workpaper` when a Node service, queue worker, test, or agent tool
needs to write workbook inputs, recalculate formulas, read values, and persist
state without a browser grid.

The quickstart source is maintained at
[`examples/headless-workpaper/npm-eval.ts`](examples/headless-workpaper/npm-eval.ts)
and mirrored to <https://proompteng.github.io/bilig/npm-eval.ts>.

To generate a starter project, use the published
`@bilig/create-workpaper` package:

```sh
npm create @bilig/workpaper@latest pricing-workpaper
npm create @bilig/workpaper@latest pricing-agent -- --agent
```

The `--agent` starter includes agent-facing files and verification scripts.

## TypeScript API Shape

The WorkPaper runtime shape is still intentionally small: build sheets, write an
input, read the calculated value, and serialize state.

```ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from '@bilig/workpaper'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 20],
    ['Average revenue', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Workbook is missing required sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)

const revenue = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 })
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))

console.log({ revenue, savedBytes: saved.length })
```

## XLSX Recalculation

If the immediate problem is stale formula values in an XLSX file, start with the
small recalculation package before adopting a larger runtime.

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc --demo --json

npx --package @bilig/xlsx-formula-recalc xlsx-recalc quote.xlsx \
  --set Inputs!B2=42 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

SheetJS / `xlsx` boundary:

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc --demo --json
```

ExcelJS boundary:

```sh
npm install exceljs @bilig/exceljs-formula-recalc
npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json
```

## Package Boundaries

| Package | Owns | Does not own |
| --- | --- | --- |
| `@bilig/workbook` | Agent-first model contracts, refs, checks, formulas, operation types, plan/run result contracts. | Runtime execution, formula calculation, UI, consumer business models. |
| `@bilig/formula` | Formula parsing and normalization. | Workbook state or runtime mutation. |
| `@bilig/protocol` | Shared workbook data and protocol shapes. | Engine behavior. |
| `@bilig/core` | Workbook engine, formula calculation, mutation flow, canonical `@bilig/workbook` adapter. | Public consumer model definitions. |
| `@bilig/workpaper` | Headless formula workbook API for Node services and agent tools. | Browser editing UI. |
| `@bilig/headless` | Full lower-level runtime package with agent metadata and subpaths. | A smaller first install for simple model authoring. |

## Package Footprint

<!-- headless-package-footprint:start -->

Current checked npm footprint for `@bilig/headless@0.42.0`:

- Pack dry run: `730 kB` tarball, `4.47 MB` unpacked, `732` package entries.
- Boundary: the main import is the WorkPaper formula/JSON runtime; XLSX
  import/export stays behind the `@bilig/headless/xlsx` subpath; MCP is the
  `bilig-workpaper-mcp` binary wrapper; reduced workbook reports use the
  `bilig-formula-clinic` binary.
- Cold-start gate: Node imports the main entrypoint, builds a two-sheet
  WorkPaper, and reads `24000` under `1000 ms` without importing
  the XLSX subpath.
- Runtime: Node `>=22.0.0`; Node 22 compatibility is covered by the runtime package workflow.
<!-- headless-package-footprint:end -->

## For Agents

When working in this repo:

- For public workbook model work, start at
  [`packages/workbook/README.md`](packages/workbook/README.md),
  [`docs/public-api.md`](docs/public-api.md), and
  [`docs/architecture.md`](docs/architecture.md).
- Do not assume spreadsheet UI behavior is the public contract. The public
  contract is the model, plan, checks, and runtime result.
- Do not hardcode Bilig-owned business models. Consumer models belong to
  consumers.
- Do not parse workbook ref ids. Use labels and `describeRef`.
- Do not treat planned checks as proof. Proof comes from runtime readbacks or
  verified check results.
- For package changes, run the focused package tests before broader checks.

Agent-facing package metadata also lives in:

- [`docs/AGENTS.md`](docs/AGENTS.md)
- [`docs/skill.md`](docs/skill.md)
- [`docs/.well-known/agent.json`](docs/.well-known/agent.json)
- [`docs/llms-full.txt`](docs/llms-full.txt)

The published package also carries `AGENTS.md` and `SKILL.md` so agents
inspecting `node_modules` can find the local contract. Use the
[agent handoff prompt](docs/headless-workpaper-agent-handbook.md#copy-paste-prompt-for-another-agent)
when handing a workbook task to another agent.

## Examples

| Example | What it proves |
| --- | --- |
| [`examples/workbook-agent-model`](examples/workbook-agent-model) | Consumer-defined `@bilig/workbook` model, plan inspection, runtime requirements, execution, and proof result. |
| [`examples/headless-workpaper`](examples/headless-workpaper) | Build, edit, recalculate, serialize, and restore a headless workbook. |
| [`examples/serverless-workpaper-api`](examples/serverless-workpaper-api) | Route-shaped workbook calculation API. |
| [`examples/xlsx-recalculation-node`](examples/xlsx-recalculation-node) | XLSX import, input edit, recalculation, export, and reimport proof. |
| [`examples/recalc-bridge-workflows`](examples/recalc-bridge-workflows) | Recalculation flows across SheetJS, xlsx-populate, and ExcelJS. |
| [`examples/n8n-workpaper-formula-readback`](examples/n8n-workpaper-formula-readback) | n8n workflow for formula-backed readback. |
| [`examples/flowise-workpaper-formula-readback`](examples/flowise-workpaper-formula-readback) | Flowise integration proof. |
| [`examples/dify-workpaper-formula-readback`](examples/dify-workpaper-formula-readback) | Dify integration proof. |

The serverless example also covers `npm run next-server-action` and
`npm run next-server-action-formdata` for framework-boundary smoke tests.

For MCP and agent handoff flows, see:

- [`docs/headless-workpaper-agent-handbook.md`](docs/headless-workpaper-agent-handbook.md)
- [`docs/mcp-workpaper-tool-server.md`](docs/mcp-workpaper-tool-server.md)
- [`docs/agent-workbook-challenge.md`](docs/agent-workbook-challenge.md)

MCP and reduced-case commands:

```sh
npm exec --package @bilig/headless@0.42.0 -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

The MCP tool surface includes `export_workpaper_document`, and `validate_formula`
alongside read and write tools.

## Proof You Can Reproduce

- The 90-second quickstart above edits one input and verifies the dependent formula result.
- Run `pnpm workpaper:bench:competitive:check` to check benchmark artifacts
  locally.
- The benchmark card is generated from
  [`docs/assets/workpaper-benchmark-card.png`](docs/assets/workpaper-benchmark-card.png).
- Read the [compatibility limits](docs/where-bilig-is-not-excel-compatible-yet.md)
  and use the Excel oracle harness when stale cached formula values are not
  trustworthy.
- Public feedback threads:
  [workflow questions](https://github.com/proompteng/bilig/discussions/157),
  [service examples](https://github.com/proompteng/bilig/discussions/213),
  [persistence adapters](https://github.com/proompteng/bilig/discussions/307),
  [JavaScript spreadsheet library guide](https://github.com/proompteng/bilig/discussions/308),
  [OpenAI Responses tool calls](https://github.com/proompteng/bilig/discussions/335),
  and [benchmark critique](https://github.com/proompteng/bilig/discussions/340).
- Star or bookmark the repo at <https://github.com/proompteng/bilig/stargazers>.
- Watch releases at <https://github.com/proompteng/bilig/subscription>.

Discovery links that must remain easy for humans and agents to find:

- [workbook automation examples](docs/workbook-automation-examples-node.md)
- [Node spreadsheet formula engine](docs/node-spreadsheet-formula-engine.md)
- [server-side spreadsheet automation](docs/server-side-spreadsheet-automation-node.md)
- [Google Sheets API boundary](docs/google-sheets-api-alternative-node-workpaper.md)
- [production adoption checklist](docs/production-adoption-checklist-headless-workpaper.md)
- [`examples/serverless-workpaper-api`](examples/serverless-workpaper-api)
- `quote-approval-api`
- [framework adapters](docs/node-framework-workpaper-adapters.md)
- [submit workbook fixture](docs/submit-workbook-fixture.md)
- <https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml>
- <https://github.com/proompteng/bilig/discussions/414>
- [MCP directory status](docs/mcp-spreadsheet-server-directory.md)
- [`examples/headless-workpaper#invoice-totals`](examples/headless-workpaper#invoice-totals)
- [`examples/headless-workpaper#agent-framework-adapters`](examples/headless-workpaper#agent-framework-adapters)
- [`examples/headless-workpaper#mcp-tool-server-shape`](examples/headless-workpaper#mcp-tool-server-shape)
- `agent:framework-adapters`, `agent:mcp-tools`, `agent:mcp-stdio`,
  `agent:openai-agents-sdk`, `agent:openai-responses`,
  `agent:ai-sdk-generate-text`, `agent:ai-sdk-stream-text`
- `openai-responses-workpaper-tool-call`
- `ai-sdk-generate-text-tool-smoke.ts`
- `ai-sdk-stream-text-tool-smoke.ts`
- `npm exec --package @bilig/headless@0.42.0 -- bilig-agent-challenge`
- <https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper>
- [AI SDK and LangChain tools](docs/vercel-ai-sdk-langchain-spreadsheet-tool.md)
- [MCP tool server](docs/mcp-workpaper-tool-server.md)
- [MCP client setup](docs/mcp-client-setup.md)
- [Claude Desktop MCPB WorkPaper](docs/claude-desktop-mcpb-workpaper.md)
- [`examples/headless-workpaper#budget-variance-alerts`](examples/headless-workpaper#budget-variance-alerts)
- [`examples/headless-workpaper#fulfillment-capacity-plan`](examples/headless-workpaper#fulfillment-capacity-plan)
- [`examples/headless-workpaper#quote-approval-threshold`](examples/headless-workpaper#quote-approval-threshold)
- [`examples/headless-workpaper#subscription-mrr-forecast`](examples/headless-workpaper#subscription-mrr-forecast)
- [`docs/javascript-spreadsheet-library-headless-node.md`](docs/javascript-spreadsheet-library-headless-node.md)
- [`docs/sheetjs-exceljs-alternative-formula-workbook-api.md`](docs/sheetjs-exceljs-alternative-formula-workbook-api.md)
- first-timers:
  <https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only>

## Repo Map

- `apps/web`: Vite/React browser workbook shell.
- `apps/bilig`: fullstack monolith runtime and backend APIs.
- `packages/workbook`: agent-first public workbook model API.
- `packages/core`: workbook engine, mutation service, formula execution.
- `packages/formula`: formula language parser and normalizer.
- `packages/protocol`: shared workbook data shapes.
- `packages/workpaper`: headless workbook API.
- `packages/headless`: broader runtime package and agent-facing metadata.
- `packages/excel-import`: XLSX import/export support.
- `packages/*formula-recalc`: focused XLSX recalculation packages.
- `docs`: architecture, public API notes, agent handbooks, package trust docs.
- `examples`: runnable proofs.
- `integrations`: n8n and Pipedream integrations.
- `scripts`: release, CI, corpus, docs, and validation automation.

## Local Development

Use Node `24+`, Bun, and `pnpm@10.32.1`.

```sh
pnpm install
pnpm dev:web-local
```

Useful checks:

```sh
pnpm --filter @bilig/workbook test
pnpm --filter @bilig/workbook build
pnpm typecheck
pnpm lint
pnpm build
pnpm run ci
```

Common dev commands:

```sh
pnpm dev:web        # browser shell
pnpm dev:sync       # monolith runtime
pnpm test           # unit tests
pnpm test:browser   # Playwright tests
pnpm coverage       # coverage run
```

If you edit protocol or formula inventory sources, regenerate and commit the
generated outputs. CI fails on dirty tracked files.

## Public Review

Start with these docs when evaluating adoption:

- [`docs/why-use-bilig.md`](docs/why-use-bilig.md)
- [`docs/public-api.md`](docs/public-api.md)
- [`docs/npm-provenance-package-trust.md`](docs/npm-provenance-package-trust.md)
- [`docs/what-workpaper-benchmark-proves.md`](docs/what-workpaper-benchmark-proves.md)
- [`docs/submit-workbook-fixture.md`](docs/submit-workbook-fixture.md)
- formula workbooks proof page:
  [`docs/formula-workbooks-node-services-agent-tools.md`](docs/formula-workbooks-node-services-agent-tools.md)
- stale XLSX formula cache:
  [`docs/stale-xlsx-formula-cache-node.md`](docs/stale-xlsx-formula-cache-node.md)
- SheetJS formula result not updating:
  [`docs/sheetjs-formula-result-not-updating-node.md`](docs/sheetjs-formula-result-not-updating-node.md)
- Microsoft Graph Excel recalculation:
  [`docs/microsoft-graph-excel-recalculation-node.md`](docs/microsoft-graph-excel-recalculation-node.md)
- Google Sheets API boundary:
  [`docs/google-sheets-api-alternative-node-workpaper.md`](docs/google-sheets-api-alternative-node-workpaper.md)
- XLSX formula recalculation:
  [`docs/xlsx-formula-recalculation-node.md`](docs/xlsx-formula-recalculation-node.md)
- agent-owned XLSX recalculation:
  [`docs/agent-xlsx-formula-recalculation-without-libreoffice.md`](docs/agent-xlsx-formula-recalculation-without-libreoffice.md)
- file-level calculation engine:
  [`docs/excel-file-calculation-engine-node.md`](docs/excel-file-calculation-engine-node.md)
- ExcelJS shared formula recalculation:
  [`docs/exceljs-shared-formula-recalculation-node.md`](docs/exceljs-shared-formula-recalculation-node.md)

Checked benchmark evidence is file-backed. The current WorkPaper vs
HyperFormula artifact shows
[`94/100` comparable WorkPaper mean wins](docs/what-workpaper-benchmark-proves.md).
The current worst p95 row is `structural-move-rows` at `4.047x`, so do not make
blanket performance claims beyond the checked evidence.

For first-time contributors, [`docs/why-use-bilig.md`](docs/why-use-bilig.md)
and the public API docs include acceptance commands for first patches.

Trust checks for `@bilig/headless@0.42.0`:

```sh
npm view @bilig/headless@latest version dist.attestations dist.signatures --json
```

The npm provenance and package trust guide explains the release path.
Repository security posture is tracked by
`https://api.scorecard.dev/projects/github.com/proompteng/bilig/badge` and
uploaded to GitHub code scanning on every `main` update.

## Security

Read [`SECURITY.md`](SECURITY.md) and [`SUPPORT.md`](SUPPORT.md) before sharing
reports. Report security issues privately through the repository security
policy. Do not open public issues with workbook data, credentials, tokens, or
exploit details.

For formula or workbook correctness bugs, open a reduced public fixture when
possible:
[`docs/submit-workbook-fixture.md`](docs/submit-workbook-fixture.md).

## License

MIT. See [`LICENSE`](LICENSE).
