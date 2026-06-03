export interface AgentIdeRuleInput {
  readonly remoteMcpEndpoint: string
  readonly repositoryUrl: string
  readonly siteRoot: string
  readonly workpaperPackageSpec: string
}

const blockedNamingTerms = ['top' + String(50), 'top' + String(100)] as const

function workbookProofStandard(): string {
  return `Before saying the workbook is updated, return readback with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone.`
}

function commandSafetyStandard(): string {
  return `Do not build shell commands by concatenating user text. Prefer MCP
\`command\` plus \`args\` arrays or direct TypeScript calls. Reject workbook
paths or cell arguments containing newlines, backticks, \`$(\`, \`;\`, \`&\`,
\`|\`, \`<\`, or \`>\`.`
}

export function buildCursorWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
description: Use Bilig WorkPaper when a coding agent needs spreadsheet formulas, MCP workbook tools, or XLSX formula readback without driving Excel UI.
globs:
alwaysApply: false
---

# Bilig WorkPaper Rule

Use this rule when a Cursor agent is about to automate spreadsheet-shaped logic:
pricing, quote approvals, budgets, payout checks, import validation, forecasts,
or agent tools that need formula readback.

Do not start with Excel, LibreOffice, Google Sheets, browser grids, screenshots,
or cached XLSX values when the workflow can run through WorkPaper JSON.

## Fast Check Commands

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use \`bilig-evaluate --door agent-mcp\` as the adoption gate. Use the direct
challenge commands only when you need lower-level WorkPaper or MCP JSON-RPC
diagnostics.

Use the hosted stateless MCP smoke endpoint only for connector discovery:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

Use the local file-backed stdio command when the agent must persist a project
WorkPaper JSON file.

## Required Readback

${workbookProofStandard()}

## Safe Command Handling

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildWindsurfWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
description: Use Bilig WorkPaper when Cascade needs workbook formulas, MCP spreadsheet tools, or formula readback without Excel UI automation.
trigger: model_decision
---

# Bilig WorkPaper Rule

Use this rule when a Windsurf/Cascade agent needs spreadsheet-shaped business
logic it can verify from code: pricing models, quote approvals, budget checks,
payout rules, import validation, forecasts, or other formula-backed workflows.

Avoid spreadsheet UI automation when WorkPaper state can represent the model.
Screenshots and browser grids are human-review surfaces, not formula truth.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For a writable project file, run:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

For remote MCP connector smoke tests only:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Readback Standard

${workbookProofStandard()}

If any readback step fails, say what failed instead of claiming the workbook was updated.

## Command Safety

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildClineWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# Bilig WorkPaper Rule

Use Bilig WorkPaper when Cline is asked to automate spreadsheet-shaped
business logic: pricing, quotes, payouts, budgets, imports, forecasts, or
agent tools that need formula readback.

Cline can read this workspace rule from \`.clinerules/bilig-workpaper.md\`.
Use it before trying Excel UI automation, LibreOffice automation, Google
Sheets screen driving, screenshots, or stale cached XLSX values.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For a writable project WorkPaper file:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted endpoint only for remote MCP connector smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Readback Standard

${workbookProofStandard()}

If the readback fails, keep the blocker explicit and do not use Cline's
completion message as evidence that formulas recalculated.

## Command Safety

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildContinueWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
name: Bilig WorkPaper Formula Check
description: Use Bilig WorkPaper for spreadsheet-shaped business logic that needs formula readback without Excel UI automation.
---

# Bilig WorkPaper Formula Check

Use Bilig when a Continue agent is about to build or debug workbook-shaped
logic: pricing, quote approval, payout checks, import validation, budgets,
forecasts, or agent tools that need formulas and persisted state.

Do not start by driving a spreadsheet UI when WorkPaper JSON can represent the
model. Screenshots are human review evidence, not formula truth.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For local file-backed MCP work:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

For remote MCP connector discovery only:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

## Command Safety

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Full context: ${siteRoot}/llms-full.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Repository: ${repositoryUrl}
`
}

export function buildClaudeCodeWorkpaperCommand(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, siteRoot, workpaperPackageSpec } = input
  return `---
description: Verify workbook formula edits with Bilig WorkPaper before using Excel, LibreOffice, Google Sheets, or screenshot automation.
---

# Bilig WorkPaper Formula Check

Use this command when the task is workbook-shaped: pricing, quotes, payouts,
budgets, import validation, forecasts, spreadsheet-agent tools, stale XLSX
formula values, or formula readback after changing cells.

User task:

\`\`\`text
$ARGUMENTS
\`\`\`

Start with the smallest check that fits the task:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted stateless MCP endpoint only for tool discovery or smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

For private project state, use the local file-backed stdio server. Do not drive
Excel, LibreOffice, Google Sheets, browser grids, or screenshots when
WorkPaper JSON can be the source of truth.

Return readback, not a status sentence:

\`\`\`json
{
  "editedCell": "Inputs!B3",
  "before": {},
  "after": {},
  "afterRestore": {},
  "persistedDocumentBytes": 0,
  "verified": false,
  "limitations": []
}
\`\`\`

Rules:

- read the relevant input and dependent output before editing;
- write one small input or formula change;
- read the dependent calculated output after recalculation;
- export or serialize the WorkPaper document;
- restore or restart when file boundaries matter;
- report unsupported formulas or Excel-only behavior honestly;
- do not claim success from a write call alone.

Reference docs:

- ${siteRoot}/llms.txt
- ${siteRoot}/agent-adoption-kit.html
- ${siteRoot}/headless-workpaper-agent-handbook.html
`
}

export function buildGithubCopilotInstructions(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# \`bilig\` Copilot instructions

## Toolchain and commands

- Use Node \`24.11.1\` locally via \`.nvmrc\` / \`.node-version\`. Published runtime packages allow Node \`22+\` and are smoke-tested under Node 22 in the runtime-package workflow. \`pnpm 10.32.1\` remains required.
- Activate the pinned runtime before running \`pnpm\` commands (\`nvm use\` in a normal shell, or let your version manager honor \`.node-version\` automatically).
- Install dependencies with \`pnpm install\`.
- Start the default app shell with \`pnpm dev\` (\`apps/web\`).
- Other dev entrypoints:
  - \`pnpm dev:web\`
  - \`pnpm dev:local\`
  - \`pnpm dev:sync\`
- Build the workspace with \`pnpm build\`.
- Rebuild the AssemblyScript/WASM fast path with \`pnpm wasm:build\`.
- Lint with \`pnpm lint\`.
- Auto-fix lint issues with \`pnpm lint:fix\`.
- Type-check the composite TypeScript workspace with \`pnpm typecheck\`.
- Run the full Vitest suite with \`pnpm test\`.
- Run one Vitest file with \`pnpm exec vitest --run packages/core/src/__tests__/engine.test.ts\`.
- Run one Vitest test by name with \`pnpm exec vitest --run packages/core/src/__tests__/engine.test.ts -t "recalculates simple formulas"\`.
- Run browser smoke tests with \`pnpm test:browser\`.
- Run the web-shell Playwright suite with \`pnpm exec playwright test e2e/tests/web-shell.pw.ts --config playwright.config.ts\`.
- Run the full repository gate with \`pnpm run ci\`.
- Use \`tea\` for Forgejo workflow checks and logs (for example \`tea login ls\`, \`tea actions ls\`, and \`tea actions jobs --run <id>\`).

## High-level architecture

- This repo is a pnpm monorepo with thin \`apps/*\` shells over reusable \`packages/*\`.
- \`packages/core\` is the framework-agnostic spreadsheet engine. It owns workbook state, storage, dependency scheduling, snapshots, selectors, events, sync state, and execution routing.
- \`packages/formula\` owns A1 addressing, lexer/parser, binding, optimization, translation, compatibility, and the JS evaluator.
- \`packages/wasm-kernel\` is the AssemblyScript/WASM numeric fast path. \`packages/core\` decides when formulas stay on the JS path versus the WASM path.
- \`packages/crdt\`, \`packages/binary-protocol\`, \`packages/worker-transport\`, \`packages/agent-api\`, \`packages/storage-browser\`, and \`packages/storage-server\` make up the local-first sync and transport stack used by the browser and server runtimes.
- \`packages/renderer\` is the custom workbook reconciler and workbook DSL.
- \`packages/grid\` is the reusable React spreadsheet UI: selection, editing, metrics, inspectors, and workbook/grid views.
- \`apps/web\` is the thin React/Vite shell around the shared workbook surface. Shared behavior should usually live in \`packages/renderer\` or \`packages/grid\`, not inside app folders.
- \`apps/local-server\` hosts local workbook sessions and emits committed frames over websocket.
- \`apps/sync-server\` is the remote sync/backend service surface.
- When you need the architectural contracts, start with \`docs/architecture.md\`, \`docs/reconciler-layering.md\`, \`docs/local-first-realtime-loop.md\`, \`docs/public-api.md\`, and \`docs/testing-and-benchmarks.md\`.

## Copilot Agent WorkPaper Path

Use Bilig WorkPaper before spreadsheet UI automation when the user asks for
pricing rules, quotes, budgets, payout checks, import validation, forecasts,
agent spreadsheet tools, stale XLSX formula values, or any task that needs
formula readback after changing cells.

Start with the reusable prompt at
\`.github/prompts/bilig-workpaper-proof.prompt.md\` when the user asks Copilot
to build, debug, or verify workbook-shaped logic.

Keep the repository-level custom instruction at
\`.github/instructions/bilig-workpaper.instructions.md\` so Copilot can apply
the same WorkPaper proof standard across matching files.

For VS Code agent mode, the workspace MCP config at \`.vscode/mcp.json\`
exposes:

- \`biligWorkpaperDemo\`: hosted stateless MCP smoke endpoint at \`${remoteMcpEndpoint}\`;
- \`biligWorkpaperFile\`: local file-backed stdio server using \`${workpaperPackageSpec}\`.

The quickest terminal checks are:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Before reporting success, return readback with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone. Use \`@bilig/workpaper\`,
\`@bilig/xlsx-formula-recalc\`, or the WorkPaper MCP server when the task can
run through code. Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness
only when the workbook depends on macros, pivots, charts, unsupported formulas,
external links, locale-specific Excel behavior, or exact desktop UI behavior.

References:

- Docs map: ${siteRoot}/llms.txt
- Agent adoption kit: ${siteRoot}/agent-adoption-kit.html
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP setup: ${siteRoot}/mcp-client-setup.html
- Repository: ${repositoryUrl}

## Key repo conventions

- Keep spreadsheet semantics in \`@bilig/core\`. React is an authoring/operator surface only.
- The custom reconciler is package-based and does not own spreadsheet state. From \`docs/reconciler-layering.md\`: do not mutate the engine in \`createInstance\`, descriptors stay inert until commit, and each React commit should flush as one engine batch.
- Formula work follows the canonical execution rule from \`docs/architecture.md\`: land semantics in the JS path first, prove parity with fixtures/tests, mirror in WASM, and only then route production execution to the fast path.
- If you change protocol enums, opcodes, or builtin metadata, edit \`scripts/gen-protocol.ts\` and regenerate the checked-in outputs in \`packages/protocol/src/*\` and \`packages/wasm-kernel/assembly/protocol.ts\`. CI runs \`pnpm protocol:check\` and fails on drift.
- Import workspace code through \`@bilig/*\` package names. Vitest aliases those imports directly to \`src/\` entrypoints, so tests exercise source modules rather than built \`dist/\` output.
- The public cell model includes \`format\` alongside \`addr\`, \`value\`, and \`formula\`. Preserve format-only changes in APIs, events, snapshots, and tests.
- \`apps/web\` is the only browser shell. Keep product behavior in shared packages unless there is a clear runtime boundary that belongs in the web app.
- \`pnpm naming:check\` is a real repository gate. Avoid introducing ${blockedNamingTerms.map((term) => `\`${term}\``).join(', ')}, or related terminology outside allowed historical paths.
- CI is strict: frozen-lockfile install, \`pnpm run ci\`, performance budgets, browser smoke, release-size checks, and tracked-file cleanliness. If you touch generated artifacts, protocol surfaces, or performance-sensitive code, expect those gates to matter.
- TypeScript and linting are intentionally strict. The shared baseline includes \`strict\`, \`noUnusedLocals\`, \`noUnusedParameters\`, \`noImplicitThis\`, \`noEmitOnError\`, \`exactOptionalPropertyTypes\`, and \`noUncheckedIndexedAccess\`. Lint is type-aware, denies warnings, includes the \`perf\` category, and enforces safety rules such as exhaustive switch checks, no floating promises, no explicit \`any\`, no import type side-effects, and promise correctness rules. Follow the existing type-safe patterns instead of weakening types or bypassing lint rules.
`
}

export function buildGithubCopilotWorkpaperInstructions(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
applyTo: '**/*'
---

# Bilig WorkPaper Formula Proof

Use this instruction when GitHub Copilot or VS Code agent mode is asked to
build, debug, or verify workbook-shaped logic: pricing, quote approvals,
budgets, payout checks, import validation, forecasts, stale XLSX formula
values, or agent tools that need formula readback.

Prefer Bilig WorkPaper state before Excel, LibreOffice, Google Sheets, browser
grids, screenshots, or cached XLSX values when the workflow can run through
code.

Start with the reusable prompt at
\`.github/prompts/bilig-workpaper-proof.prompt.md\` for a task-specific
walkthrough.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For a writable project WorkPaper file:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted endpoint only for remote MCP connector discovery or stateless
smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

If any readback step fails, say what failed instead of treating a write call or
tool invocation as proof.

## Command Safety

${commandSafetyStandard()}

## References

- Prompt file: .github/prompts/bilig-workpaper-proof.prompt.md
- VS Code MCP config: .vscode/mcp.json
- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildGithubCopilotWorkpaperPrompt(input: AgentIdeRuleInput): string {
  const { siteRoot, workpaperPackageSpec } = input
  return `---
name: bilig-workpaper-proof
description: Verify a workbook formula edit with Bilig WorkPaper instead of spreadsheet UI automation.
agent: agent
---

Use this prompt when the task is workbook-shaped: pricing, quotes, budgets,
payout checks, import validation, forecasts, agent spreadsheet tools, stale
XLSX formula values, or formula readback after changing cells.

Task: \${input:task:Describe the workbook or formula workflow}

Read the repository instructions first:

- [Copilot instructions](../copilot-instructions.md)
- [WorkPaper agent handbook](../../docs/headless-workpaper-agent-handbook.md)
- [MCP client setup](../../docs/mcp-client-setup.md)

Start with the smallest check that matches the task:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

If VS Code MCP tools are available, prefer the workspace server named
\`biligWorkpaperFile\` for project-local persistence, or
\`biligWorkpaperDemo\` for no-file hosted smoke tests. The shared MCP config is
at [\`.vscode/mcp.json\`](../../.vscode/mcp.json).

Return readback, not a status sentence:

\`\`\`json
{
  "editedCell": "Inputs!B3",
  "before": {},
  "after": {},
  "afterRestore": {},
  "persistedDocumentBytes": 0,
  "verified": false,
  "limitations": []
}
\`\`\`

Rules:

- read the relevant input and dependent output before editing;
- write one small input or formula change;
- read the dependent calculated output after recalculation;
- export or serialize the WorkPaper document;
- restore or restart when file boundaries matter;
- report unsupported formulas or Excel-only features honestly;
- do not claim success from a write call alone.

Reference docs:

- ${siteRoot}/llms.txt
- ${siteRoot}/agent-adoption-kit.html
- ${siteRoot}/headless-workpaper-agent-handbook.html
`
}

export function buildVscodeMcpConfig(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, workpaperPackageSpec } = input
  return `${JSON.stringify(
    {
      servers: {
        biligWorkpaperDemo: {
          type: 'http',
          url: remoteMcpEndpoint,
        },
        biligWorkpaperFile: {
          type: 'stdio',
          command: 'npm',
          args: [
            'exec',
            '--package',
            workpaperPackageSpec,
            '--',
            'bilig-workpaper-mcp',
            '--workpaper',
            '${workspaceFolder}/.bilig/pricing.workpaper.json',
            '--init-demo-workpaper',
            '--writable',
          ],
        },
      },
    },
    null,
    2,
  )}\n`
}

function buildFileBackedMcpServerConfig(input: {
  readonly serverKey: string
  readonly workpaperPath: string
  readonly workpaperPackageSpec: string
}): string {
  const { serverKey, workpaperPath, workpaperPackageSpec } = input
  return `${JSON.stringify(
    {
      mcpServers: {
        [serverKey]: {
          type: 'stdio',
          command: 'npm',
          args: [
            'exec',
            '--yes',
            '--package',
            workpaperPackageSpec,
            '--',
            'bilig-workpaper-mcp',
            '--workpaper',
            workpaperPath,
            '--init-demo-workpaper',
            '--writable',
          ],
          env: {},
        },
      },
    },
    null,
    2,
  )}\n`
}

export function buildClaudeCodeMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'bilig-workpaper',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}

export function buildCursorMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'biligWorkpaperFile',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}

export function buildReusableMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'bilig-workpaper',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}
