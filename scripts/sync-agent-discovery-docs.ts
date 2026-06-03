import { createHash } from 'node:crypto'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { agentNotAFitBoundaries, mcpPromptNames } from './agent-discovery-constants.ts'
import { buildDocsAgentInstructions, buildDocsAgentStart } from './agent-discovery-agent-instructions.ts'
import { buildLlmsFullSources } from './agent-discovery-llms-full-sources.ts'
import { mcpServerCardManifest } from './agent-discovery-mcp-card.ts'
import { buildWorkpaperPackageAgentInstructions, buildWorkpaperPackageSkillDocument } from './agent-discovery-package-docs.ts'
import { readTextFileIfExists } from './read-if-exists.ts'
import { syncVersionedStaticReferences } from './sync-agent-static-references.ts'
import { buildEvaluatorDoors, buildProofContract, compactProofContractJsonArrays } from './agent-discovery-evaluator-doors.ts'
import {
  buildClineWorkpaperRule,
  buildClaudeCodeWorkpaperCommand,
  buildContinueWorkpaperRule,
  buildCursorWorkpaperRule,
  buildGithubCopilotInstructions,
  buildGithubCopilotWorkpaperPrompt,
  buildVscodeMcpConfig,
  buildWindsurfWorkpaperRule,
} from './agent-discovery-ide-rules.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const siteRoot = 'https://proompteng.github.io/bilig'
const skillDiscoveryRoot = 'https://bilig.proompteng.ai'
const remoteMcpEndpoint = 'https://bilig.proompteng.ai/mcp'
const remoteMcpAliasEndpoint = 'https://bilig.proompteng.ai/mcp/workpaper'
const remoteMcpServerCard = 'https://bilig.proompteng.ai/.well-known/mcp/server-card.json'
const repositoryUrl = 'https://github.com/proompteng/bilig'
const skillName = 'bilig-workpaper'
const skillManifestUrl = `${skillDiscoveryRoot}/.well-known/agent-skills/${skillName}/SKILL.txt`
const skillDiscoverySchemaUrl = 'https://schemas.agentskills.io/discovery/0.2.0/schema.json'
const headlessPackageVersion = parsePackageVersion(await readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'))
const headlessPackageSpec = `@bilig/headless@${headlessPackageVersion}`
const workpaperPackageSpec = '@bilig/workpaper@latest'
const workbookPackageSpec = '@bilig/workbook@latest'
const unscopedWorkpaperPackageSpec = 'bilig-workpaper@latest'
const mcpbReleaseAssetUrl = `${repositoryUrl}/releases/latest/download/bilig-workpaper.mcpb`
const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`
const checkOnly = process.argv.includes('--check')

function parsePackageVersion(packageJson: string): string {
  const parsed: unknown = JSON.parse(packageJson)
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('packages/headless/package.json must be an object')
  }
  const version = Reflect.get(parsed, 'version')
  if (typeof version !== 'string') {
    throw new Error('packages/headless/package.json must define a string version')
  }
  return version
}

const docsAgentInstructions = buildDocsAgentInstructions({
  mcpbReleaseAssetUrl,
  mcpbReleaseChecksumUrl,
  remoteMcpEndpoint,
  remoteMcpServerCard,
  siteRoot,
  skillDiscoveryRoot,
  skillManifestUrl,
  workpaperPackageSpec,
})

const docsAgentStart = buildDocsAgentStart({
  mcpbReleaseAssetUrl,
  mcpbReleaseChecksumUrl,
  remoteMcpEndpoint,
  remoteMcpServerCard,
  siteRoot,
  skillDiscoveryRoot,
  skillManifestUrl,
  workpaperPackageSpec,
})

const headlessPackageAgentInstructions = `# @bilig/headless agent notes

If you are editing this package inside the Bilig monorepo, follow the root
\`AGENTS.md\` first. These notes are also published with the npm package so an
agent inspecting \`node_modules/@bilig/headless\` can use WorkPaper state without
driving a spreadsheet UI.

## Use WorkPaper first

Use \`@bilig/workpaper\` first when the task is workbook-shaped business logic:
pricing, quote approval, budget checks, payout models, import validation,
forecasting, or agent tools that need formula readback.

Use \`@bilig/headless\` only when the task explicitly targets lower-level
runtime subpaths, package-boundary audits, provenance checks, or internals of
this package.

Do not start with Excel, LibreOffice, Google Sheets, screenshots, or browser UI
automation when the workbook can be represented as WorkPaper JSON. Screenshots
are for human visual review, not formula truth.

## Handoff prompt

Paste this into another coding agent when it is about to automate a spreadsheet
through a desktop app or browser grid:

\`\`\`text
Before using spreadsheet UI automation, check whether this is workbook-shaped
logic that can run through @bilig/workpaper. If it is, use WorkPaper state as the
source of truth.

For triage, start with:
npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json

For MCP proof, run:
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable

For Node or TypeScript, import @bilig/workpaper directly. Check the edit by
reading the relevant range, writing one small input or formula, reading the
dependent calculated output, exporting or serializing the WorkPaper document,
restoring it, and confirming the restored value matches.

Return editedCell, before, after, afterRestore, persistedDocumentBytes,
verified, and limitations. Do not claim success from a write call alone.
\`\`\`

## Minimum edit loop

For every agent-owned workbook edit:

1. identify the exact sheet and A1 cell or range.
2. read the current input and dependent output.
3. validate formulas before writing them.
4. write one small change.
5. read the dependent computed output after recalculation.
6. serialize or export the WorkPaper document.
7. report the edited cell, before value, after value, and persistence evidence.

Do not report success from a write call alone.

## MCP entrypoint

For MCP clients, use the published stdio server:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Expected file-backed tools:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`set_cell_contents_and_readback\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`

Use \`--init-demo-workpaper\` when the path may not exist yet; it creates the demo
WorkPaper JSON only when the file is missing. Use \`--writable\` only when the
task should persist \`set_cell_contents\` edits back to the same WorkPaper JSON
file.

Claude Desktop users can skip manual JSON config by installing the released
MCPB bundle:

- ${mcpbReleaseAssetUrl}
- ${mcpbReleaseChecksumUrl}

## Direct TypeScript entrypoint

Use \`@bilig/workpaper\` when the workbook logic belongs in a service, queue
worker, test, or route:

\`\`\`ts
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
\`\`\`

## Verification shortcuts

From a clean project, run the package-owned check:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

\`bilig-agent-challenge\` checks the direct WorkPaper API loop.
\`bilig-mcp-challenge\` checks the file-backed MCP JSON-RPC loop. A good run
prints \`verified: true\`.

When the task explicitly targets this lower-level \`@bilig/headless\` package,
run the same checks against this package boundary:

\`\`\`sh
npm exec --package ${headlessPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Deeper docs:

- <https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html>
- <https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html>
- <https://proompteng.github.io/bilig/mcp-client-setup.html>
`

const skillDocument = `---
name: bilig-workpaper
version: 0.1.0
description: Use @bilig/workpaper WorkPaper state for workbook formulas, agent spreadsheet tools, MCP file-backed or remote demo editing, and XLSX formula bug reports without driving spreadsheet UI.
tags:
  - ai-agents
  - spreadsheet-automation
  - formulas
  - xlsx
  - mcp
  - typescript
---

# Bilig WorkPaper Agent Skill

Use this skill when an agent needs spreadsheet-style formulas but the work should run through files, terminal commands, TypeScript, HTTP routes, or MCP tools instead of Excel UI automation.

## When To Trigger

Trigger this skill for tasks involving:

- workbook-shaped business logic in Node.js services;
- formula readback after writing cells;
- quote, budget, payout, pricing, import-validation, or forecast models;
- agent spreadsheet tools that need deterministic cell addresses;
- MCP clients that can run a stdio server or call a Streamable HTTP endpoint;
- reduced XLSX formula bugs that need a local report.

Do not trigger it for manual spreadsheet editing, Office macros, VBA, pivots, charts, COM automation, or exact Excel desktop behavior unless the user explicitly asks to compare Bilig against an Excel oracle.

## Command Safety

Do not build shell commands by concatenating user text. Treat the commands below as literal templates, validate workbook paths before use, and reject values containing newlines, backticks, \`$(\`, \`;\`, \`&\`, \`|\`, \`<\`, or \`>\`. Prefer MCP client \`command\` plus \`args\` arrays or direct TypeScript calls when inserting user-provided paths or cell references.

## First Check: Agent Triage

Before wiring a client or opening a spreadsheet UI, print the compact decision
card:

\`\`\`json
{
  "command": "npm",
  "args": ["exec", "--yes", "--package", "${workpaperPackageSpec}", "--", "bilig-agent-start", "--json"]
}
\`\`\`

## First Check: Agent Evaluator

Before wiring a client, prove the published agent door with the package-owned evaluator.
It exercises MCP discovery, cell mutation, formula readback, JSON export, restart restore, and returns \`verified: true\`:

\`\`\`json
{
  "command": "npm",
  "args": ["exec", "--yes", "--package", "${workpaperPackageSpec}", "--", "bilig-evaluate", "--door", "agent-mcp", "--json"]
}
\`\`\`

For service-owned WorkPaper logic without MCP, run \`bilig-evaluate --door workpaper-service --json\`.
Use the lower-level challenge commands only when debugging the direct API loop or file-backed MCP JSON-RPC transcript:

\`\`\`json
[
  { "command": "npm", "args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-agent-challenge", "--json"] },
  { "command": "npm", "args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-mcp-challenge", "--json"] }
]
\`\`\`

## First Choice: MCP

Use MCP when the host can run a stdio server or call a Streamable HTTP server.
Configure stdio as an argument array, not a shell-concatenated string:

If the host supports installable skills, first check that the public skill
package is discoverable:

\`\`\`sh
npx --yes skills@latest add ${skillDiscoveryRoot} --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
\`\`\`

\`\`\`json
{
  "command": "npm",
  "args": [
    "exec",
    "--package",
    "${workpaperPackageSpec}",
    "--",
    "bilig-workpaper-mcp",
    "--workpaper",
    "./pricing.workpaper.json",
    "--init-demo-workpaper",
    "--writable"
  ]
}
\`\`\`

Run \`bilig-evaluate --door agent-mcp --json\` first. If the workbook contains
provider-backed formulas such as \`IMPORTRANGE\`, run
\`bilig-evaluate --door agent-mcp --scenario provider-backed --json\` to confirm
the adapter boundary. If the evaluator fails, run \`bilig-mcp-challenge\` and
treat its returned \`tools\` array as the source of truth for the currently published package. The core file-backed tools are:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`set_cell_contents_and_readback\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`

After a write, always read the dependent output cell and export the WorkPaper
document. If the listed tool set includes \`set_cell_contents_and_readback\`,
prefer it for stateless clients because the edit and dependent readback happen
in one tool call. If it is absent, call \`set_cell_contents\`, then \`read_cell\`
or \`read_range\`, then \`export_workpaper_document\`.

For remote MCP clients, use the stateless demo endpoint when the client supports
Streamable HTTP:

\`\`\`text
${remoteMcpEndpoint}
${remoteMcpAliasEndpoint}
\`\`\`

The remote endpoint is request-local and does not write user files. Use it for
connector smoke tests, tool discovery, and agent onboarding; use the file-backed
stdio command when the workflow must persist a project WorkPaper JSON file.

## Second Choice: Direct TypeScript

Use \`@bilig/workpaper\` directly when workbook logic belongs in a service, queue worker, test, or route:

\`\`\`ts
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
\`\`\`

## XLSX Formula Clinic

When the user has a reduced XLSX formula/import bug, generate a local report through an argument array:

\`\`\`json
{
  "command": "npm",
  "args": [
    "exec",
    "--package",
    "${workpaperPackageSpec}",
    "--",
    "bilig-formula-clinic",
    "./reduced.xlsx",
    "--cells",
    "Summary!B7,Inputs!B2"
  ]
}
\`\`\`

The report is local. It does not upload workbook contents. Ask for a reduced public fixture rather than private customer spreadsheets.

## Required Verification

Return readback, not vibes. A successful agent response should include:

- the exact edited sheet and A1 cell;
- before values for relevant inputs and dependent outputs;
- after values read from the recalculated workbook;
- persistence evidence from serialized or exported WorkPaper state;
- restore or reimport checks when file boundaries matter;
- limitations for unsupported formulas or Excel-only features.

If any readback step fails, report the blocker instead of claiming the workbook was updated.

## Reference URLs

- Compact docs map: ${siteRoot}/llms.txt
- Full agent context: ${siteRoot}/llms-full.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Agent workbook challenge: ${siteRoot}/agent-workbook-challenge.html
- MCP server guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Open WebUI tool setup: ${siteRoot}/open-webui-workpaper-mcp.html
- LobeHub MCP setup: ${siteRoot}/lobehub-workpaper-mcp.html
- AnythingLLM MCP setup: ${siteRoot}/anythingllm-workpaper-mcp.html
- Sim MCP setup: ${siteRoot}/sim-workpaper-mcp.html
- FastMCP Python client: ${siteRoot}/fastmcp-workpaper-client.html
- Agno WorkPaper MCP tools: ${siteRoot}/agno-workpaper-mcp.html
- Pydantic AI WorkPaper MCP tools: ${siteRoot}/pydantic-ai-workpaper-mcp.html
- smolagents WorkPaper tool: ${siteRoot}/smolagents-workpaper-tool.html
- Hugging Face WorkPaper Space template: ${siteRoot}/huggingface-workpaper-space.html
- Windmill TypeScript script: ${siteRoot}/windmill-workpaper-script.html
- Trigger.dev task: ${siteRoot}/triggerdev-workpaper-task.html
- Inngest step: ${siteRoot}/inngest-workpaper-step.html
- Airbyte validation: ${siteRoot}/airbyte-workpaper-validation.html
- Meltano utility: ${siteRoot}/meltano-workpaper-utility.html
- Temporal Activity: ${siteRoot}/temporal-workpaper-activity.html
- Airflow DAG: ${siteRoot}/airflow-workpaper-dag.html
- Dagster asset: ${siteRoot}/dagster-workpaper-asset.html
- Kestra Node flow: ${siteRoot}/kestra-workpaper-flow.html
- Prefect flow: ${siteRoot}/prefect-workpaper-flow.html
- XLSX formula clinic: ${siteRoot}/formula-bug-clinic.html
- Compatibility limits: ${siteRoot}/where-bilig-is-not-excel-compatible-yet.html
- Repository: ${repositoryUrl}
`

const workpaperPackageAgentInstructions = buildWorkpaperPackageAgentInstructions({
  headlessPackageAgentInstructions,
  headlessPackageSpec,
  unscopedWorkpaperPackageSpec,
  workpaperPackageSpec,
})

const workpaperPackageSkillDocument = buildWorkpaperPackageSkillDocument({
  skillDocument,
  workpaperPackageSpec,
  unscopedWorkpaperPackageSpec,
})

const llmsFullSources = buildLlmsFullSources(repositoryUrl)

function skillIndexJson(): string {
  const skillDigest = createHash('sha256').update(skillDocument).digest('hex')
  return `${JSON.stringify(
    {
      $schema: skillDiscoverySchemaUrl,
      skills: [
        {
          name: skillName,
          type: 'skill-md',
          description:
            'Use @bilig/workpaper WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.',
          url: skillManifestUrl,
          digest: `sha256:${skillDigest}`,
        },
      ],
    },
    null,
    2,
  )}\n`
}

function agentJsonManifest(): string {
  const json = JSON.stringify(
    {
      schema_version: 'agent-json-0.1.0',
      name: 'bilig',
      title: 'Bilig WorkPaper formula runtime',
      description:
        'Formula WorkPaper runtime for Node.js services and agent tools: edit cells, recalculate, verify readback, and persist JSON without spreadsheet UI automation.',
      url: `${siteRoot}/`,
      repository: repositoryUrl,
      license: 'MIT',
      contact: `${repositoryUrl}/discussions/new?category=general`,
      llms_txt: `${siteRoot}/llms.txt`,
      llms_full: `${siteRoot}/llms-full.txt`,
      llms_install: `${siteRoot}/llms-install.html`,
      llms_install_source: `${repositoryUrl}/blob/main/llms-install.md`,
      well_known_llms_txt: `${siteRoot}/.well-known/llms.txt`,
      well_known_llms_full: `${siteRoot}/.well-known/llms-full.txt`,
      agent_start: `${siteRoot}/agent-start.txt`,
      well_known_agent_start: `${siteRoot}/.well-known/agent-start.txt`,
      skill_file: skillManifestUrl,
      agent_instructions: `${siteRoot}/AGENTS.md`,
      adoption_kit: `${siteRoot}/agent-adoption-kit.html`,
      skills: [
        {
          name: skillName,
          url: skillManifestUrl,
          index_url: `${siteRoot}/.well-known/agent-skills/index.json`,
          description:
            'Use @bilig/workpaper WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.',
        },
      ],
      evaluator_doors: buildEvaluatorDoors({ repositoryUrl, siteRoot, workpaperPackageSpec }),
      mcp: {
        server_name: 'io.github.proompteng/bilig-workpaper',
        server_card: `${siteRoot}/.well-known/mcp/server-card.json`,
        remote_server_card: remoteMcpServerCard,
        manifest: `${siteRoot}/.well-known/mcp.json`,
        registry_search: 'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
        remote_endpoint: remoteMcpEndpoint,
        remote_alias_endpoint: remoteMcpAliasEndpoint,
        remote_transport: {
          type: 'streamable-http',
          protocol_version: '2025-11-25',
          stateless: true,
          authentication_required: false,
        },
        command: 'npm',
        args: [
          'exec',
          '--package',
          workpaperPackageSpec,
          '--',
          'bilig-workpaper-mcp',
          '--workpaper',
          './pricing.workpaper.json',
          '--init-demo-workpaper',
          '--writable',
        ],
        tools: [
          'list_sheets',
          'read_range',
          'read_cell',
          'set_cell_contents',
          'set_cell_contents_and_readback',
          'get_cell_display_value',
          'export_workpaper_document',
          'validate_formula',
        ],
        resources: [
          'bilig://workpaper/manifest',
          'bilig://workpaper/agent-handoff',
          'bilig://workpaper/sheets',
          'bilig://workpaper/current-document',
        ],
        prompts: ['edit_and_verify_workpaper', 'debug_workpaper_formula'],
      },
      capabilities: [
        {
          name: 'agent-start',
          type: 'plain-text-entrypoint',
          url: `${siteRoot}/agent-start.txt`,
          well_known_url: `${siteRoot}/.well-known/agent-start.txt`,
          decision_command: `npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json`,
          challenge_command: `npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json`,
        },
        {
          name: 'agent-adoption-kit',
          type: 'docs-entrypoint',
          docs: `${siteRoot}/agent-adoption-kit.html`,
          skill_install: `npx --yes skills@latest add ${skillDiscoveryRoot} --list`,
          challenge_command: `npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json`,
        },
        {
          name: 'workbook-agent-intent-api',
          type: 'npm-library',
          package: '@bilig/workbook',
          runtime: 'Node.js >=22',
          install: 'npm install @bilig/workbook',
          docs: `${siteRoot}/workbook-agent-intent-api.html`,
          source: `${repositoryUrl}/tree/main/examples/workbook-agent-model`,
        },
        {
          name: 'workpaper-formula-runtime',
          type: 'npm-library',
          package: '@bilig/workpaper',
          runtime: 'Node.js >=22',
          install: 'npm install @bilig/workpaper',
          docs: `${siteRoot}/try-bilig-headless-in-node.html`,
        },
        {
          name: 'file-backed-workpaper-mcp',
          type: 'mcp-stdio-server',
          docs: `${siteRoot}/mcp-workpaper-tool-server.html`,
          server_card: `${siteRoot}/.well-known/mcp/server-card.json`,
          challenge_command: `npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge`,
        },
        {
          name: 'gemini-cli-workpaper-extension',
          type: 'gemini-cli-extension',
          install: repositoryUrl,
          install_command: `gemini extensions install ${repositoryUrl} --ref main`,
          manifest: `${repositoryUrl}/blob/main/gemini-extension.json`,
          context: `${repositoryUrl}/blob/main/gemini-workpaper-context.md`,
          docs: `${siteRoot}/gemini-cli-workpaper-extension.html`,
        },
        {
          name: 'claude-desktop-mcpb',
          type: 'mcpb-desktop-extension',
          download_url: mcpbReleaseAssetUrl,
          checksum_url: mcpbReleaseChecksumUrl,
          docs: `${siteRoot}/claude-desktop-mcpb-workpaper.html`,
        },
        {
          name: 'remote-workpaper-mcp-demo',
          type: 'mcp-streamable-http-server',
          endpoint: remoteMcpEndpoint,
          alias_endpoint: remoteMcpAliasEndpoint,
          server_card: remoteMcpServerCard,
          protocol_version: '2025-11-25',
          authentication_required: false,
          docs: `${siteRoot}/mcp-workpaper-tool-server.html#remote-stateless-endpoint`,
        },
        {
          name: 'fastmcp-workpaper-client',
          type: 'python-mcp-client-smoke-test',
          endpoint: remoteMcpEndpoint,
          client: 'FastMCP',
          docs: `${siteRoot}/fastmcp-workpaper-client.html`,
          source: `${repositoryUrl}/tree/main/examples/fastmcp-workpaper-client`,
        },
        {
          name: 'agno-workpaper-mcp',
          type: 'python-agent-mcp-smoke-test',
          framework: 'Agno',
          command:
            'uv run --python 3.12 --with agno --with mcp --with openai python examples/agno-workpaper-mcp/agno_workpaper_mcp.py --output .tmp/agno-workpaper-proof.json',
          docs: `${siteRoot}/agno-workpaper-mcp.html`,
          source: `${repositoryUrl}/tree/main/examples/agno-workpaper-mcp`,
        },
        {
          name: 'pydantic-ai-workpaper-mcp',
          type: 'python-agent-mcp-smoke-test',
          framework: 'Pydantic AI',
          command:
            'uv run --python 3.12 --with pydantic-ai --with mcp --with fastmcp python examples/pydantic-ai-workpaper-mcp/pydantic_ai_workpaper_mcp.py --output .tmp/pydantic-ai-workpaper-proof.json',
          docs: `${siteRoot}/pydantic-ai-workpaper-mcp.html`,
          source: `${repositoryUrl}/tree/main/examples/pydantic-ai-workpaper-mcp`,
        },
        {
          name: 'semantic-kernel-workpaper-mcp',
          type: 'python-agent-mcp-smoke-test',
          framework: 'Microsoft Semantic Kernel',
          command:
            "uv run --python 3.12 --with 'semantic-kernel[mcp]' python examples/semantic-kernel-workpaper-mcp/semantic_kernel_workpaper_mcp.py --output .tmp/semantic-kernel-workpaper-proof.json",
          docs: `${siteRoot}/semantic-kernel-workpaper-mcp.html`,
          source: `${repositoryUrl}/tree/main/examples/semantic-kernel-workpaper-mcp`,
        },
        {
          name: 'smolagents-workpaper-tool',
          type: 'python-agent-tool-smoke-test',
          framework: 'smolagents',
          command: 'uv run --python 3.12 --with smolagents python examples/smolagents-workpaper-tool/smolagents_workpaper_tool.py',
          docs: `${siteRoot}/smolagents-workpaper-tool.html`,
          source: `${repositoryUrl}/tree/main/examples/smolagents-workpaper-tool`,
        },
        {
          name: 'huggingface-workpaper-space-template',
          type: 'gradio-mcp-space-template',
          framework: 'Hugging Face Spaces',
          command:
            "cd examples/huggingface-workpaper-space && npm install --omit=dev --package-lock=false && uv run --python 3.12 --with 'gradio[mcp]>=6.0,<7' python app.py --check",
          docs: `${siteRoot}/huggingface-workpaper-space.html`,
          source: `${repositoryUrl}/tree/main/examples/huggingface-workpaper-space`,
        },
        {
          name: 'inngest-workpaper-step',
          type: 'durable-workflow-step-smoke-test',
          framework: 'Inngest',
          command: 'cd examples/inngest-workpaper-step && pnpm install --ignore-workspace --lockfile=false && pnpm run smoke',
          docs: `${siteRoot}/inngest-workpaper-step.html`,
          source: `${repositoryUrl}/tree/main/examples/inngest-workpaper-step`,
        },
        {
          name: 'airbyte-workpaper-validation',
          type: 'post-sync-validation-smoke-test',
          framework: 'Airbyte',
          command: 'cd examples/airbyte-workpaper-validation && pnpm install --ignore-workspace --lockfile=false && pnpm run smoke',
          docs: `${siteRoot}/airbyte-workpaper-validation.html`,
          source: `${repositoryUrl}/tree/main/examples/airbyte-workpaper-validation`,
        },
        {
          name: 'meltano-workpaper-utility',
          type: 'elt-utility-smoke-test',
          framework: 'Meltano',
          command: 'cd examples/meltano-workpaper-utility && pnpm install --ignore-workspace --lockfile=false && pnpm run smoke',
          docs: `${siteRoot}/meltano-workpaper-utility.html`,
          source: `${repositoryUrl}/tree/main/examples/meltano-workpaper-utility`,
        },
        {
          name: 'formula-clinic',
          type: 'local-cli',
          command: `npm exec --package ${workpaperPackageSpec} -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"`,
          docs: `${siteRoot}/formula-bug-clinic.html`,
        },
      ],
      proof_contract: buildProofContract(),
      verification_contract: [
        'read the relevant range before editing',
        'write the target input or formula cell',
        'read the dependent calculated output after recalculation',
        'export or serialize the WorkPaper document',
        'restore or reimport when a file boundary matters',
        'return editedCell, before, after, afterRestore, persistedDocumentBytes, verified, and limitations',
      ],
      boundaries: {
        good_fit: [
          'pricing, quote approval, budget, payout, import-validation, and forecast logic',
          'agent tools that need deterministic cell addresses and formula readback',
          'service-owned workbook state that can persist as JSON',
        ],
        not_a_fit: agentNotAFitBoundaries,
      },
      public_entrypoints: [
        `${siteRoot}/`,
        `${siteRoot}/llms.txt`,
        `${siteRoot}/llms-full.txt`,
        `${siteRoot}/llms-install.html`,
        `${repositoryUrl}/blob/main/llms-install.md`,
        `${siteRoot}/.well-known/llms.txt`,
        `${siteRoot}/.well-known/llms-full.txt`,
        `${siteRoot}/agent-start.txt`,
        `${siteRoot}/.well-known/agent-start.txt`,
        `${siteRoot}/why-use-bilig.html`,
        `${siteRoot}/eval-xlsx-cache-doctor.html`,
        `${siteRoot}/eval-xlsx-recalc.html`,
        `${siteRoot}/xlsx-cache-doctor-github-action.html`,
        `${siteRoot}/agent-adoption-kit.html`,
        `${siteRoot}/headless-workpaper-agent-handbook.html`,
        `${siteRoot}/agent-workbook-challenge.html`,
        `${siteRoot}/mcp-workpaper-tool-server.html`,
        `${siteRoot}/open-webui-workpaper-mcp.html`,
        `${siteRoot}/open-multi-agent-workpaper-mcp.html`,
        `${siteRoot}/lobehub-workpaper-mcp.html`,
        `${siteRoot}/anythingllm-workpaper-mcp.html`,
        `${siteRoot}/sim-workpaper-mcp.html`,
        `${siteRoot}/fastmcp-workpaper-client.html`,
        `${siteRoot}/agno-workpaper-mcp.html`,
        `${siteRoot}/pydantic-ai-workpaper-mcp.html`,
        `${siteRoot}/smolagents-workpaper-tool.html`,
        `${siteRoot}/huggingface-workpaper-space.html`,
        remoteMcpEndpoint,
        remoteMcpServerCard,
        `${siteRoot}/agent-workpaper-tool-calling-recipe.html`,
        `${siteRoot}/agent-framework-workbook-tools.html`,
        `${siteRoot}/workbook-agent-intent-api.html`,
        `${siteRoot}/openai-agents-sdk-workpaper-tool.html`,
        `${siteRoot}/langgraph-workpaper-toolnode-spreadsheet.html`,
        `${siteRoot}/llamaindex-workpaper-spreadsheet-tool.html`,
        `${siteRoot}/crewai-workpaper-spreadsheet-tool.html`,
        `${siteRoot}/cloudflare-agents-workpaper-spreadsheet-tool.html`,
        `${siteRoot}/semantic-kernel-workpaper-mcp.html`,
        `${siteRoot}/gemini-cli-workpaper-extension.html`,
        `${siteRoot}/node-framework-workpaper-adapters.html`,
        `${siteRoot}/n8n-workpaper-formula-readback.html`,
        `${siteRoot}/dify-workpaper-formula-readback.html`,
        `${siteRoot}/flowise-workpaper-formula-readback.html`,
        `${siteRoot}/pipedream-workpaper-formula-readback.html`,
        `${siteRoot}/triggerdev-workpaper-task.html`,
        `${siteRoot}/inngest-workpaper-step.html`,
        `${siteRoot}/airbyte-workpaper-validation.html`,
        `${siteRoot}/meltano-workpaper-utility.html`,
        `${siteRoot}/temporal-workpaper-activity.html`,
        `${siteRoot}/airflow-workpaper-dag.html`,
        `${siteRoot}/dagster-workpaper-asset.html`,
        `${siteRoot}/kestra-workpaper-flow.html`,
        `${siteRoot}/prefect-workpaper-flow.html`,
        `${siteRoot}/npm-provenance-package-trust.html`,
      ],
    },
    null,
    2,
  )
  const compactPersistedBytesAlias = compactProofContractJsonArrays(json, compactStringArrayProperty)
  const compactPrompts = compactStringArrayProperty(compactPersistedBytesAlias, 'prompts', mcpPromptNames, '    ')
  return `${compactStringArrayProperty(compactPrompts, 'not_a_fit', agentNotAFitBoundaries, '    ')}\n`
}

function compactStringArrayProperty(json: string, propertyName: string, values: readonly string[], indent: string): string {
  const expanded = `${indent}"${propertyName}": [\n${values.map((value) => `${indent}  ${JSON.stringify(value)}`).join(',\n')}\n${indent}]`
  const compact = `${indent}"${propertyName}": [${values.map((value) => JSON.stringify(value)).join(', ')}]`
  return json.replace(expanded, compact)
}

function stripFrontmatter(content: string): string {
  if (!content.startsWith('---\n')) {
    return content.trim()
  }
  return content.replace(/^---\n[\s\S]*?\n---\n+/, '').trim()
}

async function buildLlmsFull(): Promise<string> {
  const sections: string[] = [
    '# Bilig llms-full',
    '',
    '> Full agent context for Bilig, a formula WorkPaper runtime for Node services, MCP clients, and coding-agent workbook tools.',
    '',
    `Repository: ${repositoryUrl}`,
    `Site: ${siteRoot}/`,
    `npm: https://www.npmjs.com/package/@bilig/workpaper`,
    `npm workbook: https://www.npmjs.com/package/@bilig/workbook`,
    `Agent start: ${siteRoot}/agent-start.txt`,
    `Agent instructions: ${siteRoot}/AGENTS.md`,
    `Agent install context: ${siteRoot}/llms-install.html`,
    `Skill manifest: ${skillManifestUrl}`,
    `Compact index: ${siteRoot}/llms.txt`,
    '',
    '## Generated Agent Instructions',
    docsAgentInstructions.trim(),
    '',
    '## Generated Skill Manifest',
    skillDocument.trim(),
  ]

  const sourceSections = await Promise.all(
    llmsFullSources.map(async (source): Promise<string[]> => {
      const content =
        source.relativePath === 'packages/headless/AGENTS.md'
          ? headlessPackageAgentInstructions
          : await readFile(join(repoRoot, source.relativePath), 'utf8')
      return ['', '---', '', `## ${source.title}`, '', `Source: ${source.url}`, '', stripFrontmatter(content)]
    }),
  )

  sourceSections.forEach((section) => sections.push(...section))

  return `${sections.join('\n')}\n`
}

async function generatedTargets(): Promise<ReadonlyArray<readonly [string, string]>> {
  const llmsFull = await buildLlmsFull()
  const llms = await readFile(join(repoRoot, 'docs', 'llms.txt'), 'utf8')
  const llmsInstall = await readFile(join(repoRoot, 'llms-install.md'), 'utf8')
  const agentJson = agentJsonManifest()
  const ideRuleInput = { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec }
  const mcpServerCard = mcpServerCardManifest({
    headlessPackageSpec: workpaperPackageSpec,
    headlessPackageVersion,
    remoteMcpEndpoint,
    repositoryUrl,
    siteRoot,
  })
  return [
    ['docs/AGENTS.md', docsAgentInstructions],
    ['docs/agent-start.txt', docsAgentStart],
    ['docs/agent.json', agentJson],
    ['docs/skill.md', skillDocument],
    ['docs/skill.txt', skillDocument],
    ['docs/llms-install.md', llmsInstall],
    ['docs/llms-full.txt', llmsFull],
    ['docs/.well-known/agent.json', agentJson],
    ['docs/.well-known/agent-start.txt', docsAgentStart],
    ['docs/.well-known/llms.txt', llms],
    ['docs/.well-known/llms-full.txt', llmsFull],
    ['docs/.well-known/agent-skills/index.json', skillIndexJson()],
    ['docs/.well-known/agent-skills/bilig-workpaper/SKILL.md', skillDocument],
    ['docs/.well-known/agent-skills/bilig-workpaper/SKILL.txt', skillDocument],
    ['docs/.well-known/skills/index.json', skillIndexJson()],
    ['docs/.well-known/skills/bilig-workpaper/SKILL.md', skillDocument],
    ['docs/.well-known/skills/bilig-workpaper/SKILL.txt', skillDocument],
    ['docs/.well-known/mcp/server-card.json', mcpServerCard],
    ['docs/.well-known/mcp.json', mcpServerCard],
    ['docs/.well-known/mcp-server-card.json', mcpServerCard],
    ['.cursor/rules/bilig-workpaper.mdc', buildCursorWorkpaperRule(ideRuleInput)],
    ['.windsurf/rules/bilig-workpaper.md', buildWindsurfWorkpaperRule(ideRuleInput)],
    ['.clinerules/bilig-workpaper.md', buildClineWorkpaperRule(ideRuleInput)],
    ['.continue/rules/bilig-workpaper.md', buildContinueWorkpaperRule(ideRuleInput)],
    ['.github/copilot-instructions.md', buildGithubCopilotInstructions(ideRuleInput)],
    ['.github/prompts/bilig-workpaper-proof.prompt.md', buildGithubCopilotWorkpaperPrompt(ideRuleInput)],
    ['.vscode/mcp.json', buildVscodeMcpConfig(ideRuleInput)],
    ['.claude/commands/bilig-workpaper-proof.md', buildClaudeCodeWorkpaperCommand(ideRuleInput)],
    ['.claude/skills/bilig-workpaper/SKILL.md', skillDocument],
    ['skills/bilig-workpaper/SKILL.md', skillDocument],
    ['packages/workpaper/SKILL.md', skillDocument],
    ['packages/workpaper/AGENTS.md', docsAgentInstructions],
    ['packages/headless/SKILL.md', skillDocument],
    ['packages/headless/AGENTS.md', headlessPackageAgentInstructions],
    ['packages/bilig/SKILL.md', workpaperPackageSkillDocument],
    ['packages/bilig/AGENTS.md', workpaperPackageAgentInstructions],
  ] as const
}

const staticReferenceMismatches = await syncVersionedStaticReferences({
  checkOnly,
  headlessPackageSpec,
  headlessPackageVersion,
  mcpbReleaseAssetUrl,
  mcpbReleaseChecksumUrl,
  repoRoot,
  workbookPackageSpec,
  workpaperPackageSpec,
})
const targetResults = await Promise.all(
  (await generatedTargets()).map(async ([relativePath, content]): Promise<string | undefined> => {
    const absolutePath = join(repoRoot, relativePath)
    const existing = await readTextFileIfExists(absolutePath)
    if (existing === content) {
      return undefined
    }

    if (checkOnly) {
      return relativePath
    }

    await mkdir(dirname(absolutePath), { recursive: true })
    await writeFile(absolutePath, content)
    return undefined
  }),
)

const mismatchedTargets = [...staticReferenceMismatches, ...targetResults.filter((target): target is string => target !== undefined)]

if (mismatchedTargets.length > 0) {
  console.error(`Agent discovery docs are stale:\n${mismatchedTargets.map((target) => `- ${target}`).join('\n')}`)
  console.error('Run `pnpm agent:discovery:generate`.')
  process.exitCode = 1
}
