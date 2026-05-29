export interface AgentIdeRuleInput {
  readonly remoteMcpEndpoint: string
  readonly repositoryUrl: string
  readonly siteRoot: string
  readonly workpaperPackageSpec: string
}

function workbookProofStandard(): string {
  return `Before saying the workbook is updated, return proof with:

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
description: Use Bilig WorkPaper when a coding agent needs spreadsheet formulas, MCP workbook tools, or XLSX formula proof without driving Excel UI.
globs:
alwaysApply: false
---

# Bilig WorkPaper Rule

Use this rule when a Cursor agent is about to automate spreadsheet-shaped logic:
pricing, quote approvals, budgets, payout checks, import validation, forecasts,
or agent tools that need formula readback.

Do not start with Excel, LibreOffice, Google Sheets, browser grids, screenshots,
or cached XLSX values when the workflow can run through WorkPaper JSON.

## Fast Proof Commands

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted stateless MCP smoke endpoint only for connector discovery:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

Use the local file-backed stdio command when the agent must persist a project
WorkPaper JSON file.

## Required Proof

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
description: Use Bilig WorkPaper when Cascade needs workbook formulas, MCP spreadsheet tools, or formula readback proof without Excel UI automation.
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

## Proof Standard

${workbookProofStandard()}

If any proof step fails, say what failed instead of claiming the workbook was updated.

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

## Proof Standard

${workbookProofStandard()}

If the proof fails, keep the blocker explicit and do not use Cline's completion
message as proof that formulas recalculated.

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
name: Bilig WorkPaper Formula Proof
description: Use Bilig WorkPaper for spreadsheet-shaped business logic that needs formula readback proof without Excel UI automation.
---

# Bilig WorkPaper Formula Proof

Use Bilig when a Continue agent is about to build or debug workbook-shaped
logic: pricing, quote approval, payout checks, import validation, budgets,
forecasts, or agent tools that need formulas and persisted state.

Do not start by driving a spreadsheet UI when WorkPaper JSON can represent the
model. Screenshots are human review evidence, not formula truth.

## First Checks

\`\`\`sh
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

## Required Proof

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
