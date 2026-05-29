export interface AgentIdeRuleInput {
  readonly remoteMcpEndpoint: string
  readonly repositoryUrl: string
  readonly siteRoot: string
  readonly workpaperPackageSpec: string
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

After any write, return proof with:

- edited sheet and A1 cell;
- input and dependent-output values before the change;
- recalculated output after the change;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not report success from a write call alone.

## Safe Command Handling

Do not concatenate user-provided paths or cell references into shell strings.
Use argument arrays where the host supports them. Reject workbook paths or cell
arguments containing newlines, backticks, \`$(\`, \`;\`, \`&\`, \`|\`, \`<\`, or \`>\`.

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

Before saying the workbook is updated, provide:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized/exported WorkPaper evidence;
- restore or restart readback when persistence matters;
- limitations for unsupported formulas or Excel-only behavior.

If any proof step fails, say what failed instead of claiming the workbook was
updated.

## Command Safety

Do not build shell commands by concatenating user text. Prefer MCP \`command\`
plus \`args\` arrays or direct TypeScript calls. Reject path or cell arguments
containing newlines, backticks, \`$(\`, \`;\`, \`&\`, \`|\`, \`<\`, or \`>\`.

## References

- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}
