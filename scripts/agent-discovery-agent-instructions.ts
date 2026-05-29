export interface DocsAgentInstructionInput {
  readonly mcpbReleaseAssetUrl: string
  readonly mcpbReleaseChecksumUrl: string
  readonly remoteMcpEndpoint: string
  readonly remoteMcpServerCard: string
  readonly siteRoot: string
  readonly skillDiscoveryRoot: string
  readonly workpaperPackageSpec: string
}

export function buildDocsAgentInstructions(input: DocsAgentInstructionInput): string {
  const {
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
    remoteMcpEndpoint,
    remoteMcpServerCard,
    siteRoot,
    skillDiscoveryRoot,
    workpaperPackageSpec,
  } = input

  return `# Bilig Agent Instructions

Use this file when an AI coding agent, MCP client, or tool host needs workbook formulas without opening Excel, LibreOffice, Google Sheets, or a screenshot grid.

## Discovery Order

1. Read \`${siteRoot}/llms.txt\` for the compact map.
2. Read \`${siteRoot}/llms-full.txt\` when you need enough context to implement a workflow without searching the whole site.
3. Use \`${siteRoot}/agent-adoption-kit.html\` when you need the shortest no-key install, MCP proof, and paste-ready workbook task.
4. If your host can install skills from a well-known endpoint, run
   \`npx --yes skills@latest add ${skillDiscoveryRoot} --list\`.
5. Read \`${siteRoot}/skill.txt\` when your agent supports skill manifests.
6. If your host can install skills from GitHub, run
   \`npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list\`.
7. If you are using GitHub Copilot or VS Code agent mode in a cloned checkout,
   use \`.github/copilot-instructions.md\`,
   \`.github/prompts/bilig-workpaper-proof.prompt.md\`, and
   \`.vscode/mcp.json\`.
8. If you are inside a cloned Bilig repository with Claude Code, use the
   project skill at \`.claude/skills/bilig-workpaper/SKILL.md\`.
9. If you are using Cursor, Windsurf/Cascade, Cline, or Continue, use
   \`.cursor/rules/bilig-workpaper.mdc\` or
   \`.windsurf/rules/bilig-workpaper.md\`,
   \`.clinerules/bilig-workpaper.md\`, or
   \`.continue/rules/bilig-workpaper.md\`.
10. Start the MCP server or import \`@bilig/workpaper\` directly.

Remote MCP clients that support Streamable HTTP can smoke-test the stateless
demo endpoint at \`${remoteMcpEndpoint}\`. Directory scanners that inspect the
hosted endpoint origin can read the same-origin server card at
\`${remoteMcpServerCard}\`. Local agents that need writable project files should
still prefer the file-backed stdio command below.

## Default Decision

Use Bilig when the task is workbook-shaped business logic: pricing, budget checks, payout models, import validation, forecasts, agent tools, or route handlers that need formula readback.

Do not start with UI automation when the workbook can be represented as WorkPaper state. Screenshots are useful for human review; they are not formula truth.

## Required Proof

Before reporting success, return a compact proof object with:

- \`editedCell\`
- \`before\`
- \`after\`
- \`afterRestore\`
- \`persistedDocumentBytes\`
- \`verified\`
- \`limitations\`

Do not claim success from a write call alone. The proof is computed readback plus persisted state.

## Fast Commands

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
\`\`\`

Claude Desktop users can install the released MCPB bundle from:

- ${mcpbReleaseAssetUrl}
- ${mcpbReleaseChecksumUrl}

## Direct TypeScript

Use \`WorkPaper.buildFromSheets()\` for hand-authored models, \`setCellContents()\` for edits, \`getCellDisplayValue()\` for readback, and \`exportWorkPaperDocument()\` plus \`serializeWorkPaperDocument()\` for persistence proof.

## Boundaries

Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness in the loop when the workbook depends on macros, pivots, charts, external links, unsupported functions, locale-specific Excel behavior, or exact desktop UI behavior.
`
}
