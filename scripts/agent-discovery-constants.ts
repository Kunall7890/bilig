export const skillTags = ['ai-agents', 'spreadsheet-automation', 'formulas', 'xlsx', 'mcp', 'typescript'] as const

export const mcpPromptNames = ['edit_and_verify_workpaper', 'debug_workpaper_formula'] as const

export const agentNotAFitBoundaries = [
  'manual spreadsheet editing as the main product',
  'Office macros or desktop Excel automation',
  'one-off arithmetic',
] as const

export const versionedStaticReferenceRoots = [
  'README.md',
  'packages/headless/README.md',
  'packages/headless/AGENTS.md',
  'packages/headless/SKILL.md',
  'packages/bilig/README.md',
  'packages/bilig/AGENTS.md',
  'packages/bilig/SKILL.md',
  'skills/bilig-workpaper/SKILL.md',
  'docs/agent-workbook-challenge.md',
  'docs/agent-xlsx-formula-recalculation-without-libreoffice.md',
  'docs/claude-desktop-mcpb-workpaper.md',
  'docs/formula-bug-clinic.md',
  'docs/formula-workbooks-node-services-agent-tools.md',
  'docs/ai-agent-spreadsheet-tool-node.md',
  'docs/agent-framework-workbook-tools.md',
  'docs/headless-workpaper-agent-handbook.md',
  'docs/index.html',
  'docs/llms.txt',
  'docs/mcp-client-setup.md',
  'docs/windmill-workpaper-script.md',
  'docs/triggerdev-workpaper-task.md',
  'docs/temporal-workpaper-activity.md',
  'docs/airflow-workpaper-dag.md',
  'docs/dagster-workpaper-asset.md',
  'docs/kestra-workpaper-flow.md',
  'docs/prefect-workpaper-flow.md',
  'docs/open-webui-workpaper-mcp.md',
  'docs/lobehub-workpaper-mcp.md',
  'docs/anythingllm-workpaper-mcp.md',
  'docs/fastmcp-workpaper-client.md',
  'docs/smolagents-workpaper-tool.md',
  'docs/sim-workpaper-mcp.md',
  'docs/mcp-spreadsheet-server-directory.md',
  'docs/mcp-workpaper-tool-server.md',
  'docs/spreadsheet-mcp-server-comparison.md',
  'docs/why-agents-need-workbook-apis.md',
] as const

export const versionedStaticReferenceExtensions = new Set(['.html', '.json', '.md', '.txt'])
