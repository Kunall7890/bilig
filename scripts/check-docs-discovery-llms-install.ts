import { requireIncludes } from './check-docs-discovery-core.ts'

export interface LlmsInstallDiscoveryInput {
  readonly parsedAgentJson: object
  readonly docsLlmsInstall: string
  readonly llmsFull: string
  readonly llmsInstall: string
}

export function requireLlmsInstallDiscovery({ docsLlmsInstall, llmsFull, llmsInstall, parsedAgentJson }: LlmsInstallDiscoveryInput): void {
  requireIncludes(llmsFull, '## Agent Install Context', 'docs/llms-full.txt')
  if (docsLlmsInstall !== llmsInstall) {
    throw new Error('docs/llms-install.md must match root llms-install.md')
  }
  requireIncludes(llmsInstall, '# Bilig WorkPaper install context', 'llms-install.md')
  requireIncludes(
    llmsInstall,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'llms-install.md',
  )
  requireIncludes(llmsInstall, 'CLAUDE.md', 'llms-install.md')
  requireIncludes(llmsInstall, 'CONVENTIONS.md', 'llms-install.md')
  requireIncludes(llmsInstall, '.aider.conf.yml', 'llms-install.md')
  requireIncludes(llmsInstall, '.clinerules/bilig-workpaper.md', 'llms-install.md')
  requireIncludes(llmsInstall, '.cursor/rules/bilig-workpaper.mdc', 'llms-install.md')
  requireIncludes(llmsInstall, '.devin/rules/bilig-workpaper.md', 'llms-install.md')
  requireIncludes(llmsInstall, '.kiro/settings/mcp.json', 'llms-install.md')
  requireIncludes(llmsInstall, '.kiro/steering/bilig-workpaper.md', 'llms-install.md')
  requireIncludes(llmsInstall, '.trae/mcp.json', 'llms-install.md')
  requireIncludes(llmsInstall, '.trae/rules/bilig-workpaper.md', 'llms-install.md')
  requireIncludes(llmsInstall, '.zed/settings.json', 'llms-install.md')
  requireIncludes(llmsInstall, '.junie/mcp/mcp.json', 'llms-install.md')
  requireIncludes(llmsInstall, 'opencode.jsonc', 'llms-install.md')
  requireIncludes(llmsInstall, '.opencode/agents/bilig-workpaper.md', 'llms-install.md')
  requireIncludes(llmsInstall, 'https://proompteng.github.io/bilig/agent-rule-chooser.html', 'llms-install.md')
  requireIncludes(llmsInstall, 'set_cell_contents_and_readback', 'llms-install.md')
  requireIncludes(llmsInstall, 'bilig-workpaper-mcp --from-xlsx ./pricing.xlsx', 'llms-install.md')
  requireIncludes(llmsInstall, 'analyze_workbook_risk', 'llms-install.md')
  requireIncludes(llmsInstall, 'It does not certify\nExcel compatibility.', 'llms-install.md')
  requireIncludes(llmsInstall, 'Reject answers that only say a cell was written.', 'llms-install.md')
  if (Reflect.get(parsedAgentJson, 'llms_install') !== 'https://proompteng.github.io/bilig/llms-install.html') {
    throw new Error('docs/.well-known/agent.json must advertise the agent install context page')
  }
  if (Reflect.get(parsedAgentJson, 'llms_install_source') !== 'https://github.com/proompteng/bilig/blob/main/llms-install.md') {
    throw new Error('docs/.well-known/agent.json must advertise the agent install context source')
  }
}
