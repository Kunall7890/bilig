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
  requireIncludes(llmsInstall, '.clinerules/bilig-workpaper.md', 'llms-install.md')
  requireIncludes(llmsInstall, '.cursor/rules/bilig-workpaper.mdc', 'llms-install.md')
  requireIncludes(llmsInstall, '.devin/rules/bilig-workpaper.md', 'llms-install.md')
  requireIncludes(llmsInstall, 'https://proompteng.github.io/bilig/agent-rule-chooser.html', 'llms-install.md')
  requireIncludes(llmsInstall, 'set_cell_contents_and_readback', 'llms-install.md')
  requireIncludes(llmsInstall, 'Reject answers that only say a cell was written.', 'llms-install.md')
  if (Reflect.get(parsedAgentJson, 'llms_install') !== 'https://proompteng.github.io/bilig/llms-install.html') {
    throw new Error('docs/.well-known/agent.json must advertise the agent install context page')
  }
  if (Reflect.get(parsedAgentJson, 'llms_install_source') !== 'https://github.com/proompteng/bilig/blob/main/llms-install.md') {
    throw new Error('docs/.well-known/agent.json must advertise the agent install context source')
  }
}
