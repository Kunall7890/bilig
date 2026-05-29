import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

export function requireReadmeAgentWorkflowRecipeLinks(readme: string): void {
  const context = 'README.md Integration Recipes After The Proof section'
  const sectionStart = readme.indexOf('## Integration Recipes After The Proof')
  const sectionEnd = readme.indexOf('## Choose An Evaluation Path')
  const section = sectionStart === -1 || sectionEnd === -1 ? '' : readme.slice(sectionStart, sectionEnd)

  requireNotIncludes(section, '](docs/', context)
  requireIncludes(section, 'https://proompteng.github.io/bilig/open-webui-workpaper-mcp.html', context)
  requireIncludes(section, 'https://proompteng.github.io/bilig/pipedream-workpaper-formula-readback.html', context)
  requireIncludes(section, 'https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html', context)
}
