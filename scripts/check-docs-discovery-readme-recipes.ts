import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

export function requireReadmeAgentWorkflowRecipeLinks(readme: string): void {
  const sectionStart = readme.indexOf('## Agent And Workflow Recipes')
  const sectionEnd = readme.indexOf('### Stale XLSX Formula Values? Run This First')
  const section = sectionStart === -1 || sectionEnd === -1 ? '' : readme.slice(sectionStart, sectionEnd)

  requireNotIncludes(section, '](docs/', 'README.md Agent And Workflow Recipes section')
  requireIncludes(
    section,
    'https://proompteng.github.io/bilig/open-webui-workpaper-mcp.html',
    'README.md Agent And Workflow Recipes section',
  )
  requireIncludes(
    section,
    'https://proompteng.github.io/bilig/pipedream-workpaper-formula-readback.html',
    'README.md Agent And Workflow Recipes section',
  )
  requireIncludes(
    section,
    'https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html',
    'README.md Agent And Workflow Recipes section',
  )
}
