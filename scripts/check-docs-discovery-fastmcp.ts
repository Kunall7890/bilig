import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

export interface FastMcpWorkpaperClientDiscoveryContext {
  readonly repoRoot: string
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly scopedWorkpaperPackageReadme: string
}

export async function requireFastMcpWorkpaperClientDiscovery(context: FastMcpWorkpaperClientDiscoveryContext): Promise<void> {
  const fastmcpWorkpaperClient = await readFile(join(context.docsRoot, 'fastmcp-workpaper-client.md'), 'utf8')

  await Promise.all(
    ['README.md', 'fastmcp_workpaper_client.py', 'scripts/check-client.py'].map((sourceFile) =>
      requireFile(join(context.repoRoot, 'examples', 'fastmcp-workpaper-client', sourceFile)),
    ),
  )

  requireIncludes(context.index, './fastmcp-workpaper-client.html', 'docs/index.html')
  requireIncludes(fastmcpWorkpaperClient, 'examples/fastmcp-workpaper-client', 'docs/fastmcp-workpaper-client.md')
  requireIncludes(fastmcpWorkpaperClient, 'from fastmcp import Client', 'docs/fastmcp-workpaper-client.md')
  requireIncludes(
    fastmcpWorkpaperClient,
    'FastMCP owns the MCP client session. Bilig owns the WorkPaper formula tools',
    'docs/fastmcp-workpaper-client.md',
  )
  requireIncludes(fastmcpWorkpaperClient, 'The hosted endpoint is stateless.', 'docs/fastmcp-workpaper-client.md')
  requireIncludes(fastmcpWorkpaperClient, 'https://gofastmcp.com/community/showcase', 'docs/fastmcp-workpaper-client.md')
  requireIncludes(context.llms, 'https://proompteng.github.io/bilig/fastmcp-workpaper-client.html', 'docs/llms.txt')
  requireIncludes(context.llms, 'https://github.com/proompteng/bilig/tree/main/examples/fastmcp-workpaper-client', 'docs/llms.txt')
  requireIncludes(context.llmsFull, 'FastMCP WorkPaper Client', 'docs/llms-full.txt')
  requireIncludes(context.scopedWorkpaperPackageReadme, 'FastMCP Python client for hosted MCP smoke tests', 'packages/workpaper/README.md')
  requireIncludes(
    context.scopedWorkpaperPackageReadme,
    'https://proompteng.github.io/bilig/fastmcp-workpaper-client.html',
    'packages/workpaper/README.md',
  )
}
