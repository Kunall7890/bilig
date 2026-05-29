export const agentFrameworkLlmsRequiredLinks = [
  'https://proompteng.github.io/bilig/mastra-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/llamaindex-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html',
  'https://proompteng.github.io/bilig/open-multi-agent-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/copilotkit-workpaper-spreadsheet-action.html',
  'https://proompteng.github.io/bilig/cloudflare-agents-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/crewai-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/semantic-kernel-workpaper-mcp.html',
] as const

export const agentFrameworkDocRequirements = [
  {
    path: 'docs/mastra-workpaper-spreadsheet-tool.md',
    includes: [
      'Mastra WorkPaper spreadsheet tool',
      '@bilig/workpaper',
      'createTool',
      'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
    ],
  },
  {
    path: 'docs/llamaindex-workpaper-spreadsheet-tool.md',
    includes: [
      'LlamaIndex.TS WorkPaper spreadsheet tool',
      '@bilig/workpaper',
      'tool(fn, { parameters })',
      'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
    ],
  },
  {
    path: 'docs/langgraph-workpaper-toolnode-spreadsheet.md',
    includes: [
      'LangGraph.js WorkPaper ToolNode spreadsheet tool',
      '@langchain/langgraph',
      '@langchain/mcp-adapters',
      'new ToolNode(tools)',
      'examples/langgraph-workpaper-tool-state',
      'examples/langchain-mcp-workpaper-toolnode',
      'pnpm run smoke',
    ],
  },
  {
    path: 'docs/open-multi-agent-workpaper-mcp.md',
    includes: ['Open Multi-Agent WorkPaper MCP example', 'connectMCPTools()', '@bilig/workpaper@0.96.0', 'verified'],
  },
  {
    path: 'docs/copilotkit-workpaper-spreadsheet-action.md',
    includes: [
      'CopilotKit WorkPaper spreadsheet action',
      '@bilig/workpaper',
      'useCopilotAction',
      'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
    ],
  },
  {
    path: 'docs/cloudflare-agents-workpaper-spreadsheet-tool.md',
    includes: [
      'Cloudflare Agents WorkPaper spreadsheet tool',
      '@bilig/workpaper',
      'agentTool',
      'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
    ],
  },
  {
    path: 'docs/crewai-workpaper-spreadsheet-tool.md',
    includes: [
      'CrewAI WorkPaper spreadsheet tool',
      '@bilig/workpaper',
      'JSON contract',
      'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
    ],
  },
  {
    path: 'docs/semantic-kernel-workpaper-mcp.md',
    includes: ['Semantic Kernel WorkPaper MCP Plugin', 'MCPStdioPlugin', 'examples/semantic-kernel-workpaper-mcp', 'verified'],
  },
] as const
