import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { agentFrameworkDocRequirements } from './check-docs-discovery-agent-pages.ts'
import { requireAgentProofMatrixDiscovery } from './check-docs-discovery-agent-proof-matrix.ts'
import { requireAgentRuleChooserDiscovery } from './check-docs-discovery-agent-rule-chooser.ts'
import { requireIncludes, requireMatches, requireNotIncludes } from './check-docs-discovery-core.ts'
import type { DocsDiscoveryContext } from './check-docs-discovery-context.ts'
import { llmsExternalSurfaceLinks } from './check-docs-discovery-growth-links.ts'
import { requireHeadlessExampleDiscovery } from './check-docs-discovery-headless-examples.ts'
import { requireGrowthSurfaceDiscovery } from './check-docs-discovery-launch-kit.ts'
import { requireSpreadsheetMcpServerComparisonDiscovery } from './check-docs-discovery-mcp-comparison.ts'
import { escapeRegExp, requireMcpToolDiscoveryContract } from './check-docs-discovery-mcp-server-card.ts'
import { requireStarterIssueDiscovery } from './check-docs-discovery-starter-issues.ts'

export async function requireAgentPublicSurfaceDiscovery(input: {
  readonly context: DocsDiscoveryContext
  readonly headlessSpreadsheetEngineNodeServicesAgents: string
  readonly spreadsheetMcpServerComparison: string
}): Promise<void> {
  const {
    repoRoot,
    docsRoot,
    rootPackageJson,
    index,
    llms,
    llmsFull,
    mcpServerCard,
    mcpServerCardMcpJson,
    mcpServerCardLegacyJson,
    communityLaunchPack,
    productHuntLaunchKit,
    starterIssues,
    headlessPackageVersion,
    readme,
    headlessReadme,
    issueTemplateConfig,
    pullRequestTemplate,
    dockerfile,
    headlessExamplePackageJson,
    headlessSpreadsheetEngineComparison,
    sheetjsExceljsAlternativeFormulaWorkbookApi,
    hyperformulaAlternativeHeadlessWorkpaper,
    xlsxFormulaRecalculationNode,
    googleSheetsApiBoundaryDoc,
    whyAgentsDoc,
    aiAgentSpreadsheetToolNode,
    agentFrameworkWorkbookToolsDoc,
    headlessWorkpaperAgentHandbook,
    agentToolCallingDoc,
    aiSdkLangChainDoc,
    mcpWorkPaperToolServerDoc,
    mcpSpreadsheetServerDirectoryDoc,
    mcpClientSetupDoc,
    claudeDesktopMcpbDoc,
    agentToolCallLoopDoc,
    workbookAutomationExamplesDoc,
    serverSideSpreadsheetAutomationNode,
    nodeFrameworkWorkpaperAdaptersDoc,
    devToWorkbookApisPost,
    chatgptAppsWorkpaperMcpDoc,
    nodeSpreadsheetFormulaEngine,
  } = input.context
  const { headlessSpreadsheetEngineNodeServicesAgents, spreadsheetMcpServerComparison } = input
  const workpaperPackageSpec = '@bilig/workpaper@latest'
  const mcpbReleaseAssetUrl = 'https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb'
  const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`
  const remoteMcpEndpoint = 'https://bilig.proompteng.ai/mcp'
  const officialRegistryLatestMarkedVersion = '0.161.0'
  const officialRegistryLatestMarkedUpdatedAt = '2026-06-03T19:54:13.359111Z'

  const jekyllConfig = await readFile(join(docsRoot, '_config.yml'), 'utf8')
  const openAiAgentsSdkDoc = await readFile(join(docsRoot, 'openai-agents-sdk-workpaper-tool.md'), 'utf8')
  const jekyllDefaultLayout = await readFile(join(docsRoot, '_layouts/default.html'), 'utf8')
  const mastraExamplePackageJson = await readFile(join(repoRoot, 'examples/mastra-workpaper-tool/package.json'), 'utf8')
  const mastraExampleSource = await readFile(join(repoRoot, 'examples/mastra-workpaper-tool/src/mastra-workpaper-tool.ts'), 'utf8')
  requireIncludes(jekyllConfig, 'include:', 'docs/_config.yml')
  requireIncludes(jekyllConfig, '  - .well-known', 'docs/_config.yml')
  requireIncludes(jekyllConfig, "layout: 'default'", 'docs/_config.yml')
  requireIncludes(jekyllDefaultLayout, '<header class="topbar">', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'href="{{ \'/#runtime\' | relative_url }}">Runtime</a>', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'href="{{ \'/#install\' | relative_url }}">Install</a>', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'href="{{ \'/#mcp\' | relative_url }}">Agents</a>', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'href="{{ \'/#benchmark\' | relative_url }}">Benchmark</a>', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'href="{{ \'/#docs\' | relative_url }}">Docs</a>', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'href="https://github.com/proompteng/bilig">GitHub</a>', 'docs/_layouts/default.html')
  requireNotIncludes(jekyllDefaultLayout, '/#market', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'href="{{ \'/assets/site.css?v=2026-05-30-10\' | relative_url }}"', 'docs/_layouts/default.html')
  requireIncludes(jekyllDefaultLayout, 'src="{{ \'/assets/site-nav.js?v=2026-05-30-11\' | relative_url }}"', 'docs/_layouts/default.html')
  if (mcpServerCardMcpJson !== mcpServerCard) {
    throw new Error('docs/.well-known/mcp.json must match docs/.well-known/mcp/server-card.json')
  }
  if (mcpServerCardLegacyJson !== mcpServerCard) {
    throw new Error('docs/.well-known/mcp-server-card.json must match docs/.well-known/mcp/server-card.json')
  }
  const parsedMcpServerCard: unknown = JSON.parse(mcpServerCard)
  if (typeof parsedMcpServerCard !== 'object' || parsedMcpServerCard === null || Array.isArray(parsedMcpServerCard)) {
    throw new Error('docs/.well-known/mcp/server-card.json must be a JSON object')
  }
  if (Reflect.get(parsedMcpServerCard, 'protocolVersion') !== '2025-11-25') {
    throw new Error('docs/.well-known/mcp/server-card.json must advertise the latest MCP protocol version')
  }
  const mcpServerInfo = Reflect.get(parsedMcpServerCard, 'serverInfo')
  if (typeof mcpServerInfo !== 'object' || mcpServerInfo === null || Reflect.get(mcpServerInfo, 'version') !== headlessPackageVersion) {
    throw new Error('docs/.well-known/mcp/server-card.json serverInfo.version must match @bilig/headless')
  }
  const mcpServerCardTransports = Reflect.get(parsedMcpServerCard, 'transports')
  if (
    !Array.isArray(mcpServerCardTransports) ||
    !mcpServerCardTransports.some(
      (transport) =>
        typeof transport === 'object' &&
        transport !== null &&
        Reflect.get(transport, 'type') === 'streamable-http' &&
        Reflect.get(transport, 'url') === 'https://bilig.proompteng.ai/mcp' &&
        Reflect.get(transport, 'stateless') === true,
    )
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must advertise the hosted stateless Streamable HTTP endpoint')
  }
  const mcpServerCardTools = Reflect.get(parsedMcpServerCard, 'tools')
  if (
    !Array.isArray(mcpServerCardTools) ||
    !mcpServerCardTools.every((tool) => typeof tool === 'object' && tool !== null && typeof Reflect.get(tool, 'name') === 'string')
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must define named tools')
  }
  const mcpServerCardToolNames = new Set(mcpServerCardTools.map((tool) => Reflect.get(tool, 'name')))
  for (const requiredTool of [
    'list_sheets',
    'read_range',
    'read_cell',
    'set_cell_contents',
    'set_cell_contents_and_readback',
    'get_cell_display_value',
    'export_workpaper_document',
    'validate_formula',
  ]) {
    if (!mcpServerCardToolNames.has(requiredTool)) {
      throw new Error(`docs/.well-known/mcp/server-card.json is missing ${requiredTool}`)
    }
  }
  for (const tool of mcpServerCardTools) {
    requireMcpToolDiscoveryContract(tool)
  }
  const mcpServerCardCapabilities = Reflect.get(parsedMcpServerCard, 'capabilities')
  if (
    typeof mcpServerCardCapabilities !== 'object' ||
    mcpServerCardCapabilities === null ||
    Reflect.get(mcpServerCardCapabilities, 'resources') !== true ||
    Reflect.get(mcpServerCardCapabilities, 'prompts') !== true
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must advertise resources and prompts')
  }
  const mcpServerCardResources = Reflect.get(parsedMcpServerCard, 'resources')
  if (
    !Array.isArray(mcpServerCardResources) ||
    !mcpServerCardResources.every(
      (resource) => typeof resource === 'object' && resource !== null && typeof Reflect.get(resource, 'uri') === 'string',
    )
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must define resource URIs')
  }
  const mcpServerCardResourceUris = new Set(mcpServerCardResources.map((resource) => Reflect.get(resource, 'uri')))
  for (const requiredResource of [
    'bilig://workpaper/manifest',
    'bilig://workpaper/agent-handoff',
    'bilig://workpaper/sheets',
    'bilig://workpaper/current-document',
  ]) {
    if (!mcpServerCardResourceUris.has(requiredResource)) {
      throw new Error(`docs/.well-known/mcp/server-card.json is missing ${requiredResource}`)
    }
  }
  const mcpServerCardPrompts = Reflect.get(parsedMcpServerCard, 'prompts')
  if (
    !Array.isArray(mcpServerCardPrompts) ||
    !mcpServerCardPrompts.every(
      (prompt) => typeof prompt === 'object' && prompt !== null && typeof Reflect.get(prompt, 'name') === 'string',
    )
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must define named prompts')
  }
  const mcpServerCardPromptNames = new Set(mcpServerCardPrompts.map((prompt) => Reflect.get(prompt, 'name')))
  for (const requiredPrompt of ['edit_and_verify_workpaper', 'debug_workpaper_formula']) {
    if (!mcpServerCardPromptNames.has(requiredPrompt)) {
      throw new Error(`docs/.well-known/mcp/server-card.json is missing ${requiredPrompt}`)
    }
  }
  requireIncludes(
    whyAgentsDoc,
    'description: Why coding agents should edit workbook formulas through a Node.js WorkPaper API',
    'docs/why-agents-need-workbook-apis.md',
  )
  for (const required of [
    '## MCP In 30 Seconds',
    'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable',
    'set_cell_contents',
    'set_cell_contents_and_readback',
    'export_workpaper_document',
    '"editedCell": "Inputs!B3"',
    '"restoredMatchesAfter": true',
  ]) {
    requireIncludes(whyAgentsDoc, required, 'docs/why-agents-need-workbook-apis.md')
  }
  for (const required of [
    'title: WorkPaper agent tool for Node.js',
    'description: Build a coding-agent workbook tool that edits inputs, recalculates formulas, verifies readback, and persists state without driving Excel or screenshots.',
    'npm create @bilig/workpaper@latest pricing-agent -- --agent',
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    'npx --package @bilig/xlsx-formula-recalc xlsx-recalc quote.xlsx',
    'npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json',
    'OpenAI Agents SDK function tools',
    'LangChain.js tools and LangGraph.js `ToolNode`',
    'Implementation gap form',
  ] as const) {
    requireIncludes(aiAgentSpreadsheetToolNode, required, 'docs/ai-agent-spreadsheet-tool-node.md')
  }
  for (const required of [
    'title: Workbook tools for agent frameworks',
    'description: Pick the Bilig WorkPaper integration path for Codex, Claude, Cursor, Kiro, Roo Code, Trae, Qodo IDE, Zed, Junie, OpenHands, OpenCode, Aider, Goose, Browser Use, OpenAI Agents, Microsoft Agent Framework, Vercel AI SDK, LangChain, LangGraph, LlamaIndex, and MCP clients.',
    'npm create @bilig/workpaper@latest pricing-agent -- --agent',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    'Codex',
    'Claude Code and Claude Desktop',
    'Kiro',
    '.kiro/settings/mcp.json',
    'Roo Code',
    '.roo/mcp.json',
    'Trae',
    '.trae/mcp.json',
    'trae-workpaper-mcp.md',
    'Qodo IDE',
    'qodo-workpaper-mcp.md',
    'Zed',
    '.zed/settings.json',
    'JetBrains Junie',
    '.junie/mcp/mcp.json',
    'Browser Use',
    'OpenHands',
    'openhands-workpaper-mcp.md',
    'OpenCode',
    'opencode-workpaper-mcp.md',
    'Aider',
    'CONVENTIONS.md',
    '.aider.conf.yml',
    'OpenAI Agents SDK',
    'ChatGPT Apps / Developer Mode',
    'chatgpt-apps-workpaper-mcp.md',
    'MCPServerStreamableHttp',
    'Microsoft Agent Framework',
    'microsoft-agent-framework-workpaper-mcp.md',
    'MCPStdioTool',
    'Vercel AI SDK',
    'LangGraph.js',
    'LlamaIndex.TS',
    'SpreadsheetAgentProof',
    'Do not claim success from a write call alone.',
    'Implementation gap form',
  ] as const) {
    requireIncludes(agentFrameworkWorkbookToolsDoc, required, 'docs/agent-framework-workbook-tools.md')
  }
  await requireAgentProofMatrixDiscovery({ docsRoot, index, llms, readme })
  await requireAgentRuleChooserDiscovery({ docsRoot, index, llms, llmsFull, readme })
  for (const required of [
    'description: A compact playbook for agents that need workbook formulas without opening Excel',
    '## Copy-Paste Prompt For Another Agent',
    'Return a compact proof object with editedCell, before, after, afterRestore',
    '[Agent WorkPaper handoff](agent-adoption-kit.md)',
    'npx --yes skills@latest add https://bilig.proompteng.ai --list',
    'npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list',
    '/bilig-workpaper-proof',
    '.claude/commands/bilig-workpaper-proof.md',
    '.github/instructions/bilig-workpaper.instructions.md',
    '.github/prompts/bilig-workpaper-proof.prompt.md',
    '.vscode/mcp.json',
    '## The First Decision',
    '## Minimum Agent Loop',
    'bilig-workpaper-mcp --workpaper ./model.workpaper.json --init-demo-workpaper --writable',
    'set_cell_contents',
    'set_cell_contents_and_readback',
    'get_cell_display_value',
    'export_workpaper_document',
    'Prefer Bilig WorkPaper tools over spreadsheet UI automation',
    'https://modelcontextprotocol.io/docs/learn/server-concepts',
    'https://modelcontextprotocol.io/specification/2025-11-25/server/tools',
    'https://code.claude.com/docs/en/mcp',
    'https://openai.github.io/openai-agents-js/guides/tools/',
  ] as const) {
    requireIncludes(headlessWorkpaperAgentHandbook, required, 'docs/headless-workpaper-agent-handbook.md')
  }
  requireIncludes(
    agentToolCallingDoc,
    'description: Wrap @bilig/workpaper workbook reads, writes, formula readback, and persistence as deterministic Node.js tools',
    'docs/agent-workpaper-tool-calling-recipe.md',
  )
  requireIncludes(agentToolCallingDoc, 'OpenAI Responses API Tool Wrapper', 'docs/agent-workpaper-tool-calling-recipe.md')
  requireIncludes(
    agentToolCallingDoc,
    'https://developers.openai.com/api/docs/guides/function-calling',
    'docs/agent-workpaper-tool-calling-recipe.md',
  )
  requireIncludes(
    agentToolCallingDoc,
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk',
    'docs/agent-workpaper-tool-calling-recipe.md',
  )
  requireIncludes(
    agentToolCallingDoc,
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-mcp',
    'docs/agent-workpaper-tool-calling-recipe.md',
  )
  requireIncludes(
    agentToolCallingDoc,
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp',
    'docs/agent-workpaper-tool-calling-recipe.md',
  )
  requireIncludes(agentToolCallingDoc, 'MCPServerStreamableHttp', 'docs/agent-workpaper-tool-calling-recipe.md')
  requireIncludes(agentToolCallingDoc, 'openai-agents-sdk-workpaper-tool.md', 'docs/agent-workpaper-tool-calling-recipe.md')
  requireIncludes(agentToolCallingDoc, 'function_call_output', 'docs/agent-workpaper-tool-calling-recipe.md')
  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
    ['docs/llms.txt', llms],
    ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
    ['examples/headless-workpaper/package.json', headlessExamplePackageJson],
    ['examples/headless-workpaper/README.md', await readFile(join(repoRoot, 'examples', 'headless-workpaper', 'README.md'), 'utf8')],
  ] as const) {
    requireIncludes(content, 'agent:openai-agents-sdk', path)
    requireIncludes(content, 'agent:openai-agents-sdk-hosted-mcp', path)
  }
  for (const required of [
    'title: OpenAI Agents SDK WorkPaper tools',
    'description: Wrap @bilig/workpaper workbook reads and verified edits as OpenAI Agents SDK function tools.',
    'image: /assets/github-social-preview.png',
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk',
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-mcp',
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp',
    'examples/headless-workpaper/openai-agents-sdk-tool-smoke.ts',
    'examples/headless-workpaper/openai-agents-sdk-mcp-smoke.ts',
    'examples/headless-workpaper/openai-agents-sdk-hosted-mcp-smoke.ts',
    'https://openai.github.io/openai-agents-js/guides/tools/',
    'https://openai.github.io/openai-agents-js/guides/mcp/',
    'OpenAI Agents SDK Agent -> tool() -> invokeFunctionTool()',
    'OpenAI Agents SDK Agent -> MCPServerStdio -> getAllMcpTools() -> invokeFunctionTool()',
    'OpenAI Agents SDK Agent -> MCPServerStreamableHttp -> getAllMcpTools() -> invokeFunctionTool()',
    'MCPServerStdio',
    'MCPServerStreamableHttp',
    'https://bilig.proompteng.ai/mcp',
    'set_cell_contents_and_readback',
    'persistence.persisted',
    'restoredMatchesAfter',
  ] as const) {
    requireIncludes(openAiAgentsSdkDoc, required, 'docs/openai-agents-sdk-workpaper-tool.md')
  }
  for (const required of [
    'title: ChatGPT Apps WorkPaper MCP',
    'description: Add Bilig as a ChatGPT Developer Mode remote MCP app for WorkPaper workbook readback proof.',
    'https://bilig.proompteng.ai/mcp',
    'Settings -> Apps & Connectors -> Advanced settings',
    'https://developers.openai.com/api/docs/mcp',
    'https://developers.openai.com/apps-sdk/deploy/connect-chatgpt',
    'set_cell_contents_and_readback',
    'Summary!A1:B4',
    'Summary!B3',
    'persistence.persisted: false',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp',
    'Access-Control-Request-Headers: accept, content-type, mcp-protocol-version',
    'Use an Apps SDK component resource later',
  ] as const) {
    requireIncludes(chatgptAppsWorkpaperMcpDoc, required, 'docs/chatgpt-apps-workpaper-mcp.md')
  }
  for (const [path, content] of [
    ['README.md', readme],
    ['docs/agent-framework-workbook-tools.md', agentFrameworkWorkbookToolsDoc],
    ['docs/llms.txt', llms],
  ] as const) {
    requireIncludes(content, 'ChatGPT Apps WorkPaper MCP', path)
    requireIncludes(content, 'chatgpt-apps-workpaper-mcp', path)
  }
  requireIncludes(
    agentToolCallingDoc,
    'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
    'docs/agent-workpaper-tool-calling-recipe.md',
  )
  requireIncludes(
    aiSdkLangChainDoc,
    'description: Wrap @bilig/workpaper WorkPaper reads, verified edits, formula contracts, and persistence checks as AI SDK, LangChain, Mastra, LlamaIndex.TS, LangGraph.js, CopilotKit, and Cloudflare Agents tools',
    'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
  )
  requireIncludes(
    aiSdkLangChainDoc,
    'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
    'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
  )
  requireIncludes(aiSdkLangChainDoc, 'Mastra `createTool()`', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
  requireIncludes(aiSdkLangChainDoc, 'LlamaIndex.TS tools', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
  requireIncludes(aiSdkLangChainDoc, 'LangGraph.js `ToolNode`', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
  requireIncludes(aiSdkLangChainDoc, 'CopilotKit `useCopilotAction`', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
  requireIncludes(aiSdkLangChainDoc, 'Cloudflare Agents API and agent tools', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
  const agentFrameworkDocs = await Promise.all(
    agentFrameworkDocRequirements.map(async ({ path, includes }) => ({
      path,
      includes,
      content: await readFile(join(repoRoot, path), 'utf8'),
    })),
  )
  for (const { path, includes, content } of agentFrameworkDocs) {
    for (const required of includes) {
      requireIncludes(content, required, path)
    }
  }
  for (const required of [
    '"@mastra/core": "1.38.0"',
    '"smoke": "node --disable-warning=DEP0205 --import tsx src/mastra-workpaper-tool.ts"',
  ] as const) {
    requireIncludes(mastraExamplePackageJson, required, 'examples/mastra-workpaper-tool/package.json')
  }
  for (const required of [
    'createTool',
    'Mastra createTool -> execute -> WorkPaper readback',
    'set-workpaper-input-cell',
    'restoredMatchesAfter',
  ] as const) {
    requireIncludes(mastraExampleSource, required, 'examples/mastra-workpaper-tool/src/mastra-workpaper-tool.ts')
  }
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'description: Expose @bilig/workpaper workbook reads, verified edits, formula contracts, persistence checks, resources, and prompts through MCP.',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'pnpm --dir examples/headless-workpaper run agent:mcp-tools',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, 'npm run --silent agent:mcp-stdio', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, '## Copy-Paste JSON-RPC Transcript', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'ChatGPT Apps WorkPaper MCP', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'https://chatgpt.com', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'pnpm --dir examples/headless-workpaper run agent:mcp-transcript',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, '"structuredContent": {', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, '"restoredMatchesAfter": true', 'docs/mcp-workpaper-tool-server.md')
  for (const required of [
    '## MCP Inspector Smoke',
    'https://modelcontextprotocol.io/docs/tools/inspector',
    '@modelcontextprotocol/inspector@latest',
    '--method tools/list',
    '--method tools/call',
    '--tool-name set_workpaper_input_cell',
    '--tool-arg value=0.4',
    'read_workpaper_summary',
    'set_workpaper_input_cell',
    '"before": { "expectedArr": 60000 }',
    '"after": { "expectedArr": 96000 }',
    '"restored": { "expectedArr": 96000 }',
    '"serializedBytes": 1162',
    'This Inspector smoke launches default demo mode',
    'Keep the Inspector proxy on localhost',
  ]) {
    requireIncludes(mcpWorkPaperToolServerDoc, required, 'docs/mcp-workpaper-tool-server.md')
  }
  for (const required of [
    '"ai": "6.0.195"',
    '"agent:ai-sdk-generate-text": "node --disable-warning=DEP0205 --import tsx ai-sdk-generate-text-tool-smoke.ts"',
    '"agent:ai-sdk-stream-text": "node --disable-warning=DEP0205 --import tsx ai-sdk-stream-text-tool-smoke.ts"',
    '"agent:mcp-transcript": "node --disable-warning=DEP0205 --import tsx mcp-stdio-transcript.ts"',
  ] as const) {
    requireIncludes(headlessExamplePackageJson, required, 'examples/headless-workpaper/package.json')
  }
  requireIncludes(rootPackageJson, '"workpaper:smoke:external": "bun scripts/workpaper-external-smoke.ts"', 'package.json')
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    `npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge`,
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    `npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp`,
    'docs/mcp-workpaper-tool-server.md',
  )
  for (const required of [
    '### Cursor demo server config',
    '"bilig-workpaper": {',
    '"command": "npm"',
    '"@bilig/headless@latest"',
    '"bilig-workpaper-mcp"',
    'That default demo mode exposes two tools: `read_workpaper_summary` and',
    '`set_workpaper_input_cell`.',
    'Use the `bilig-workpaper` MCP server. Read `Summary!A1:B5`',
    '`Inputs!B3` to `0.4` with `set_workpaper_input_cell`',
    '`expectedArr` before and after',
    '`expectedArr` `60000` before the edit and `96000`',
  ]) {
    requireIncludes(mcpWorkPaperToolServerDoc, required, 'docs/mcp-workpaper-tool-server.md')
  }
  requireIncludes(mcpWorkPaperToolServerDoc, '[MCP client setup guide](mcp-client-setup.md#cursor)', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, `\`${workpaperPackageSpec}\` in file-backed mode`, 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    `npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable`,
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, '`list_sheets`, `read_range`, `read_cell`', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'WorkPaper JSON back to the same file after `set_cell_contents`',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, 'resources/list', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'prompts/list', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'bilig://workpaper/agent-handoff', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'edit_and_verify_workpaper', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'io.github.proompteng/bilig-workpaper', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, '/workpaper/pricing.workpaper.json', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, '`validate_formula`', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'https://bilig.proompteng.ai/.well-known/mcp/server-card.json',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
    'docs/mcp-workpaper-tool-server.md',
  )
  for (const required of [
    `ARG BILIG_WORKPAPER_VERSION=${headlessPackageVersion}`,
    'npm install --omit=dev "@bilig/workpaper@${BILIG_WORKPAPER_VERSION}"',
    'ENTRYPOINT ["./node_modules/.bin/bilig-workpaper-mcp", "--workpaper", "/workpaper/pricing.workpaper.json", "--init-demo-workpaper", "--writable"]',
    'io.modelcontextprotocol.server.name="io.github.proompteng/bilig-workpaper"',
  ]) {
    requireIncludes(dockerfile, required, 'Dockerfile')
  }
  requireIncludes(mcpWorkPaperToolServerDoc, 'tools/list', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'tools/call', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'MCP tool annotations', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    '`read_workpaper_summary` is read-only, idempotent, and closed-world',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    '`set_workpaper_input_cell` mutates the local WorkPaper state, is idempotent',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'https://modelcontextprotocol.io/specification/2025-11-25/server/tools',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, 'https://github.com/proompteng/bilig/discussions/230', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(agentToolCallingDoc, 'https://github.com/proompteng/bilig/discussions/335', 'docs/agent-workpaper-tool-calling-recipe.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'mcp-client-setup.md', 'docs/mcp-workpaper-tool-server.md')
  requireSpreadsheetMcpServerComparisonDiscovery({ spreadsheetMcpServerComparison })
  for (const required of [
    'description: Live directory and install status for the Bilig WorkPaper MCP server',
    `npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp`,
    'io.github.proompteng/bilig-workpaper',
    'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
    'https://glama.ai/mcp/servers/proompteng/bilig',
    'https://bilig.proompteng.ai/.well-known/mcp/server-card.json',
    'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
    'https://proompteng.github.io/bilig/.well-known/mcp.json',
    'https://proompteng.github.io/bilig/.well-known/mcp-server-card.json',
    'https://smithery.ai/servers/gkonushev/bilig-workpaper',
    'npx -y smithery mcp add gkonushev/bilig-workpaper',
    "npx -y smithery tool call bilig-workpaper list_sheets '{}'",
    'Static MCP server card',
    'https://github.com/chatmcp/mcpso/issues/2295',
    'https://github.com/cline/mcp-marketplace/issues/1557',
    'https://github.com/docker/mcp-registry/pull/3606',
    'https://github.com/aaif-goose/goose/pull/9315',
    'Docker MCP Registry             | Open for maintainer review; current proof refreshed in the existing pull request',
    'Goose MCP catalog               | Closed by maintainer while Goose pauses new MCP server additions',
    'mcp.so                          | Submitted for maintainer review; issue body refreshed on May 19',
    'Cline MCP Marketplace           | Submitted for maintainer review; issue body refreshed on May 19',
    'The Docker MCP Registry pull request was refreshed in place, not duplicated.',
    'Latest checked result on June 3, 2026: the pull request is still open',
    'reports merge state `BLOCKED`',
    'do not open a second Docker listing',
    'The Goose MCP catalog pull request was closed on May 19, 2026',
    'Do not resubmit there until maintainers reopen that path',
    'The mcp.so and Cline MCP Marketplace submissions were refreshed on May 19, 2026',
    'by editing the existing issue bodies, not by adding more comments',
    'https://mcpserver.cc/en?q=bilig',
    'bcdce4e1-3b05-4be2-b611-2a2abb8baf79',
    'https://agentndx.ai/browse?q=bilig',
    'AgentNDX submission was accepted for review on May 13, 2026',
    'https://github.com/YuzeHao2023/Awesome-MCP-Servers/pull/244',
    'https://github.com/toolsdk-ai/toolsdk-mcp-registry/pull/309',
    'https://github.com/ever-works/awesome-mcp-servers-data/pull/4',
    'https://github.com/jmstfv/mcpserve/pull/19',
    'https://github.com/MCPFind/mcp-find/pull/37',
    'https://github.com/mctrinh/awesome-mcp-servers/pull/46',
    'https://mcprepository.com/proompteng/bilig',
    'MCPRepository search returns a live Bilig page',
    'Live; `smithery mcp add` smoke connected and listed demo workbook sheets',
    'Live; latest marker and repo package versions are recorded below',
    'Live with `Try in Browser`; file-backed tools indexed with A-grade TDQS',
    'Live in PulseMCP-backed lookup as `Bilig WorkPaper`',
    'https://www.pulsemcp.com/servers?search=bilig&q=bilig',
    'https://github.com/proompteng/bilig/issues/384',
    'Live verification on May 19, 2026 found Bilig WorkPaper\nthrough a PulseMCP-backed lookup query',
    'Smithery lists Bilig WorkPaper as `gkonushev/bilig-workpaper`',
    'Live verification on May 19, 2026 returned a connected Smithery server',
    'Glama lists Bilig WorkPaper publicly with TypeScript, Developer Tools',
    'file-backed tools',
    'A-grade Tool Definition Quality',
    "Glama's source crawl, hosted smoke build, and JSON API can refresh on\ndifferent cadences",
    'Latest checked result on June 3, 2026: Live.',
    'official Registry',
    'latest-marked server `io.github.proompteng/bilig-workpaper` is version',
    `\`${officialRegistryLatestMarkedVersion}\`, package \`@bilig/workpaper\` is version \`${officialRegistryLatestMarkedVersion}\``,
    'Keep the repo package version separate from\nRegistry readback because npm and Registry publication can refresh on different\ncadences.',
    'entry was',
    `\`${officialRegistryLatestMarkedUpdatedAt}\``,
    'hosted server-card path still advertises remote',
    'https://bilig.proompteng.ai/mcp',
    'limit=100',
    'read_workpaper_summary',
    'set_workpaper_input_cell',
    'file-backed mode',
    '/workpaper/pricing.workpaper.json',
    '--init-demo-workpaper',
    'set_cell_contents',
    'validate_formula',
  ]) {
    requireIncludes(mcpSpreadsheetServerDirectoryDoc, required, 'docs/mcp-spreadsheet-server-directory.md')
  }
  requireMatches(
    mcpSpreadsheetServerDirectoryDoc,
    new RegExp(`current\\s+repo\\s+package\\s+version\\s+is\\s+\`${escapeRegExp(headlessPackageVersion)}\``, 'u'),
    `current repo package version is \`${headlessPackageVersion}\``,
    'docs/mcp-spreadsheet-server-directory.md',
  )
  requireIncludes(
    mcpClientSetupDoc,
    'description: Remote MCP smoke endpoint and local stdio configuration for Bilig WorkPaper in Claude, Cursor, Junie, Trae, Zed, VS Code, Cline, and Codex.',
    'docs/mcp-client-setup.md',
  )
  for (const required of [
    '## Remote smoke in 30 seconds',
    '## Smithery install',
    'https://smithery.ai/servers/gkonushev/bilig-workpaper',
    'npx -y smithery mcp add gkonushev/bilig-workpaper',
    "npx -y smithery tool call bilig-workpaper list_sheets '{}'",
    'https://bilig.proompteng.ai/mcp',
    'https://bilig.proompteng.ai/.well-known/mcp/server-card.json',
    'mcp-protocol-version: 2025-11-25',
    '## Persistent file-backed stdio server',
    `npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp`,
    `bilig-workpaper-mcp --from-xlsx ./pricing.xlsx`,
    'analyze_workbook_risk',
    'does not certify Excel compatibility',
    'edits stay in memory',
    'The third command imports an XLSX into an in-memory WorkPaper server',
    'The client configs below use file-backed mode',
    `"args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]`,
    `args = ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]`,
    'pnpm mcpb:workpaper:build',
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
    'claude-desktop-mcpb-workpaper.md',
    'claude mcp add-json bilig-workpaper',
    '.mcp.json',
    '.cursor/mcp.json',
    '.junie/mcp/mcp.json',
    '.trae/mcp.json',
    'Use the bilig-workpaper MCP server from .trae/mcp.json after Project MCP is',
    'enabled. List sheets, read Inputs!B3 and Summary!B3',
    'Trae reads the project rule from `.trae/rules/bilig-workpaper.md`',
    '.zed/settings.json',
    'mcp/bilig-workpaper.mcp.json',
    'Use the biligWorkpaperFile MCP server. List sheets, read Summary!A1:B5',
    'Use the biligWorkpaperFile MCP server from .junie/mcp/mcp.json.',
    'Junie reads project guidelines from `.junie/AGENTS.md`',
    '"context_servers"',
    '"bilig-workpaper"',
    'mcp:bilig-workpaper:set_cell_contents_and_readback',
    'Use the bilig-workpaper context server from .zed/settings.json.',
    'set Inputs!B3 to 0.4 with set_cell_contents_and_readback',
    `code --add-mcp '{"name":"biligWorkpaperFile","type":"stdio","command":"npm","args":["exec","--package","${workpaperPackageSpec}","--","bilig-workpaper-mcp","--workpaper","\${workspaceFolder}/.bilig/pricing.workpaper.json","--init-demo-workpaper","--writable"]}'`,
    `code --add-mcp '{"name":"biligWorkpaperDemo","type":"http","url":"${remoteMcpEndpoint}"}'`,
    'The useful Cursor tool set includes `list_sheets`, `read_range`',
    '.vscode/mcp.json',
    '"autoApprove": []',
    'cline config mcp',
    'cline config mcp --json',
    '~/.cline/mcp.json',
    '[mcp_servers.bilig-workpaper]',
    'Codex spreadsheet MCP server',
    'Use the Bilig WorkPaper MCP server from Codex.',
    'https://code.visualstudio.com/docs/copilot/reference/mcp-configuration',
    'https://docs.cline.bot/mcp/configuring-mcp-servers',
    'https://zed.dev/docs/ai/mcp',
    'https://zed.dev/docs/ai/rules',
    'https://zed.dev/docs/ai/tool-permissions',
    'https://docs.trae.ai/ide/model-context-protocol',
    'https://docs.trae.ai/ide/add-mcp-servers',
    'https://docs.trae.ai/ide/rules',
    'https://docs.trae.ai/ide/skills',
    'https://platform.openai.com/docs/docs-mcp',
  ]) {
    requireIncludes(mcpClientSetupDoc, required, 'docs/mcp-client-setup.md')
  }
  for (const forbidden of [
    'cline_mcp_settings.json',
    '~/.cline/data/settings/cline_mcp_settings.json',
    'data/settings/cline_mcp_settings.json',
    'https://docs.cline.bot/mcp/adding-and-configuring-servers',
  ] as const) {
    requireNotIncludes(mcpClientSetupDoc, forbidden, 'docs/mcp-client-setup.md')
  }
  requireIncludes(rootPackageJson, '"mcpb:workpaper:build": "tsx scripts/build-workpaper-mcpb.ts"', 'package.json')
  for (const required of [
    'description: Download or reproduce the Claude Desktop MCPB bundle for the published @bilig/headless WorkPaper MCP server',
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
    'Open the downloaded `.mcpb` file with Claude Desktop',
    'pnpm mcpb:workpaper:build',
    'BILIG_HEADLESS_VERSION=$(npm view @bilig/headless version)',
    'pnpm mcpb:workpaper:build -- --package-version "$BILIG_HEADLESS_VERSION"',
    'build/mcpb/bilig-workpaper.mcpb',
    'open build/mcpb/bilig-workpaper.mcpb',
    'list_sheets',
    'read_range',
    'read_cell',
    'set_cell_contents',
    'set_cell_contents_and_readback',
    'get_cell_display_value',
    'export_workpaper_document',
    'validate_formula',
    '[Bilig WorkPaper MCPB privacy policy](workpaper-mcpb-privacy.md)',
    '"entry_point": "server/index.js"',
    'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
  ]) {
    requireIncludes(claudeDesktopMcpbDoc, required, 'docs/claude-desktop-mcpb-workpaper.md')
  }
  requireGrowthSurfaceDiscovery(communityLaunchPack, headlessPackageVersion, llms, productHuntLaunchKit, requireIncludes)
  requireNotIncludes(llms, '## launch and feedback', 'docs/llms.txt')
  requireNotIncludes(llms, 'conversion-feedback comment after npm download and clone traffic review', 'docs/llms.txt')
  requireNotIncludes(llms, 'published dev article source', 'docs/llms.txt')
  for (const removedGrowthLink of llmsExternalSurfaceLinks) {
    requireNotIncludes(llms, removedGrowthLink, 'docs/llms.txt')
  }
  requireIncludes(
    aiSdkLangChainDoc,
    'https://ai-sdk.dev/docs/ai-sdk-core/tools-and-tool-calling',
    'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
  )
  requireIncludes(
    aiSdkLangChainDoc,
    'https://docs.langchain.com/oss/javascript/langchain/tools',
    'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
  )
  requireIncludes(
    agentToolCallLoopDoc,
    'description: A runnable @bilig/workpaper loop where an agent writes one workbook input',
    'docs/agent-spreadsheet-tool-call-loop.md',
  )
  for (const [path, content] of [
    ['docs/why-agents-need-workbook-apis.md', whyAgentsDoc],
    ['docs/agent-framework-workbook-tools.md', agentFrameworkWorkbookToolsDoc],
    ['docs/headless-workpaper-agent-handbook.md', headlessWorkpaperAgentHandbook],
    ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
    ['docs/openai-agents-sdk-workpaper-tool.md', openAiAgentsSdkDoc],
    ['docs/vercel-ai-sdk-langchain-spreadsheet-tool.md', aiSdkLangChainDoc],
    ['docs/mcp-workpaper-tool-server.md', mcpWorkPaperToolServerDoc],
    ['docs/mcp-spreadsheet-server-directory.md', mcpSpreadsheetServerDirectoryDoc],
    ['docs/mcp-client-setup.md', mcpClientSetupDoc],
    ['docs/claude-desktop-mcpb-workpaper.md', claudeDesktopMcpbDoc],
    ['docs/agent-spreadsheet-tool-call-loop.md', agentToolCallLoopDoc],
    ['docs/workbook-automation-examples-node.md', workbookAutomationExamplesDoc],
    ['docs/server-side-spreadsheet-automation-node.md', serverSideSpreadsheetAutomationNode],
    ['docs/google-sheets-api-alternative-node-workpaper.md', googleSheetsApiBoundaryDoc],
    ['docs/node-framework-workpaper-adapters.md', nodeFrameworkWorkpaperAdaptersDoc],
    ['docs/dev-to-workbook-apis-post.md', devToWorkbookApisPost],
  ] as const) {
    requireIncludes(content, 'image: /assets/github-social-preview.png', path)
  }

  const exceljsFormulaRecalculationNode = await readFile(join(docsRoot, 'exceljs-formula-recalculation-node.md'), 'utf8')
  for (const required of [
    'title: ExcelJS formula recalculation in Node.js',
    'canonical_url: https://proompteng.github.io/bilig/exceljs-formula-recalculation-node.html',
    'npm install exceljs @bilig/exceljs-formula-recalc',
    "import { recalculateExceljsWorkbook } from '@bilig/exceljs-formula-recalc'",
    'ExcelJS formula result not updating',
    'get computed value of Excel sheet cell in Node.js',
    'Use `@bilig/exceljs-formula-recalc` for an ExcelJS workbook that needs fresh formula',
    'Do not choose `@bilig/workpaper` only to generate styled XLSX files',
    'release feed nearby',
  ] as const) {
    requireIncludes(exceljsFormulaRecalculationNode, required, 'docs/exceljs-formula-recalculation-node.md')
  }

  requireIncludes(workbookAutomationExamplesDoc, '## 90-second npm-only check', 'docs/workbook-automation-examples-node.md')
  requireIncludes(
    workbookAutomationExamplesDoc,
    'curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts',
    'docs/workbook-automation-examples-node.md',
  )

  requireIncludes(issueTemplateConfig, 'https://github.com/proompteng/bilig/discussions/213', '.github/ISSUE_TEMPLATE/config.yml')
  requireIncludes(
    pullRequestTemplate,
    'For public docs or example work, include the page or discussion that a new',
    '.github/PULL_REQUEST_TEMPLATE.md',
  )

  for (const required of [
    '## Use-Case Chooser',
    'Formula-backed calculations inside a Node service',
    'Agent writeback that must prove the value after an edit',
    'XLSX parsing, export, styling, images, and workbook-file metadata',
    'Persisting a workbook document as JSON and restoring it later',
    'Embedding a spreadsheet UI that users edit directly',
    '[Node quickstart](try-bilig-headless-in-node.md)',
    '[agent tool-calling recipe](agent-workpaper-tool-calling-recipe.md)',
    '[SheetJS and ExcelJS boundary guide](sheetjs-exceljs-alternative-formula-workbook-api.md)',
    '[HyperFormula alternative notes](hyperformula-alternative-headless-workpaper.md)',
    '[documented Excel gaps](where-bilig-is-not-excel-compatible-yet.md)',
  ]) {
    requireIncludes(headlessSpreadsheetEngineComparison, required, 'docs/headless-spreadsheet-engine-comparison.md')
  }

  for (const required of [
    '## If you arrived from HN or LibHunt',
    'workbook-shaped calculation boundary',
    '[XLSX recalculation proof](xlsx-recalculation-proof.md)',
    '[LibHunt headless-spreadsheet topic](https://www.libhunt.com/topic/headless-spreadsheet)',
    'repo is the current source of truth for releases',
    'open an implementation gap',
  ]) {
    requireIncludes(headlessSpreadsheetEngineNodeServicesAgents, required, 'docs/headless-spreadsheet-engine-node-services-agents.md')
  }

  for (const [path, content] of [
    ['docs/sheetjs-exceljs-alternative-formula-workbook-api.md', sheetjsExceljsAlternativeFormulaWorkbookApi],
    ['docs/hyperformula-alternative-headless-workpaper.md', hyperformulaAlternativeHeadlessWorkpaper],
  ] as const) {
    requireIncludes(
      content,
      '[headless spreadsheet engine use-case chooser](headless-spreadsheet-engine-comparison.md#use-case-chooser)',
      path,
    )
  }

  for (const required of [
    'title: XLSX formula recalculation in Node.js',
    'canonical_url: https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html',
    'npm install @bilig/xlsx-formula-recalc',
    'npx --package @bilig/xlsx-formula-recalc xlsx-recalc pricing.xlsx',
    'cd bilig/examples/xlsx-recalculation-node',
    '"exportedReimportMatchesAfter": true',
    '"formulasSurvivedXlsxRoundTrip": true',
    "import { recalculateXlsx } from '@bilig/xlsx-formula-recalc'",
    'Use ExcelJS or SheetJS first when the job is workbook-file manipulation',
    'Use `@bilig/xlsx-formula-recalc` when the Node process must own the recalculated answer',
    'release feed nearby',
  ] as const) {
    requireIncludes(xlsxFormulaRecalculationNode, required, 'docs/xlsx-formula-recalculation-node.md')
  }

  for (const required of [
    'title: SheetJS and ExcelJS alternative for formula-backed workbook APIs',
    'canonical_url: https://proompteng.github.io/bilig/sheetjs-exceljs-alternative-formula-workbook-api.html',
    'Research date: 2026-05-20.',
    '## TypeScript WorkPaper Evaluation Path',
    '## Use The Narrow Bridge First',
    'npm --prefix examples/recalc-bridge-workflows run smoke',
    'xlsx-formula-recalc',
    'exceljs-formula-recalc',
    'npm install -D tsx typescript @types/node',
    'const workbook = WorkPaper.buildFromSheets({',
    'workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40)',
    'verified: before === 36864 && after === 46080 && afterRestore === after',
    'SheetJS Pro has a formula calculator component',
    'ExcelJS can store formulas and supplied results',
  ] as const) {
    requireIncludes(sheetjsExceljsAlternativeFormulaWorkbookApi, required, 'docs/sheetjs-exceljs-alternative-formula-workbook-api.md')
  }

  requireIncludes(nodeSpreadsheetFormulaEngine, 'cat > formula-engine-smoke.ts', 'docs/node-spreadsheet-formula-engine.md')

  const discussionDocs = {
    readme: ['README.md', readme],
    headless: ['packages/headless/README.md', headlessReadme],
    agent: ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
    index: ['docs/index.html', index],
    launch: ['internal/growth/community-launch-pack.md', communityLaunchPack],
    llms: ['docs/llms.txt', llms],
    mcp: ['docs/mcp-workpaper-tool-server.md', mcpWorkPaperToolServerDoc],
  } as const

  const discussionDocChecks = [
    ['https://github.com/proompteng/bilig/discussions/157', ['readme', 'headless', 'launch', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/213', ['readme', 'launch', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/230', ['mcp', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/167', ['launch', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/307', ['readme', 'headless', 'launch', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/308', ['readme', 'headless', 'launch', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/335', ['readme', 'headless', 'agent', 'launch', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/340', ['readme', 'headless', 'launch', 'llms']],
    ['https://github.com/proompteng/bilig/discussions/382', ['launch', 'llms']],
  ] as const

  for (const [url, docKeys] of discussionDocChecks) {
    for (const docKey of docKeys) {
      const [path, content] = discussionDocs[docKey]
      requireIncludes(content, url, path)
    }
  }

  requireStarterIssueDiscovery(starterIssues, llms)

  await requireHeadlessExampleDiscovery({
    repoRoot,
    docsRoot,
    readme,
    headlessReadme,
    index,
    llms,
    agentToolCallingDoc,
    aiSdkLangChainDoc,
  })
}
