const mcpbReleaseAssetUrl = 'https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb'
const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`

export function requireAgentJsonPublicDiscovery(parsedAgentJson: object): void {
  const agentJsonCapabilities = Reflect.get(parsedAgentJson, 'capabilities')
  if (!Array.isArray(agentJsonCapabilities)) {
    throw new Error('docs/.well-known/agent.json capabilities must be an array')
  }

  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'agent-start' &&
        Reflect.get(capability, 'url') === 'https://proompteng.github.io/bilig/agent-start.txt' &&
        Reflect.get(capability, 'well_known_url') === 'https://proompteng.github.io/bilig/.well-known/agent-start.txt',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the compact agent-start capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'file-backed-workpaper-mcp' &&
        Reflect.get(capability, 'server_card') === 'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the file-backed MCP capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'repo-local-mcp-configs' &&
        Reflect.get(capability, 'type') === 'project-mcp-configs' &&
        Reflect.get(capability, 'claude_code') === 'https://github.com/proompteng/bilig/blob/main/.mcp.json' &&
        Reflect.get(capability, 'cursor') === 'https://github.com/proompteng/bilig/blob/main/.cursor/mcp.json' &&
        Reflect.get(capability, 'vscode') === 'https://github.com/proompteng/bilig/blob/main/.vscode/mcp.json' &&
        Reflect.get(capability, 'reusable') === 'https://github.com/proompteng/bilig/blob/main/mcp/bilig-workpaper.mcp.json' &&
        Reflect.get(capability, 'workpaper_state_path') === './.bilig/pricing.workpaper.json',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the repo-local MCP config capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'remote-workpaper-mcp-demo' &&
        Reflect.get(capability, 'endpoint') === 'https://bilig.proompteng.ai/mcp',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the remote MCP demo capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'openai-agents-sdk-hosted-workpaper-mcp' &&
        Reflect.get(capability, 'framework') === 'OpenAI Agents SDK' &&
        Reflect.get(capability, 'endpoint') === 'https://bilig.proompteng.ai/mcp' &&
        Reflect.get(capability, 'command') === 'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the OpenAI Agents SDK hosted MCP smoke capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'chatgpt-apps-workpaper-mcp' &&
        Reflect.get(capability, 'framework') === 'ChatGPT Apps' &&
        Reflect.get(capability, 'endpoint') === 'https://bilig.proompteng.ai/mcp' &&
        Reflect.get(capability, 'authentication_required') === false &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/chatgpt-apps-workpaper-mcp.html',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the ChatGPT Apps remote MCP capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'claude-desktop-mcpb' &&
        Reflect.get(capability, 'type') === 'mcpb-desktop-extension' &&
        Reflect.get(capability, 'download_url') === mcpbReleaseAssetUrl &&
        Reflect.get(capability, 'checksum_url') === mcpbReleaseChecksumUrl,
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the Claude Desktop MCPB release asset')
  }

  const agentJsonPublicEntrypoints = Reflect.get(parsedAgentJson, 'public_entrypoints')
  if (!Array.isArray(agentJsonPublicEntrypoints) || !agentJsonPublicEntrypoints.every((entrypoint) => typeof entrypoint === 'string')) {
    throw new Error('docs/.well-known/agent.json public_entrypoints must be a string array')
  }
  for (const requiredEntrypoint of requiredPublicEntrypoints) {
    if (!agentJsonPublicEntrypoints.includes(requiredEntrypoint)) {
      throw new Error(`docs/.well-known/agent.json public_entrypoints is missing ${requiredEntrypoint}`)
    }
  }
}

function hasCapability(capabilities: readonly unknown[], predicate: (capability: object) => boolean): boolean {
  return capabilities.some(
    (capability) => typeof capability === 'object' && capability !== null && !Array.isArray(capability) && predicate(capability),
  )
}

const requiredPublicEntrypoints = [
  'https://proompteng.github.io/bilig/llms.txt',
  'https://proompteng.github.io/bilig/llms-full.txt',
  'https://proompteng.github.io/bilig/llms-install.html',
  'https://github.com/proompteng/bilig/blob/main/llms-install.md',
  'https://proompteng.github.io/bilig/.well-known/llms.txt',
  'https://proompteng.github.io/bilig/.well-known/llms-full.txt',
  'https://proompteng.github.io/bilig/agent-start.txt',
  'https://proompteng.github.io/bilig/.well-known/agent-start.txt',
  'https://github.com/proompteng/bilig/blob/main/.mcp.json',
  'https://github.com/proompteng/bilig/blob/main/.cursor/mcp.json',
  'https://github.com/proompteng/bilig/blob/main/mcp/bilig-workpaper.mcp.json',
  'https://proompteng.github.io/bilig/chatgpt-apps-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/openai-agents-sdk-workpaper-tool.html',
  'https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html',
  'https://proompteng.github.io/bilig/llamaindex-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/agno-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/pydantic-ai-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/crewai-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/cloudflare-agents-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/semantic-kernel-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/gemini-cli-workpaper-extension.html',
  'https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html',
  'https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html',
  'https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html',
  'https://proompteng.github.io/bilig/triggerdev-workpaper-task.html',
  'https://proompteng.github.io/bilig/temporal-workpaper-activity.html',
] as const
