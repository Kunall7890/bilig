const mcpbReleaseAssetUrl = 'https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb'
const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`

export function requireAgentJsonPublicDiscovery(parsedAgentJson: object): void {
  const agentJsonCapabilities = Reflect.get(parsedAgentJson, 'capabilities')
  if (!Array.isArray(agentJsonCapabilities)) {
    throw new Error('docs/.well-known/agent.json capabilities must be an array')
  }

  if (Reflect.get(parsedAgentJson, 'claude_code_instructions') !== 'https://github.com/proompteng/bilig/blob/main/CLAUDE.md') {
    throw new Error('docs/.well-known/agent.json must advertise root Claude Code instructions')
  }

  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'repo-local-agent-instructions' &&
        Reflect.get(capability, 'type') === 'project-agent-instructions' &&
        Reflect.get(capability, 'claude_code') === 'https://github.com/proompteng/bilig/blob/main/CLAUDE.md' &&
        Reflect.get(capability, 'codex') === 'https://github.com/proompteng/bilig/blob/main/AGENTS.md' &&
        Reflect.get(capability, 'claude_code_skill') ===
          'https://github.com/proompteng/bilig/blob/main/.claude/skills/bilig-workpaper/SKILL.md' &&
        Reflect.get(capability, 'claude_code_command') ===
          'https://github.com/proompteng/bilig/blob/main/.claude/commands/bilig-workpaper-proof.md' &&
        Reflect.get(capability, 'openhands_skill') ===
          'https://github.com/proompteng/bilig/blob/main/.agents/skills/bilig-workpaper/SKILL.md' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/agent-rule-chooser.html',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise repo-local agent instructions')
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
        Reflect.get(capability, 'name') === 'agent-proof-transcripts' &&
        Reflect.get(capability, 'type') === 'docs-proof-pack' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/agent-proof-transcripts.html' &&
        Reflect.get(capability, 'proof_command') ===
          'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the agent proof transcript pack')
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
        Reflect.get(capability, 'name') === 'openai-responses-workpaper-tool-calls' &&
        Reflect.get(capability, 'framework') === 'OpenAI Responses API' &&
        Reflect.get(capability, 'api_shape') === 'function_call -> function_call_output' &&
        Reflect.get(capability, 'command') === 'pnpm --dir examples/headless-workpaper run agent:openai-responses' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/openai-responses-workpaper-tool-call.html' &&
        Reflect.get(capability, 'source') ===
          'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/openai-responses-tool-wrapper.ts',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the OpenAI Responses function-calling capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'ai-sdk-generate-text-workpaper-tool' &&
        Reflect.get(capability, 'framework') === 'Vercel AI SDK' &&
        Reflect.get(capability, 'package') === 'ai' &&
        Reflect.get(capability, 'adapter') === '@bilig/workpaper/ai-sdk' &&
        Reflect.get(capability, 'api_shape') === 'generateText -> tool -> execute' &&
        Reflect.get(capability, 'command') === 'pnpm --dir examples/headless-workpaper run agent:ai-sdk-generate-text' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/vercel-ai-sdk-langchain-spreadsheet-tool.html' &&
        Reflect.get(capability, 'source') ===
          'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/ai-sdk-generate-text-tool-smoke.ts',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the AI SDK generateText WorkPaper tool capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'ai-sdk-stream-text-workpaper-tool' &&
        Reflect.get(capability, 'framework') === 'Vercel AI SDK' &&
        Reflect.get(capability, 'package') === 'ai' &&
        Reflect.get(capability, 'adapter') === '@bilig/workpaper/ai-sdk' &&
        Reflect.get(capability, 'api_shape') === 'streamText -> tool -> execute' &&
        Reflect.get(capability, 'command') === 'pnpm --dir examples/headless-workpaper run agent:ai-sdk-stream-text' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/vercel-ai-sdk-langchain-spreadsheet-tool.html' &&
        Reflect.get(capability, 'source') ===
          'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/ai-sdk-stream-text-tool-smoke.ts',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the AI SDK streamText WorkPaper tool capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'browser-use-workpaper-formula-tool' &&
        Reflect.get(capability, 'framework') === 'Browser Use' &&
        Reflect.get(capability, 'command') ===
          'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/browser-use-workpaper-formula-tool.html' &&
        Reflect.get(capability, 'source') === 'https://github.com/proompteng/bilig/blob/main/docs/browser-use-workpaper-formula-tool.md' &&
        Reflect.get(capability, 'upstream_pr') === 'https://github.com/browser-use/browser-use/pull/4909',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the Browser Use WorkPaper formula tool capability')
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
        Reflect.get(capability, 'name') === 'google-adk-workpaper-mcp' &&
        Reflect.get(capability, 'framework') === 'Google Agent Development Kit' &&
        Reflect.get(capability, 'command') ===
          'uv run --python 3.12 --with google-adk --with mcp python examples/google-adk-workpaper-mcp/google_adk_workpaper_mcp.py --output .tmp/google-adk-workpaper-proof.json' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/google-adk-workpaper-mcp.html' &&
        Reflect.get(capability, 'source') === 'https://github.com/proompteng/bilig/tree/main/examples/google-adk-workpaper-mcp',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the Google ADK WorkPaper MCP capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'openhands-workpaper-mcp' &&
        Reflect.get(capability, 'framework') === 'OpenHands' &&
        Reflect.get(capability, 'command') ===
          'openhands mcp add bilig-workpaper --transport stdio npm -- exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable' &&
        Reflect.get(capability, 'config_path') === '~/.openhands/mcp.json' &&
        Reflect.get(capability, 'skill_path') === 'https://github.com/proompteng/bilig/blob/main/.agents/skills/bilig-workpaper/SKILL.md' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/openhands-workpaper-mcp.html' &&
        Reflect.get(capability, 'source') === 'https://github.com/proompteng/bilig/blob/main/docs/openhands-workpaper-mcp.md',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the OpenHands WorkPaper MCP capability')
  }
  if (
    !hasCapability(
      agentJsonCapabilities,
      (capability) =>
        Reflect.get(capability, 'name') === 'mastra-workpaper-tool' &&
        Reflect.get(capability, 'framework') === 'Mastra' &&
        Reflect.get(capability, 'api_shape') === 'createTool -> execute -> WorkPaper readback' &&
        Reflect.get(capability, 'command') === 'pnpm --dir examples/mastra-workpaper-tool run smoke' &&
        Reflect.get(capability, 'docs') === 'https://proompteng.github.io/bilig/mastra-workpaper-spreadsheet-tool.html' &&
        Reflect.get(capability, 'source') === 'https://github.com/proompteng/bilig/tree/main/examples/mastra-workpaper-tool',
    )
  ) {
    throw new Error('docs/.well-known/agent.json must advertise the Mastra WorkPaper tool capability')
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
  'https://proompteng.github.io/bilig/agent-proof-transcripts.html',
  'https://github.com/proompteng/bilig/blob/main/CLAUDE.md',
  'https://github.com/proompteng/bilig/blob/main/.mcp.json',
  'https://github.com/proompteng/bilig/blob/main/.cursor/mcp.json',
  'https://github.com/proompteng/bilig/blob/main/mcp/bilig-workpaper.mcp.json',
  'https://proompteng.github.io/bilig/chatgpt-apps-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/openai-agents-sdk-workpaper-tool.html',
  'https://proompteng.github.io/bilig/openai-responses-workpaper-tool-call.html',
  'https://proompteng.github.io/bilig/browser-use-workpaper-formula-tool.html',
  'https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html',
  'https://proompteng.github.io/bilig/mastra-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/llamaindex-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/agno-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/pydantic-ai-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/google-adk-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/openhands-workpaper-mcp.html',
  'https://github.com/proompteng/bilig/blob/main/docs/openhands-workpaper-mcp.md',
  'https://github.com/proompteng/bilig/blob/main/.agents/skills/bilig-workpaper/SKILL.md',
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
