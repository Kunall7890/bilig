import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

export async function requireAgentRuleChooserDiscovery(input: {
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly readme: string
}): Promise<void> {
  const { docsRoot, index, llms, llmsFull, readme } = input
  const [agentRuleChooser, agentStart] = await Promise.all([
    readFile(join(docsRoot, 'agent-rule-chooser.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-start.txt'), 'utf8'),
  ])

  for (const required of [
    'title: Coding agent rule chooser for Bilig WorkPaper',
    'description: Pick the Bilig instruction, rule, prompt, or MCP config for Codex',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'schemaVersion: "bilig-evaluator.v1"',
    'door: "agent-mcp"',
    '| Codex | `AGENTS.md`',
    '| Claude Code | `CLAUDE.md`, then `.claude/skills/bilig-workpaper/SKILL.md`',
    '| GitHub Copilot | `.github/copilot-instructions.md`',
    '| VS Code agent mode | `.github/copilot-instructions.md`',
    '| Cursor | `.cursor/rules/bilig-workpaper.mdc`',
    '| Kiro | `.kiro/steering/bilig-workpaper.md`',
    '`.kiro/settings/mcp.json` defines the project-local file-backed WorkPaper MCP server',
    '| Roo Code | `.roo/rules/bilig-workpaper.md`',
    '`.roo/mcp.json` defines the project-local file-backed WorkPaper MCP server',
    '| JetBrains Junie | `AGENTS.md` in the repo root',
    '`.junie/mcp/mcp.json` defines the file-backed WorkPaper MCP server',
    '| OpenHands | `AGENTS.md`, then `.agents/skills/bilig-workpaper/SKILL.md`',
    'openhands mcp add bilig-workpaper --transport stdio npm -- exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable',
    'https://docs.openhands.dev/openhands/usage/cli/mcp-servers',
    'https://docs.openhands.dev/overview/skills',
    '| OpenCode | `opencode.jsonc`, then `.opencode/agents/bilig-workpaper.md`',
    'https://opencode.ai/docs/config/',
    'https://opencode.ai/docs/mcp-servers/',
    'https://opencode.ai/docs/agents/',
    '[OpenCode WorkPaper MCP setup](opencode-workpaper-mcp.md)',
    '| Aider | `CONVENTIONS.md`, loaded by `.aider.conf.yml`.',
    'https://aider.chat/docs/usage/conventions.html',
    'https://aider.chat/docs/config/aider_conf.html',
    '| Goose | `examples/goose-workpaper-mcp/recipe.yaml`.',
    'https://goose-docs.ai/docs/getting-started/using-extensions/',
    'https://goose-docs.ai/docs/guides/recipes/recipe-reference/',
    'https://goose-docs.ai/docs/guides/goose-cli-commands/',
    '[Goose WorkPaper MCP recipe](goose-workpaper-mcp.md)',
    '| Windsurf/Cascade | `.devin/rules/bilig-workpaper.md`',
    '| Cline | `.clinerules/bilig-workpaper.md`',
    '| Continue | `.continue/rules/bilig-workpaper.md`',
    '| Gemini CLI | `gemini-extension.json` plus `gemini-workpaper-context.md`',
    'npm create @bilig/workpaper@latest . -- --add-agent',
    'npm create @bilig/workpaper@latest pricing-agent -- --agent',
    '`.vscode/mcp.json` uses the VS Code `servers` shape.',
    'Kiro reads workspace steering from `.kiro/steering/`',
    'project MCP servers',
    'from `.kiro/settings/mcp.json`',
    'Roo Code reads workspace rules from `.roo/rules/`',
    'Junie project MCP config lives at `.junie/mcp/mcp.json`',
    'Aider loads `CONVENTIONS.md` through `.aider.conf.yml`',
    'Cascade/Devin docs currently prefer `.devin/rules`',
    'mirror remains for compatible Windsurf/Cascade installs',
    'https://docs.windsurf.com/windsurf/cascade/memories',
    'https://docs.cline.bot/customization/cline-rules',
    'https://docs.continue.dev/customize/rules',
    'https://kiro.dev/docs/steering/',
    'https://kiro.dev/docs/mcp/configuration/',
    'https://roocodeinc.github.io/Roo-Code/features/custom-instructions',
    'https://roocodeinc.github.io/Roo-Code/features/mcp/using-mcp-in-roo/',
    'https://junie.jetbrains.com/docs/junie-plugin-mcp-settings.html',
    'https://junie.jetbrains.com/docs/guidelines-and-memory.html',
    '[Agent WorkPaper proof matrix](agent-proof-matrix.md)',
  ] as const) {
    requireIncludes(agentRuleChooser, required, 'docs/agent-rule-chooser.md')
  }

  requireNotIncludes(agentRuleChooser, 'No code changes are required', 'docs/agent-rule-chooser.md')

  for (const [path, content] of [
    ['README.md', readme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
  ] as const) {
    requireIncludes(content, 'agent-rule-chooser', path)
  }

  requireIncludes(llmsFull, '## Coding Agent Rule Chooser', 'docs/llms-full.txt')
  requireIncludes(llmsFull, 'Source: https://github.com/proompteng/bilig/blob/main/docs/agent-rule-chooser.md', 'docs/llms-full.txt')
  requireIncludes(agentStart, 'Agent rule chooser: https://proompteng.github.io/bilig/agent-rule-chooser.html', 'docs/agent-start.txt')
}
