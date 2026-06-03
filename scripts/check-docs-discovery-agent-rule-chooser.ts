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
    '| OpenHands | `AGENTS.md`, then `.agents/skills/bilig-workpaper/SKILL.md`',
    'openhands mcp add bilig-workpaper --transport stdio npm -- exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable',
    'https://docs.openhands.dev/openhands/usage/cli/mcp-servers',
    'https://docs.openhands.dev/overview/skills',
    '| Windsurf/Cascade | `.devin/rules/bilig-workpaper.md`',
    '| Cline | `.clinerules/bilig-workpaper.md`',
    '| Continue | `.continue/rules/bilig-workpaper.md`',
    '| Gemini CLI | `gemini-extension.json` plus `gemini-workpaper-context.md`',
    'npm create @bilig/workpaper@latest . -- --add-agent',
    'npm create @bilig/workpaper@latest pricing-agent -- --agent',
    '`.vscode/mcp.json` uses the VS Code `servers` shape.',
    'Cascade/Devin docs currently prefer `.devin/rules`',
    'mirror remains for compatible Windsurf/Cascade installs',
    'https://docs.windsurf.com/windsurf/cascade/memories',
    'https://docs.cline.bot/customization/cline-rules',
    'https://docs.continue.dev/customize/rules',
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
