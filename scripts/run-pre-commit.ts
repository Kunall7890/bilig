#!/usr/bin/env tsx

import { spawnSync } from 'node:child_process'

const generatedAgentDiscoveryFiles = [
  'docs/AGENTS.md',
  'docs/agent.json',
  'docs/skill.md',
  'docs/skill.txt',
  'docs/llms-full.txt',
  'docs/.well-known/agent.json',
  'docs/.well-known/agent-skills/index.json',
  'docs/.well-known/agent-skills/bilig-workpaper/SKILL.md',
  'docs/.well-known/agent-skills/bilig-workpaper/SKILL.txt',
  'docs/.well-known/skills/index.json',
  'docs/.well-known/skills/bilig-workpaper/SKILL.md',
  'docs/.well-known/skills/bilig-workpaper/SKILL.txt',
  'docs/.well-known/mcp/server-card.json',
  'docs/.well-known/mcp.json',
  'docs/.well-known/mcp-server-card.json',
  'skills/bilig-workpaper/SKILL.md',
  'packages/headless/SKILL.md',
  'packages/headless/AGENTS.md',
  'packages/bilig/SKILL.md',
  'packages/bilig/AGENTS.md',
  'packages/workpaper/SKILL.md',
  'packages/workpaper/AGENTS.md',
]

function run(command: string, args: string[]) {
  const result = spawnSync(command, args, {
    stdio: 'inherit',
    shell: false,
  })

  if (result.status !== 0) {
    process.exit(result.status ?? 1)
  }
}

run('pnpm', ['agent:discovery:generate'])
run('git', ['add', ...generatedAgentDiscoveryFiles])
run('pnpm', ['exec', 'lint-staged'])
