#!/usr/bin/env tsx

import { spawnSync } from 'node:child_process'

const generatedAgentDiscoveryFiles = [
  'packages/headless/SKILL.md',
  'packages/bilig/SKILL.md',
  'packages/workpaper/SKILL.md',
  'packages/xlsx-formula-recalc/SKILL.md',
  'packages/sheetjs-formula-recalc/SKILL.md',
  'packages/exceljs-formula-recalc/SKILL.md',
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
