import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireIncludes } from './check-docs-discovery-core.ts'

export async function requireAgentEvaluatorDiscovery(input: {
  readonly docsRoot: string
  readonly readme: string
  readonly index: string
  readonly llms: string
  readonly runtimePackageVersion: string
}): Promise<void> {
  const { docsRoot, readme, index, llms, runtimePackageVersion } = input
  const agentAdoptionKit = await readFile(join(docsRoot, 'agent-adoption-kit.md'), 'utf8')
  const agentMcpEvaluator = await readFile(join(docsRoot, 'eval-agent-mcp.md'), 'utf8')
  const workpaperServiceEvaluator = await readFile(join(docsRoot, 'eval-workpaper-service.md'), 'utf8')
  const expectedWorkpaperVersion = `"@bilig/workpaper": "${runtimePackageVersion}"`
  const expectedXlsxVersion = `"xlsx-formula-recalc": "${runtimePackageVersion}"`

  requireIncludes(readme, 'The published package also carries `AGENTS.md`', 'README.md')
  requireIncludes(readme, 'CLAUDE.md', 'README.md')
  requireIncludes(readme, '.claude/skills/bilig-workpaper/SKILL.md', 'README.md')
  requireIncludes(readme, '.claude/commands/bilig-workpaper-proof.md', 'README.md')
  requireIncludes(readme, '.cursor/rules/bilig-workpaper.mdc', 'README.md')
  requireIncludes(readme, '.devin/rules/bilig-workpaper.md', 'README.md')
  requireIncludes(readme, '.windsurf/rules/bilig-workpaper.md', 'README.md')
  requireIncludes(readme, '.clinerules/bilig-workpaper.md', 'README.md')
  requireIncludes(readme, '.continue/rules/bilig-workpaper.md', 'README.md')
  requireIncludes(readme, '.zed/settings.json', 'README.md')
  requireIncludes(readme, 'opencode.jsonc', 'README.md')
  requireIncludes(readme, '.opencode/agents/bilig-workpaper.md', 'README.md')
  requireIncludes(readme, 'host handoff prompt', 'README.md')
  requireIncludes(readme, 'docs/agent-adoption-kit.md', 'README.md')
  requireIncludes(readme, 'npx --yes skills@latest add https://bilig.proompteng.ai --list', 'README.md')
  requireIncludes(readme, 'npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list', 'README.md')
  requireIncludes(index, './headless-workpaper-agent-handbook.html', 'docs/index.html')
  requireIncludes(index, './agent-adoption-kit.html', 'docs/index.html')
  requireIncludes(
    agentAdoptionKit,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'docs/agent-adoption-kit.md',
  )
  requireIncludes(
    agentAdoptionKit,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
    'docs/agent-adoption-kit.md',
  )
  requireIncludes(agentAdoptionKit, 'provider-backed-adapter-missing', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, '## Agent Manifest Gate', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'https://proompteng.github.io/bilig/.well-known/agent.json', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, '`public_entrypoints`', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, '`evaluator_doors`', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, '`proof_contract`', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'bilig-agent-start --json', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'schemaVersion: "bilig-evaluator.v1"', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'door: "agent-mcp"', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, expectedWorkpaperVersion, 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, expectedXlsxVersion, 'docs/agent-adoption-kit.md')
  requireIncludes(
    agentAdoptionKit,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
    'docs/agent-adoption-kit.md',
  )
  requireIncludes(agentAdoptionKit, '## Avoid Stored-Result False Negatives', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'Do not grade workbook automation from unchecked stored formula results.', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, '@bilig/exceljs-formula-recalc', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, '## Upstream Maintainer Notes', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'keep one thread per project and update it in place', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'pnpm agent:public-rules:check', 'docs/agent-adoption-kit.md')
  requireIncludes(agentAdoptionKit, 'The first upstream message should be a maintainer question', 'docs/agent-adoption-kit.md')
  requireIncludes(
    agentAdoptionKit,
    'Do not open duplicate issues, duplicate PRs, or broad directory submissions',
    'docs/agent-adoption-kit.md',
  )
  requireIncludes(
    agentMcpEvaluator,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'docs/eval-agent-mcp.md',
  )
  requireIncludes(
    agentMcpEvaluator,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
    'docs/eval-agent-mcp.md',
  )
  requireIncludes(agentMcpEvaluator, 'scenario: "provider-backed"', 'docs/eval-agent-mcp.md')
  requireIncludes(agentMcpEvaluator, 'provider-backed-adapter-missing', 'docs/eval-agent-mcp.md')
  requireIncludes(agentMcpEvaluator, '"schemaVersion": "bilig-evaluator.v1"', 'docs/eval-agent-mcp.md')
  requireIncludes(agentMcpEvaluator, '"door": "agent-mcp"', 'docs/eval-agent-mcp.md')
  requireIncludes(agentMcpEvaluator, expectedWorkpaperVersion, 'docs/eval-agent-mcp.md')
  requireIncludes(agentMcpEvaluator, expectedXlsxVersion, 'docs/eval-agent-mcp.md')
  requireIncludes(
    workpaperServiceEvaluator,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
    'docs/eval-workpaper-service.md',
  )
  requireIncludes(workpaperServiceEvaluator, '## Current evaluator transcript', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"door": "workpaper-service"', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, expectedWorkpaperVersion, 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"editedCell": "Inputs!B2"', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"dependentCell": "Summary!B2"', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"before": 24000', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"after": 38400', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"afterRestore": 38400', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"persistedDocumentBytes": 999', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"formulaReadbackChanged": true', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"exportedWorkPaperDocument": true', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"restoredMatchesAfter": true', 'docs/eval-workpaper-service.md')
  requireIncludes(workpaperServiceEvaluator, '"verified": true', 'docs/eval-workpaper-service.md')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/eval-agent-mcp.html', 'docs/llms.txt')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/agent-adoption-kit.html', 'docs/llms.txt')
  requireIncludes(llms, 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json', 'docs/llms.txt')
  requireIncludes(
    llms,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
    'docs/llms.txt',
  )
}
