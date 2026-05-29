import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

export async function requireAgentInstructionDiscovery(input: {
  readonly repoRoot: string
  readonly docsRoot: string
  readonly headlessPackageVersion: string
}): Promise<void> {
  const { repoRoot, docsRoot, headlessPackageVersion } = input
  const headlessPackageSpec = `@bilig/headless@${headlessPackageVersion}`
  const [
    docsAgentNotes,
    docsSkill,
    claudeProjectSkillNotes,
    claudeProjectCommandNotes,
    cursorProjectRuleNotes,
    windsurfProjectRuleNotes,
    clineProjectRuleNotes,
    continueProjectRuleNotes,
    copilotInstructions,
    copilotPrompt,
    vscodeMcpConfig,
    rootSkillNotes,
    workpaperPackageJson,
    workpaperPackageReadme,
    workpaperPackageAgentNotes,
    workpaperPackageSkillNotes,
    headlessAgentNotes,
    headlessSkillNotes,
  ] = await Promise.all([
    readFile(join(docsRoot, 'AGENTS.md'), 'utf8'),
    readFile(join(docsRoot, 'skill.md'), 'utf8'),
    readFile(join(repoRoot, '.claude', 'skills', 'bilig-workpaper', 'SKILL.md'), 'utf8'),
    readFile(join(repoRoot, '.claude', 'commands', 'bilig-workpaper-proof.md'), 'utf8'),
    readFile(join(repoRoot, '.cursor', 'rules', 'bilig-workpaper.mdc'), 'utf8'),
    readFile(join(repoRoot, '.windsurf', 'rules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.clinerules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.continue', 'rules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.github', 'copilot-instructions.md'), 'utf8'),
    readFile(join(repoRoot, '.github', 'prompts', 'bilig-workpaper-proof.prompt.md'), 'utf8'),
    readFile(join(repoRoot, '.vscode', 'mcp.json'), 'utf8'),
    readFile(join(repoRoot, 'skills', 'bilig-workpaper', 'SKILL.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'package.json'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'AGENTS.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'SKILL.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'AGENTS.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'SKILL.md'), 'utf8'),
  ])

  requireIncludes(workpaperPackageJson, '"bilig-agent-challenge": "./dist/agent-workbook-challenge-bin.js"', 'packages/bilig/package.json')
  requireIncludes(workpaperPackageJson, '"bilig-workpaper-mcp": "./dist/work-paper-mcp-stdio-bin.js"', 'packages/bilig/package.json')
  requireIncludes(workpaperPackageJson, '"AGENTS.md"', 'packages/bilig/package.json')
  requireIncludes(workpaperPackageJson, '"SKILL.md"', 'packages/bilig/package.json')
  requireIncludes(workpaperPackageReadme, 'The npm tarball includes `AGENTS.md`, `SKILL.md`', 'packages/bilig/README.md')
  requireIncludes(workpaperPackageReadme, 'npm exec --package bilig-workpaper -- bilig-agent-challenge', 'packages/bilig/README.md')
  requireIncludes(workpaperPackageAgentNotes, 'node_modules/bilig-workpaper', 'packages/bilig/AGENTS.md')
  requireIncludes(workpaperPackageAgentNotes, 'npm exec --package bilig-workpaper@', 'packages/bilig/AGENTS.md')
  requireIncludes(workpaperPackageSkillNotes, 'Use bilig-workpaper WorkPaper state', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, '## First Choice: Direct npm proof or TypeScript', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, 'npm, ExcelJS, SheetJS, xlsx-populate, and StackOverflow', 'packages/bilig/SKILL.md')
  requireIncludes(
    workpaperPackageSkillNotes,
    'Use MCP only when the host specifically needs an MCP client boundary',
    'packages/bilig/SKILL.md',
  )
  requireNotIncludes(workpaperPackageSkillNotes, '## First Choice: MCP', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, '"--package", "bilig-workpaper@', 'packages/bilig/SKILL.md')

  requireIncludes(headlessAgentNotes, '## Handoff prompt', 'packages/headless/AGENTS.md')
  requireIncludes(headlessAgentNotes, 'Do not claim success from a write call alone.', 'packages/headless/AGENTS.md')
  requireIncludes(
    headlessAgentNotes,
    `npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge --json`,
    'packages/headless/AGENTS.md',
  )
  requireIncludes(
    headlessAgentNotes,
    `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable`,
    'packages/headless/AGENTS.md',
  )
  requireIncludes(headlessSkillNotes, 'name: bilig-workpaper', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, '"bilig-formula-clinic"', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, '"./reduced.xlsx"', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, 'Do not trigger it for manual spreadsheet editing', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, '## Command Safety', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, 'argument array, not a shell-concatenated string', 'packages/headless/SKILL.md')
  requireNotIncludes(headlessSkillNotes, 'allowed-tools:', 'packages/headless/SKILL.md')
  requireNotIncludes(headlessSkillNotes, 'argument-hint:', 'packages/headless/SKILL.md')

  requireIncludes(docsAgentNotes, '## Discovery Order', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'Do not claim success from a write call alone.', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.claude/skills/bilig-workpaper/SKILL.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.claude/commands/bilig-workpaper-proof.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.cursor/rules/bilig-workpaper.mdc', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.windsurf/rules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.clinerules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.continue/rules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.github/copilot-instructions.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.github/prompts/bilig-workpaper-proof.prompt.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.vscode/mcp.json', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json', 'docs/AGENTS.md')
  requireIncludes(docsSkill, 'name: bilig-workpaper', 'docs/skill.md')
  requireIncludes(docsSkill, '## Required Verification', 'docs/skill.md')
  requireIncludes(docsSkill, '## Command Safety', 'docs/skill.md')
  requireIncludes(docsSkill, 'npx --yes skills@latest add https://bilig.proompteng.ai --list', 'docs/skill.md')
  requireIncludes(docsSkill, 'npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list', 'docs/skill.md')
  requireIncludes(docsSkill, 'returned `tools` array as the source', 'docs/skill.md')
  requireIncludes(docsSkill, 'currently published package', 'docs/skill.md')
  requireIncludes(docsSkill, '"bilig-agent-challenge", "--json"', 'docs/skill.md')
  requireIncludes(docsSkill, '"bilig-mcp-challenge", "--json"', 'docs/skill.md')
  requireNotIncludes(docsSkill, 'allowed-tools:', 'docs/skill.md')
  requireNotIncludes(docsSkill, 'argument-hint:', 'docs/skill.md')

  requireIncludes(claudeProjectSkillNotes, 'name: bilig-workpaper', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(claudeProjectSkillNotes, '## Command Safety', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(claudeProjectSkillNotes, '"bilig-mcp-challenge", "--json"', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(claudeProjectSkillNotes, 'Return proof, not vibes.', '.claude/skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(claudeProjectSkillNotes, 'allowed-tools:', '.claude/skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(claudeProjectSkillNotes, 'argument-hint:', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(
    claudeProjectCommandNotes,
    'description: Prove workbook formula edits with Bilig WorkPaper',
    '.claude/commands/bilig-workpaper-proof.md',
  )
  requireIncludes(claudeProjectCommandNotes, '$ARGUMENTS', '.claude/commands/bilig-workpaper-proof.md')
  requireIncludes(
    claudeProjectCommandNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.claude/commands/bilig-workpaper-proof.md',
  )
  requireIncludes(claudeProjectCommandNotes, 'Do not drive', '.claude/commands/bilig-workpaper-proof.md')
  requireIncludes(claudeProjectCommandNotes, 'do not claim success from a write call alone.', '.claude/commands/bilig-workpaper-proof.md')

  requireIncludes(cursorProjectRuleNotes, 'description: Use Bilig WorkPaper', '.cursor/rules/bilig-workpaper.mdc')
  requireIncludes(cursorProjectRuleNotes, 'alwaysApply: false', '.cursor/rules/bilig-workpaper.mdc')
  requireIncludes(
    cursorProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    '.cursor/rules/bilig-workpaper.mdc',
  )
  requireIncludes(cursorProjectRuleNotes, 'Do not claim success from a write call alone.', '.cursor/rules/bilig-workpaper.mdc')
  requireIncludes(cursorProjectRuleNotes, 'https://proompteng.github.io/bilig/llms.txt', '.cursor/rules/bilig-workpaper.mdc')

  requireIncludes(windsurfProjectRuleNotes, 'trigger: model_decision', '.windsurf/rules/bilig-workpaper.md')
  requireIncludes(
    windsurfProjectRuleNotes,
    'Windsurf/Cascade agent needs spreadsheet-shaped business',
    '.windsurf/rules/bilig-workpaper.md',
  )
  requireIncludes(
    windsurfProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.windsurf/rules/bilig-workpaper.md',
  )
  requireIncludes(windsurfProjectRuleNotes, 'If any proof step fails', '.windsurf/rules/bilig-workpaper.md')
  requireIncludes(windsurfProjectRuleNotes, 'https://proompteng.github.io/bilig/llms.txt', '.windsurf/rules/bilig-workpaper.md')

  requireIncludes(clineProjectRuleNotes, 'Cline can read this workspace rule', '.clinerules/bilig-workpaper.md')
  requireIncludes(
    clineProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    '.clinerules/bilig-workpaper.md',
  )
  requireIncludes(clineProjectRuleNotes, 'Do not claim success from a write call alone.', '.clinerules/bilig-workpaper.md')
  requireIncludes(
    clineProjectRuleNotes,
    'https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html',
    '.clinerules/bilig-workpaper.md',
  )

  requireIncludes(continueProjectRuleNotes, 'name: Bilig WorkPaper Formula Proof', '.continue/rules/bilig-workpaper.md')
  requireIncludes(
    continueProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.continue/rules/bilig-workpaper.md',
  )
  requireIncludes(continueProjectRuleNotes, 'Do not claim success from a write call alone.', '.continue/rules/bilig-workpaper.md')
  requireIncludes(continueProjectRuleNotes, 'https://proompteng.github.io/bilig/llms-full.txt', '.continue/rules/bilig-workpaper.md')

  requireIncludes(copilotInstructions, '## Copilot Agent WorkPaper Path', '.github/copilot-instructions.md')
  requireIncludes(copilotInstructions, '.github/prompts/bilig-workpaper-proof.prompt.md', '.github/copilot-instructions.md')
  requireIncludes(copilotInstructions, '.vscode/mcp.json', '.github/copilot-instructions.md')
  requireIncludes(
    copilotInstructions,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.github/copilot-instructions.md',
  )
  requireIncludes(copilotInstructions, 'Do not claim success from a write call alone.', '.github/copilot-instructions.md')
  requireIncludes(copilotPrompt, 'name: bilig-workpaper-proof', '.github/prompts/bilig-workpaper-proof.prompt.md')
  requireIncludes(copilotPrompt, 'Task: ${input:task:', '.github/prompts/bilig-workpaper-proof.prompt.md')
  requireIncludes(
    copilotPrompt,
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    '.github/prompts/bilig-workpaper-proof.prompt.md',
  )
  requireIncludes(copilotPrompt, '`biligWorkpaperFile`', '.github/prompts/bilig-workpaper-proof.prompt.md')
  requireIncludes(vscodeMcpConfig, '"biligWorkpaperDemo"', '.vscode/mcp.json')
  requireIncludes(vscodeMcpConfig, '"biligWorkpaperFile"', '.vscode/mcp.json')
  requireIncludes(vscodeMcpConfig, '"${workspaceFolder}/.bilig/pricing.workpaper.json"', '.vscode/mcp.json')

  requireIncludes(rootSkillNotes, '## Command Safety', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(rootSkillNotes, 'argument array, not a shell-concatenated string', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(
    rootSkillNotes,
    'npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list',
    'skills/bilig-workpaper/SKILL.md',
  )
  requireIncludes(rootSkillNotes, 'returned `tools` array as the source', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(rootSkillNotes, 'currently published package', 'skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(rootSkillNotes, 'allowed-tools:', 'skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(rootSkillNotes, 'argument-hint:', 'skills/bilig-workpaper/SKILL.md')
}
