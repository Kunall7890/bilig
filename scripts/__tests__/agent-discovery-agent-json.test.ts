import { describe, expect, it } from 'vitest'

import { buildAgentJsonManifest } from '../agent-discovery-agent-json.ts'

const manifestInput = {
  siteRoot: 'https://proompteng.github.io/bilig',
  repositoryUrl: 'https://github.com/proompteng/bilig',
  skillDiscoveryRoot: 'https://bilig.proompteng.ai',
  skillManifestUrl: 'https://bilig.proompteng.ai/.well-known/agent-skills/bilig-workpaper/SKILL.txt',
  skillName: 'bilig-workpaper',
  workpaperPackageSpec: '@bilig/workpaper@latest',
  remoteMcpEndpoint: 'https://bilig.proompteng.ai/mcp',
  remoteMcpAliasEndpoint: 'https://bilig.proompteng.ai/mcp/workpaper',
  remoteMcpServerCard: 'https://bilig.proompteng.ai/.well-known/mcp/server-card.json',
  mcpbReleaseAssetUrl: 'https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb',
  mcpbReleaseChecksumUrl: 'https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb.sha256',
} as const

describe('agent discovery agent.json manifest', () => {
  it('emits repo-formatted JSON for compact primitive arrays', () => {
    const manifest = buildAgentJsonManifest(manifestInput)

    expect(manifest).toContain('"hosts": ["Codex", "Claude Code", "GitHub Copilot", "Cursor", "Continue"]')
    expect(JSON.parse(manifest)).toMatchObject({
      schema_version: 'agent-json-0.1.0',
      name: 'bilig',
    })
  })
})
