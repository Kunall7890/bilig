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
    const parsed = JSON.parse(manifest)
    expect(parsed).toMatchObject({
      schema_version: 'agent-json-0.1.0',
      name: 'bilig',
      title: 'Bilig WorkPaper spreadsheet formula readback',
      description:
        'Spreadsheet formula readback for Node.js services and agent MCP tools: edit cells, recalculate, verify outputs, and persist JSON without UI automation.',
      keywords: expect.arrayContaining(['spreadsheet formula readback', 'MCP spreadsheet tools', 'agent workbook automation']),
      tags: expect.arrayContaining(['spreadsheet-formula-readback', 'mcp-spreadsheet-tools', 'agent-workbook-automation']),
    })
    expect(parsed.capabilities).toContainEqual(
      expect.objectContaining({
        name: 'workbook-compatibility-risk-report',
        type: 'local-cli-evaluator',
        command: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json',
        expected_result:
          'bilig-evaluator.v1 JSON with workbook risk reasons, unsupported functions, external links, VBA payloads, pivots, volatile functions, stale cache counts, no compatibility score, and verified true',
        boundary:
          'Diagnoses workbook risks before agent or service use; does not certify Excel compatibility, execute macros, refresh pivots or external data, or assign a compatibility percentage.',
      }),
    )
    expect(parsed.mcp).toMatchObject({
      xlsx_import_tools: ['analyze_workbook_risk'],
      xlsx_import_args: ['exec', '--package', '@bilig/workpaper@latest', '--', 'bilig-workpaper-mcp', '--from-xlsx', './pricing.xlsx'],
    })
    expect(parsed.capabilities).toContainEqual(
      expect.objectContaining({
        name: 'file-backed-workpaper-mcp',
        xlsx_import_tool: 'analyze_workbook_risk',
        boundary:
          'The XLSX import tool is local and fixed to the --from-xlsx source file; it reports workbook risk indicators and does not certify Excel compatibility. Without --workpaper --writable, edits stay in memory.',
      }),
    )
  })
})
