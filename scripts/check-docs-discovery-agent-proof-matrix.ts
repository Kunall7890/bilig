import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

export async function requireAgentProofMatrixDiscovery({
  docsRoot,
  index,
  llms,
  readme,
}: {
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly readme: string
}): Promise<void> {
  const [
    agentProofMatrixDoc,
    agentProofTranscriptsDoc,
    codexTranscriptDoc,
    claudeCodeTranscriptDoc,
    copilotTranscriptDoc,
    cursorTranscriptDoc,
    continueTranscriptDoc,
    mcpSpreadsheetFormulaServerDoc,
    aiSdkFormulaReadbackDoc,
    exceljsFormulaResultNotUpdatingDoc,
  ] = await Promise.all([
    readFile(join(docsRoot, 'agent-proof-matrix.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-proof-transcripts.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-proof-transcripts', 'codex.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-proof-transcripts', 'claude-code.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-proof-transcripts', 'copilot.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-proof-transcripts', 'cursor.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-proof-transcripts', 'continue.md'), 'utf8'),
    readFile(join(docsRoot, 'mcp-spreadsheet-formula-server-for-coding-agents.md'), 'utf8'),
    readFile(join(docsRoot, 'vercel-ai-sdk-spreadsheet-tool-formula-readback.md'), 'utf8'),
    readFile(join(docsRoot, 'exceljs-formula-result-not-updating-after-node-edits.md'), 'utf8'),
  ])

  for (const [path, content] of [
    ['docs/agent-proof-matrix.md', agentProofMatrixDoc],
    ['docs/agent-proof-transcripts.md', agentProofTranscriptsDoc],
    ['docs/agent-proof-transcripts/codex.md', codexTranscriptDoc],
    ['docs/agent-proof-transcripts/claude-code.md', claudeCodeTranscriptDoc],
    ['docs/agent-proof-transcripts/copilot.md', copilotTranscriptDoc],
    ['docs/agent-proof-transcripts/cursor.md', cursorTranscriptDoc],
    ['docs/agent-proof-transcripts/continue.md', continueTranscriptDoc],
    ['docs/mcp-spreadsheet-formula-server-for-coding-agents.md', mcpSpreadsheetFormulaServerDoc],
    ['docs/vercel-ai-sdk-spreadsheet-tool-formula-readback.md', aiSdkFormulaReadbackDoc],
    ['docs/exceljs-formula-result-not-updating-after-node-edits.md', exceljsFormulaResultNotUpdatingDoc],
  ] as const) {
    requireIncludes(content, 'image: /assets/github-social-preview.png', path)
  }

  for (const required of [
    'title: Agent WorkPaper proof matrix',
    'description: Pick the smallest Bilig WorkPaper proof for coding agents',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    '| WorkPaper service |',
    '| Agent MCP evaluator |',
    '| Provider-backed formula boundary |',
    '| Workbook Compatibility Report |',
    'bilig-evaluate --door workbook-compatibility --json',
    'noCompatibilityScore',
    '| Agent XLSX risk preflight |',
    'pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight',
    'bilig-agent-xlsx-risk-preflight.v1',
    '[Agent XLSX risk preflight](agent-xlsx-risk-preflight.md)',
    '| ExcelJS recalculation | `npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json` | `commandSucceeded: true`, `recalculationCompleted: true`, `expectedValueMatched: true` |',
    '| MCP Inspector |',
    'analyze_workbook_risk',
    'Excel compatibility certification',
    '| Vercel AI SDK `generateText()` |',
    '| Vercel AI SDK `streamText()` |',
    '| OpenAI Responses function call |',
    '| LangGraph ToolNode |',
    '| Semantic Kernel MCP plugin |',
    '| Mastra tool |',
    'listedResourcesAndPrompts',
    'restartReadbackMatchesAfter',
    'Do not duplicate that outreach.',
    '[agent proof transcripts](agent-proof-transcripts.md)',
    '[MCP spreadsheet formula server for coding agents](mcp-spreadsheet-formula-server-for-coding-agents.md)',
  ] as const) {
    requireIncludes(agentProofMatrixDoc, required, 'docs/agent-proof-matrix.md')
  }

  for (const required of [
    'title: Agent proof transcripts',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    '"@bilig/workpaper": "0.157.0"',
    '| Prompt |',
    '| Tool call |',
    '| Result |',
    '| Workbook state change |',
    '| Formula readback |',
    '| JSON export |',
    '| Restart readback verification |',
    '[Codex](agent-proof-transcripts/codex.md)',
    '[Claude Code](agent-proof-transcripts/claude-code.md)',
    '[GitHub Copilot and VS Code agent mode](agent-proof-transcripts/copilot.md)',
    '[Cursor](agent-proof-transcripts/cursor.md)',
    '[Continue](agent-proof-transcripts/continue.md)',
    '"verified": true',
    'It does not prove every desktop spreadsheet feature',
  ] as const) {
    requireIncludes(agentProofTranscriptsDoc, required, 'docs/agent-proof-transcripts.md')
  }

  for (const [path, content, hostNeedle] of [
    ['docs/agent-proof-transcripts/codex.md', codexTranscriptDoc, 'Codex should read `AGENTS.md` first'],
    ['docs/agent-proof-transcripts/claude-code.md', claudeCodeTranscriptDoc, 'Claude Code should load `CLAUDE.md`'],
    ['docs/agent-proof-transcripts/copilot.md', copilotTranscriptDoc, 'Copilot should use `.github/copilot-instructions.md`'],
    ['docs/agent-proof-transcripts/cursor.md', cursorTranscriptDoc, 'Cursor should load `.cursor/rules/bilig-workpaper.mdc`'],
    ['docs/agent-proof-transcripts/continue.md', continueTranscriptDoc, 'Continue should load `.continue/rules/bilig-workpaper.md`'],
  ] as const) {
    for (const required of [
      hostNeedle,
      '## Prompt',
      '## Tool Call',
      '## Result',
      '## Workbook State Change',
      '## Formula Readback',
      '## JSON Export',
      '## Restart Readback Verification',
      'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
      '"editedCell": "Inputs!B3"',
      '"dependentCell": "Summary!B3"',
      '"after": 96000',
      '"afterRestart": 96000',
      '"persistedDocumentBytes": 1162',
      '"restartReadbackMatchesAfter": true',
      '[Agent proof transcripts](../agent-proof-transcripts.md)',
    ] as const) {
      requireIncludes(content, required, path)
    }
  }

  for (const required of [
    'title: MCP spreadsheet formula server for coding agents',
    'https://modelcontextprotocol.io/docs/learn/server-concepts',
    'https://modelcontextprotocol.io/specification/2025-11-25/server/tools',
    'https://github.com/modelcontextprotocol/typescript-sdk',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'listedResourcesAndPrompts',
    'dependentCellChanged',
    'persistedToDisk',
    'restartReadbackMatchesAfter',
    '@modelcontextprotocol/inspector@latest',
    'set_cell_contents_and_readback',
    'https://bilig.proompteng.ai/mcp',
    '[Agent WorkPaper proof matrix](agent-proof-matrix.md)',
  ] as const) {
    requireIncludes(mcpSpreadsheetFormulaServerDoc, required, 'docs/mcp-spreadsheet-formula-server-for-coding-agents.md')
  }

  for (const required of [
    'title: Vercel AI SDK spreadsheet tool: generateText and streamText with formula readback',
    'https://ai-sdk.dev/docs/ai-sdk-core/tools-and-tool-calling',
    'pnpm --dir examples/headless-workpaper run agent:ai-sdk-generate-text',
    'AI SDK generateText -> tool -> execute',
    'pnpm --dir examples/headless-workpaper run agent:ai-sdk-stream-text',
    'AI SDK streamText -> tool -> execute',
    'streamChunkTypes',
    'createAiSdkWorkPaperTools',
    'MockLanguageModelV3',
    '[Agent WorkPaper proof matrix](agent-proof-matrix.md)',
  ] as const) {
    requireIncludes(aiSdkFormulaReadbackDoc, required, 'docs/vercel-ai-sdk-spreadsheet-tool-formula-readback.md')
  }

  for (const required of [
    'title: ExcelJS formula result not updating after Node edits',
    'https://github.com/exceljs/exceljs#formula-value',
    'npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json',
    '"commandSucceeded": true',
    '"recalculationCompleted": true',
    '"expectedValueMatched": true',
    '"value": 72000',
    'npm --prefix examples/recalc-bridge-workflows run so:exceljs-44199441',
    'recalculateExceljsWorkbook',
    'workbook.calcProperties.fullCalcOnLoad = true',
    'Use ExcelJS for workbook files and Bilig only at the recalculation boundary',
    '[ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)',
    '[Agent WorkPaper proof matrix](agent-proof-matrix.md)',
  ] as const) {
    requireIncludes(exceljsFormulaResultNotUpdatingDoc, required, 'docs/exceljs-formula-result-not-updating-after-node-edits.md')
  }
  requireNotIncludes(exceljsFormulaResultNotUpdatingDoc, '"verified": true', 'docs/exceljs-formula-result-not-updating-after-node-edits.md')
  requireNotIncludes(exceljsFormulaResultNotUpdatingDoc, '"value": 96000', 'docs/exceljs-formula-result-not-updating-after-node-edits.md')
  requireNotIncludes(exceljsFormulaResultNotUpdatingDoc, 'result.verified', 'docs/exceljs-formula-result-not-updating-after-node-edits.md')

  for (const [path, content] of [
    ['README.md', readme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
  ] as const) {
    requireIncludes(content, 'agent-proof-matrix', path)
    requireIncludes(content, 'agent-proof-transcripts', path)
    requireIncludes(content, 'mcp-spreadsheet-formula-server-for-coding-agents', path)
    requireIncludes(content, 'vercel-ai-sdk-spreadsheet-tool-formula-readback', path)
    requireIncludes(content, 'exceljs-formula-result-not-updating-after-node-edits', path)
  }
}
