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
  const [agentProofMatrixDoc, mcpSpreadsheetFormulaServerDoc, aiSdkFormulaReadbackDoc, exceljsFormulaResultNotUpdatingDoc] =
    await Promise.all([
      readFile(join(docsRoot, 'agent-proof-matrix.md'), 'utf8'),
      readFile(join(docsRoot, 'mcp-spreadsheet-formula-server-for-coding-agents.md'), 'utf8'),
      readFile(join(docsRoot, 'vercel-ai-sdk-spreadsheet-tool-formula-readback.md'), 'utf8'),
      readFile(join(docsRoot, 'exceljs-formula-result-not-updating-after-node-edits.md'), 'utf8'),
    ])

  for (const [path, content] of [
    ['docs/agent-proof-matrix.md', agentProofMatrixDoc],
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
    '| XLSX cache detector |',
    '| ExcelJS recalculation | `npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json` | `commandSucceeded: true`, `recalculationCompleted: true`, `expectedValueMatched: true` |',
    '| MCP Inspector |',
    '| Vercel AI SDK `generateText()` |',
    '| Vercel AI SDK `streamText()` |',
    '| OpenAI Responses function call |',
    '| LangGraph ToolNode |',
    '| Semantic Kernel MCP plugin |',
    '| Mastra tool |',
    'listedResourcesAndPrompts',
    'restartReadbackMatchesAfter',
    'Do not duplicate that outreach.',
    '[MCP spreadsheet formula server for coding agents](mcp-spreadsheet-formula-server-for-coding-agents.md)',
  ] as const) {
    requireIncludes(agentProofMatrixDoc, required, 'docs/agent-proof-matrix.md')
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
    requireIncludes(content, 'mcp-spreadsheet-formula-server-for-coding-agents', path)
    requireIncludes(content, 'vercel-ai-sdk-spreadsheet-tool-formula-readback', path)
    requireIncludes(content, 'exceljs-formula-result-not-updating-after-node-edits', path)
  }
}
