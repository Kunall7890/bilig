import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { agentFrameworkLlmsRequiredLinks } from './check-docs-discovery-agent-pages.ts'
import {
  requireFile,
  requireIncludes,
  requireNoUnsupportedGoogleSheetsTenXClaims,
  requireNotIncludes,
  requirePackageKeywords,
  requirePublishedSource,
} from './check-docs-discovery-core.ts'
import { loadDocsDiscoveryContext } from './check-docs-discovery-context.ts'
import { requireSitemapPublishedSources } from './check-docs-discovery-sitemap.ts'
import { requireHomepageDiscovery } from './check-docs-discovery-homepage.ts'
import { productHuntLaunchAssetFiles } from './check-docs-discovery-launch-kit.ts'
import { requireFormulaProofDiscovery } from './check-docs-discovery-proof-pages.ts'
import { requireReadmeAgentWorkflowRecipeLinks } from './check-docs-discovery-readme-recipes.ts'
import { requireTypeScriptFirstPublicSnippets } from './check-docs-discovery-typescript-snippets.ts'
import { requireXlsxCorpusVerifierDiscovery } from './check-docs-discovery-xlsx-verifier.ts'
import { requireXlsxCalcAlternativeDiscovery } from './check-docs-discovery-xlsx-calc.ts'
import { requireXlsxRecalcPublicDiscovery } from './check-docs-discovery-xlsx-recalc-public.ts'
import { requireExternalWorkbookRecalcProofDiscovery } from './check-docs-discovery-external-workbook.ts'
import { requireSharedPublicDocsDiscovery } from './check-docs-discovery-public-docs.ts'
import { requirePackageCliSurfaceDiscovery } from './check-docs-discovery-package-cli-surfaces.ts'
import { homepageRequiredLinks, llmsRequiredLinks } from './check-docs-discovery-public-link-manifest.ts'
import { requireAgentPublicSurfaceDiscovery } from './check-docs-discovery-agent-surfaces.ts'
import { requireAgentInstructionDiscovery } from './check-docs-discovery-agent-instructions.ts'
import { requireAgentEvaluatorDiscovery } from './check-docs-discovery-agent-evaluators.ts'
import { requireTemporalWorkpaperActivityDiscovery } from './check-docs-discovery-temporal.ts'
import { requireFastMcpWorkpaperClientDiscovery } from './check-docs-discovery-fastmcp.ts'
import { requireSmolagentsWorkpaperToolDiscovery } from './check-docs-discovery-smolagents.ts'
import { requireInngestWorkpaperStepDiscovery, requireInngestWorkpaperStepExampleFiles } from './check-docs-discovery-inngest.ts'
import { requireStarterIssueDiscovery } from './check-docs-discovery-starter-issues.ts'
import { requireAgentJsonDiscoveryContract } from './agent-discovery-evaluator-doors.ts'

const docsDiscoveryContext = await loadDocsDiscoveryContext()
const skillManifestUrl = 'https://bilig.proompteng.ai/.well-known/agent-skills/bilig-workpaper/SKILL.txt'
const {
  repoRoot,
  docsRoot,
  siteRoot,
  expectedSitemapUrls,
  sourceFilesByUrl,
  benchmarkEvidence,
  headlessPackageVersion,
  readme,
  contributing,
  index,
  siteCss,
  productCss,
  robots,
  sitemap,
  llms,
  llmsFull,
  agentJson,
  agentJsonRoot,
  agentSkillsIndex,
  legacySkillsIndex,
  communityLaunchPack,
  starterIssues,
  newContributorGuide,
  headlessPackageJson,
  headlessReadme,
  excelImportReadme,
  publicApi,
  issueTemplateConfig,
  issueTemplateRoot,
  featureRequestTemplate,
  ideasDiscussionTemplate,
  qaDiscussionTemplate,
  showAndTellDiscussionTemplate,
  generalDiscussionTemplate,
  dominanceScorecard,
  xlsxFormulaRecalculationNode,
  agentXlsxFormulaRecalculationWithoutLibreOffice,
  staleXlsxFormulaCacheNode,
  microsoftGraphExcelRecalculationNode,
  formulaWorkbooksProof,
  showHnFormulaWorkbooksProof,
  googleSheetsApiBoundaryDoc,
  npmProvenancePackageTrustDoc,
  xlsxCorpusVerifierWalkthrough,
  serverSideSpreadsheetAutomationNode,
  evaluateExcelFormulasInNodeTypescript,
} = docsDiscoveryContext
const workpaperPackageSpec = '@bilig/workpaper@latest'
const fileBackedMcpArgsNeedles = [
  `"--package",\n      "${workpaperPackageSpec}",\n      "--",\n      "bilig-workpaper-mcp",\n      "--workpaper",\n      "./pricing.workpaper.json",\n      "--init-demo-workpaper",\n      "--writable"`,
  `"--package",\n        "${workpaperPackageSpec}",\n        "--",\n        "bilig-workpaper-mcp",\n        "--workpaper",\n        "./pricing.workpaper.json",\n        "--init-demo-workpaper",\n        "--writable"`,
] as const
const mcpbReleaseAssetUrl = 'https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb'
const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`
const xlsxRecalcCli = 'xlsx-recalc --demo --json'
const xlsxCacheDoctorCli = 'xlsx-cache-doctor pricing.xlsx --json'
const liveSheetjsRecalcCli = 'sheetjs-recalc --demo --json'
const liveSheetjsRecalcPackage = '@bilig/sheetjs-formula-recalc'

const headlessSpreadsheetEngineNodeServicesAgents = await readFile(
  join(docsRoot, 'headless-spreadsheet-engine-node-services-agents.md'),
  'utf8',
)
const spreadsheetMcpServerComparison = await readFile(join(docsRoot, 'spreadsheet-mcp-server-comparison.md'), 'utf8')
const mcpWorkpaperToolServerDoc = await readFile(join(docsRoot, 'mcp-workpaper-tool-server.md'), 'utf8')
const mcpServerCard = await readFile(join(docsRoot, '.well-known', 'mcp', 'server-card.json'), 'utf8')
const mcpServerCardMcpJson = await readFile(join(docsRoot, '.well-known', 'mcp.json'), 'utf8')
const mcpServerCardLegacyJson = await readFile(join(docsRoot, '.well-known', 'mcp-server-card.json'), 'utf8')
const sheetjsFormulaResultNotUpdatingNode = await readFile(join(docsRoot, 'sheetjs-formula-result-not-updating-node.md'), 'utf8')
const geminiExtensionJson = await readFile(join(repoRoot, 'gemini-extension.json'), 'utf8')
const scopedWorkpaperPackageJson = await readFile(join(repoRoot, 'packages', 'workpaper', 'package.json'), 'utf8')
const scopedWorkpaperPackageReadme = await readFile(join(repoRoot, 'packages', 'workpaper', 'README.md'), 'utf8')
const xlsxRecalcPackageJson = await readFile(join(repoRoot, 'packages', 'xlsx-formula-recalc', 'package.json'), 'utf8')
const xlsxRecalcPackageReadme = await readFile(join(repoRoot, 'packages', 'xlsx-formula-recalc', 'README.md'), 'utf8')
const xlsxRecalcPackageAgentNotes = await readFile(join(repoRoot, 'packages', 'xlsx-formula-recalc', 'AGENTS.md'), 'utf8')
const xlsxRecalcPackageSkillNotes = await readFile(join(repoRoot, 'packages', 'xlsx-formula-recalc', 'SKILL.md'), 'utf8')
const recalcBridgeReadme = await readFile(join(repoRoot, 'examples', 'recalc-bridge-workflows', 'README.md'), 'utf8')
const sheetjsRecalcPackageJson = await readFile(join(repoRoot, 'packages', 'sheetjs-formula-recalc', 'package.json'), 'utf8')
const sheetjsRecalcPackageReadme = await readFile(join(repoRoot, 'packages', 'sheetjs-formula-recalc', 'README.md'), 'utf8')
const sheetjsRecalcPackageAgentNotes = await readFile(join(repoRoot, 'packages', 'sheetjs-formula-recalc', 'AGENTS.md'), 'utf8')
const sheetjsRecalcPackageSkillNotes = await readFile(join(repoRoot, 'packages', 'sheetjs-formula-recalc', 'SKILL.md'), 'utf8')
const exceljsRecalcPackageJson = await readFile(join(repoRoot, 'packages', 'exceljs-formula-recalc', 'package.json'), 'utf8')
const exceljsRecalcPackageReadme = await readFile(join(repoRoot, 'packages', 'exceljs-formula-recalc', 'README.md'), 'utf8')
const exceljsRecalcPackageAgentNotes = await readFile(join(repoRoot, 'packages', 'exceljs-formula-recalc', 'AGENTS.md'), 'utf8')
const exceljsRecalcPackageSkillNotes = await readFile(join(repoRoot, 'packages', 'exceljs-formula-recalc', 'SKILL.md'), 'utf8')
const exceljsFormulaRecalculationNode = await readFile(join(docsRoot, 'exceljs-formula-recalculation-node.md'), 'utf8')
const directusWorkpaperFlowOperation = await readFile(join(docsRoot, 'directus-workpaper-flow-operation.md'), 'utf8')
const windmillWorkpaperScript = await readFile(join(docsRoot, 'windmill-workpaper-script.md'), 'utf8')
const triggerdevWorkpaperTask = await readFile(join(docsRoot, 'triggerdev-workpaper-task.md'), 'utf8')
const inngestWorkpaperStep = await readFile(join(docsRoot, 'inngest-workpaper-step.md'), 'utf8')
const airflowWorkpaperDag = await readFile(join(docsRoot, 'airflow-workpaper-dag.md'), 'utf8')
const dagsterWorkpaperAsset = await readFile(join(docsRoot, 'dagster-workpaper-asset.md'), 'utf8')
const kestraWorkpaperFlow = await readFile(join(docsRoot, 'kestra-workpaper-flow.md'), 'utf8')
const prefectWorkpaperFlow = await readFile(join(docsRoot, 'prefect-workpaper-flow.md'), 'utf8')
const parsedScopedWorkpaperPackageJson: unknown = JSON.parse(scopedWorkpaperPackageJson)
if (
  typeof parsedScopedWorkpaperPackageJson !== 'object' ||
  parsedScopedWorkpaperPackageJson === null ||
  Array.isArray(parsedScopedWorkpaperPackageJson)
) {
  throw new Error('packages/workpaper/package.json must be a JSON object')
}
const scopedWorkpaperPackageVersion = Reflect.get(parsedScopedWorkpaperPackageJson, 'version')
if (typeof scopedWorkpaperPackageVersion !== 'string') {
  throw new Error('packages/workpaper/package.json must define a string version')
}
const parsedGeminiExtensionJson: unknown = JSON.parse(geminiExtensionJson)
if (typeof parsedGeminiExtensionJson !== 'object' || parsedGeminiExtensionJson === null || Array.isArray(parsedGeminiExtensionJson)) {
  throw new Error('gemini-extension.json must be a JSON object')
}
if (Reflect.get(parsedGeminiExtensionJson, 'version') !== scopedWorkpaperPackageVersion) {
  throw new Error('gemini-extension.json version must match packages/workpaper/package.json')
}

requireHomepageDiscovery(index, siteCss, productCss, docsRoot)
await requireXlsxCalcAlternativeDiscovery(docsRoot)
await requireTypeScriptFirstPublicSnippets(repoRoot)
requireNoUnsupportedGoogleSheetsTenXClaims(dominanceScorecard, {
  'README.md': readme,
  'docs/index.html': index,
  'docs/google-sheets-api-alternative-node-workpaper.md': googleSheetsApiBoundaryDoc,
  'packages/headless/README.md': headlessReadme,
})
requirePackageKeywords(
  headlessPackageJson,
  [
    'agent-tools',
    'excel',
    'excel-formulas',
    'formula-recalculation',
    'formula-engine',
    'headless-spreadsheet',
    'hyperformula',
    'mcp',
    'mcp-server',
    'node',
    'spreadsheet-automation',
    'spreadsheet-engine',
    'spreadsheet-formulas',
    'spreadsheet-mcp',
    'typescript',
    'workbook-api',
    'workpaper',
    'xlsx',
  ],
  'packages/headless/package.json',
)
requirePackageKeywords(
  scopedWorkpaperPackageJson,
  [
    'agent-tools',
    'ai-agents',
    'excel-formulas',
    'exceljs',
    'formula-recalculation',
    'formula-engine',
    'mcp',
    'mcp-server',
    'model-context-protocol',
    'sheetjs',
    'spreadsheet-agent',
    'spreadsheet-automation',
    'spreadsheet-engine',
    'spreadsheet-formulas',
    'workbook-agent',
    'workbook-api',
    'workpaper',
    'xlsx',
    'xlsx-calc',
  ],
  'packages/workpaper/package.json',
)
requireIncludes(index, '"downloadUrl": "https://www.npmjs.com/package/@bilig/headless"', 'docs/index.html')
requireIncludes(index, '"https://www.npmjs.com/package/@bilig/headless"', 'docs/index.html')
requireIncludes(index, '"https://www.npmjs.com/package/@bilig/workpaper"', 'docs/index.html')
requireIncludes(index, '"https://www.npmjs.com/package/@bilig/xlsx-formula-recalc"', 'docs/index.html')
requireIncludes(index, '"https://www.npmjs.com/package/@bilig/exceljs-formula-recalc"', 'docs/index.html')
requireIncludes(index, '<h2 id="runtime-title">Teams that keep business rules in workbook formulas.</h2>', 'docs/index.html')
requireIncludes(index, 'first run', 'docs/index.html')
requireIncludes(index, 'npm install @bilig/headless', 'docs/index.html')
requireIncludes(index, '"applicationCategory": "DeveloperApplication"', 'docs/index.html')
requireIncludes(index, '"@type": "FAQPage"', 'docs/index.html')
requireIncludes(
  index,
  '<link rel="alternate" type="application/json" href="https://proompteng.github.io/bilig/.well-known/agent.json" title="agent.json" />',
  'docs/index.html',
)
for (const [mcpCardPath, mcpCardContent] of [
  ['docs/.well-known/mcp/server-card.json', mcpServerCard],
  ['docs/.well-known/mcp.json', mcpServerCardMcpJson],
  ['docs/.well-known/mcp-server-card.json', mcpServerCardLegacyJson],
] as const) {
  for (const needle of fileBackedMcpArgsNeedles) {
    requireIncludes(mcpCardContent, needle, mcpCardPath)
  }
  requireNotIncludes(mcpCardContent, '--demo-workpaper-tools', mcpCardPath)
}
requireIncludes(
  mcpWorkpaperToolServerDoc,
  `'${workpaperPackageSpec}',\n      '--',\n      'bilig-workpaper-mcp',\n      '--workpaper',\n      './pricing.workpaper.json',\n      '--init-demo-workpaper',\n      '--writable'`,
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkpaperToolServerDoc, 'Read Summary!A1:B5 with read_range.', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkpaperToolServerDoc,
  'Then set Inputs!B3 to =0.4 with set_cell_contents_and_readback.',
  'docs/mcp-workpaper-tool-server.md',
)
for (const required of homepageRequiredLinks) {
  requireIncludes(index, required, 'docs/index.html')
}

requireIncludes(robots, 'User-agent: *', 'docs/robots.txt')
requireIncludes(robots, 'Allow: /', 'docs/robots.txt')
requireIncludes(robots, `Sitemap: ${siteRoot}sitemap.xml`, 'docs/robots.txt')

const { actualSitemapUrls, sourceFilesToVerify } = requireSitemapPublishedSources({
  expectedSitemapUrls,
  sitemap,
  siteRoot,
  sourceFilesByUrl,
})

await Promise.all(sourceFilesToVerify.map((sourceFile) => requirePublishedSource(join(docsRoot, sourceFile))))
await Promise.all(
  ['README.md', 'package.json', 'quote-approval-api.ts', 'route.ts', 'smoke.ts'].map((sourceFile) =>
    requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', sourceFile)),
  ),
)
await Promise.all(
  ['README.md', 'package.json', 'tsconfig.json', 'src/api.ts', 'src/app.ts', 'src/workpaper-calculated-fields.ts', 'src/smoke.ts'].map(
    (sourceFile) => requireFile(join(repoRoot, 'examples', 'directus-workpaper-flow-operation', sourceFile)),
  ),
)
await Promise.all(
  ['README.md', 'package.json', 'tsconfig.json', 'src/workpaper-script.ts', 'src/smoke.ts'].map((sourceFile) =>
    requireFile(join(repoRoot, 'examples', 'windmill-workpaper-script', sourceFile)),
  ),
)
await Promise.all(
  ['README.md', 'package.json', 'tsconfig.json', 'src/workpaper-quote.ts', 'src/trigger-workpaper-task.ts', 'src/smoke.ts'].map(
    (sourceFile) => requireFile(join(repoRoot, 'examples', 'triggerdev-workpaper-task', sourceFile)),
  ),
)
await requireInngestWorkpaperStepExampleFiles(repoRoot)
await Promise.all(
  [
    'README.md',
    'package.json',
    'requirements.txt',
    'tsconfig.json',
    'dags/bilig_workpaper_quote_dag.py',
    'workpaper-quote.ts',
    'scripts/check-dag.ts',
  ].map((sourceFile) => requireFile(join(repoRoot, 'examples', 'airflow-workpaper-dag', sourceFile))),
)
await Promise.all(
  [
    'README.md',
    'package.json',
    'requirements.txt',
    'tsconfig.json',
    'defs/bilig_workpaper_asset.py',
    'workpaper-asset.ts',
    'scripts/check-asset.ts',
  ].map((sourceFile) => requireFile(join(repoRoot, 'examples', 'dagster-workpaper-asset', sourceFile))),
)
await Promise.all(
  ['README.md', 'package.json', 'tsconfig.json', 'flow.yml', 'blueprint.yaml', 'kestra-workpaper-flow.ts', 'scripts/check-flow.ts'].map(
    (sourceFile) => requireFile(join(repoRoot, 'examples', 'kestra-workpaper-flow', sourceFile)),
  ),
)
await Promise.all(
  ['README.md', 'package.json', 'requirements.txt', 'tsconfig.json', 'flow.py', 'workpaper-quote.ts', 'scripts/check-flow.ts'].map(
    (sourceFile) => requireFile(join(repoRoot, 'examples', 'prefect-workpaper-flow', sourceFile)),
  ),
)
await Promise.all(
  ['README.md', 'package.json', 'smoke.mjs', 'stackoverflow-exceljs-44199441.mjs', 'stackoverflow-sheetjs-63085785.mjs'].map((sourceFile) =>
    requireFile(join(repoRoot, 'examples', 'recalc-bridge-workflows', sourceFile)),
  ),
)
requireIncludes(recalcBridgeReadme, 'so:sheetjs-63085785', 'examples/recalc-bridge-workflows/README.md')
requireIncludes(recalcBridgeReadme, 'so:exceljs-44199441', 'examples/recalc-bridge-workflows/README.md')
requireIncludes(
  recalcBridgeReadme,
  'https://stackoverflow.com/questions/63085785/how-to-recalculate-all-formulas-in-excel-file-through-javascript',
  'examples/recalc-bridge-workflows/README.md',
)
requireIncludes(
  recalcBridgeReadme,
  'https://stackoverflow.com/questions/44199441/get-computed-value-of-excel-sheet-cell-in-node-js',
  'examples/recalc-bridge-workflows/README.md',
)
await requireFile(join(repoRoot, 'scripts', 'build-workpaper-mcpb.ts'))
await Promise.all(
  [
    'bilig-hero-workbook-api.png',
    'bilig-hero-workbook-api.svg',
    'bilig-hero-ambient.png',
    'hero-scene.js',
    'github-social-preview.png',
    'workpaper-benchmark-card.png',
    ...productHuntLaunchAssetFiles,
  ].map((sourceFile) => requireFile(join(docsRoot, 'assets', sourceFile))),
)
await Promise.all(
  [
    'fonts.css',
    'product-demo.css',
    'fonts/LICENSE.txt',
    'fonts/README.md',
    'fonts/ibm-plex-mono-400.woff2',
    'fonts/ibm-plex-mono-500.woff2',
    'fonts/ibm-plex-mono-600.woff2',
    'fonts/ibm-plex-sans-400.woff2',
    'fonts/ibm-plex-sans-500.woff2',
    'fonts/ibm-plex-sans-600.woff2',
    'fonts/ibm-plex-sans-700.woff2',
    'fonts/ibm-plex-sans-condensed-600.woff2',
    'fonts/ibm-plex-sans-condensed-700.woff2',
  ].map((sourceFile) => requireFile(join(docsRoot, 'assets', sourceFile))),
)

for (const required of llmsRequiredLinks) {
  requireIncludes(llms, required, 'docs/llms.txt')
}
for (const required of agentFrameworkLlmsRequiredLinks) {
  requireIncludes(llms, required, 'docs/llms.txt')
}

requireFormulaProofDiscovery({
  benchmarkEvidence,
  communityLaunchPack,
  formulaWorkbooksProof,
  headlessReadme,
  index,
  llms,
  readme,
  requireIncludes,
  showHnFormulaWorkbooksProof,
})

for (const required of [
  'title: Fix stale XLSX formula values in Node.js',
  'An `.xlsx` can store both the formula text',
  'Run a formula runtime before reading',
  'If you arrived from SheetJS, ExcelJS, or xlsx-populate',
  'I changed a cell in JavaScript; how do\nI recompute the formula value before reading it?',
  'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
  xlsxRecalcCli,
  '`@bilig/workpaper` when the service can own the workbook state locally',
  'https://github.com/proompteng/bilig/stargazers',
] as const) {
  requireIncludes(staleXlsxFormulaCacheNode, required, 'docs/stale-xlsx-formula-cache-node.md')
}
requireIncludes(index, './stale-xlsx-formula-cache-node.html', 'docs/index.html')
requireIncludes(readme, 'docs/stale-xlsx-formula-cache-node.md', 'README.md')
requireIncludes(headlessReadme, 'docs/stale-xlsx-formula-cache-node.md', 'packages/headless/README.md')
requireIncludes(llms, 'https://proompteng.github.io/bilig/stale-xlsx-formula-cache-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/stale-xlsx-formula-cache-node.md', 'docs/llms.txt')

for (const required of [
  'title: SheetJS formula result not updating in Node.js',
  'keep SheetJS for file I/O, but add a recalculation step',
  liveSheetjsRecalcCli,
  'so:sheetjs-63085785',
  'https://stackoverflow.com/questions/63085785/how-to-recalculate-all-formulas-in-excel-file-through-javascript',
  'npm --prefix examples/recalc-bridge-workflows run smoke',
  `\`${liveSheetjsRecalcPackage}\``,
  'https://github.com/proompteng/bilig/stargazers',
] as const) {
  requireIncludes(sheetjsFormulaResultNotUpdatingNode, required, 'docs/sheetjs-formula-result-not-updating-node.md')
}
requireIncludes(index, './sheetjs-formula-result-not-updating-node.html', 'docs/index.html')
requireIncludes(readme, 'docs/sheetjs-formula-result-not-updating-node.md', 'README.md')
requireIncludes(headlessReadme, 'docs/sheetjs-formula-result-not-updating-node.md', 'packages/headless/README.md')
requireIncludes(llms, 'https://proompteng.github.io/bilig/sheetjs-formula-result-not-updating-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/sheetjs-formula-result-not-updating-node.md', 'docs/llms.txt')

for (const required of [
  'title: Microsoft Graph Excel recalculation vs local Node WorkPaper',
  'POST /me/drive/items/{id}/workbook/application/calculate',
  'Files.ReadWrite',
  'application permissions are not supported for that API',
  'Use `@bilig/workpaper` when the workbook is service-owned state',
  'https://learn.microsoft.com/en-us/graph/api/workbookapplication-calculate',
  'https://github.com/proompteng/bilig/stargazers',
] as const) {
  requireIncludes(microsoftGraphExcelRecalculationNode, required, 'docs/microsoft-graph-excel-recalculation-node.md')
}
requireIncludes(index, './microsoft-graph-excel-recalculation-node.html', 'docs/index.html')
requireIncludes(readme, 'docs/microsoft-graph-excel-recalculation-node.md', 'README.md')
requireIncludes(headlessReadme, 'docs/microsoft-graph-excel-recalculation-node.md', 'packages/headless/README.md')
requireIncludes(llms, 'https://proompteng.github.io/bilig/microsoft-graph-excel-recalculation-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/microsoft-graph-excel-recalculation-node.md', 'docs/llms.txt')

await requireSharedPublicDocsDiscovery({
  docsRoot,
  readme,
  headlessReadme,
  contributing,
  newContributorGuide,
  starterIssues,
  llms,
  index,
  issueTemplateConfig,
  issueTemplateRoot,
  featureRequestTemplate,
  ideasDiscussionTemplate,
  qaDiscussionTemplate,
  showAndTellDiscussionTemplate,
  generalDiscussionTemplate,
  excelImportReadme,
  publicApi,
})

requireIncludes(readme, 'acceptance commands for first patches.', 'README.md')
requireIncludes(readme, 'docs/why-use-bilig.md', 'README.md')
await requireAgentEvaluatorDiscovery({ docsRoot, readme, index, llms, runtimePackageVersion: scopedWorkpaperPackageVersion })
requireIncludes(llms, '## agent handoff prompt', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/AGENTS.md', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/agent.json', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/agent.json', 'docs/llms.txt')
requireIncludes(llms, skillManifestUrl, 'docs/llms.txt')
requireNotIncludes(llms, 'https://proompteng.github.io/bilig/skill.txt', 'docs/llms.txt')
requireNotIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/agent-skills/bilig-workpaper/SKILL.txt', 'docs/llms.txt')
requireNotIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/skills/bilig-workpaper/SKILL.txt', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.claude/skills/bilig-workpaper/SKILL.md', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.claude/commands/bilig-workpaper-proof.md', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.github/copilot-instructions.md', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.github/prompts/bilig-workpaper-proof.prompt.md', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.vscode/mcp.json', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.cursor/rules/bilig-workpaper.mdc', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.windsurf/rules/bilig-workpaper.md', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.clinerules/bilig-workpaper.md', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/.continue/rules/bilig-workpaper.md', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/llms-full.txt', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/agent-skills/index.json', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/skills/index.json', 'docs/llms.txt')
requireIncludes(llms, '## evaluator doors', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/eval-xlsx-cache-doctor.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/eval-xlsx-recalc.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/eval-workpaper-service.html', 'docs/llms.txt')
requireIncludes(llms, 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-challenge --json', 'docs/llms.txt')
requireIncludes(llms, 'npm exec --yes --package @bilig/workpaper@latest -- bilig-mcp-challenge --json', 'docs/llms.txt')
requireIncludes(llms, 'A write call by itself is not success.', 'docs/llms.txt')
requireIncludes(llms, 'low-code agent workflow formula readback:', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/windmill-workpaper-script.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/triggerdev-workpaper-task.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/temporal-workpaper-activity.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/airflow-workpaper-dag.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/dagster-workpaper-asset.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/kestra-workpaper-flow.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/prefect-workpaper-flow.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html', 'docs/llms.txt')
requireIncludes(llms, 'n8n community node package: @bilig/n8n-nodes-workpaper', 'docs/llms.txt')
requireIncludes(llms, 'https://www.npmjs.com/package/@bilig/n8n-nodes-workpaper', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html', 'docs/llms.txt')
requireIncludes(llms, `npm exec --package ${workpaperPackageSpec} -- bilig-n8n-formula-server --port 4321`, 'docs/llms.txt')
requireIncludes(readme, '@bilig/n8n-nodes-workpaper', 'README.md')
requireIncludes(readme, 'Directus Persisted Calculated Fields', 'README.md')
requireIncludes(readme, 'Trigger.dev Durable Formula Tasks', 'README.md')
requireIncludes(readme, 'Dagster Formula Assets', 'README.md')
requireIncludes(
  directusWorkpaperFlowOperation,
  'Run Script executes in an isolated sandbox without access to npm modules',
  'docs/directus-workpaper-flow-operation.md',
)
requireIncludes(directusWorkpaperFlowOperation, 'examples/directus-workpaper-flow-operation', 'docs/directus-workpaper-flow-operation.md')
requireIncludes(windmillWorkpaperScript, 'examples/windmill-workpaper-script', 'docs/windmill-workpaper-script.md')
requireIncludes(windmillWorkpaperScript, 'Windmill can infer inputs from the `main` parameters', 'docs/windmill-workpaper-script.md')
requireIncludes(triggerdevWorkpaperTask, 'examples/triggerdev-workpaper-task', 'docs/triggerdev-workpaper-task.md')
requireIncludes(triggerdevWorkpaperTask, 'task({ id, run })', 'docs/triggerdev-workpaper-task.md')
requireIncludes(triggerdevWorkpaperTask, 'Trigger.dev owns durable execution.', 'docs/triggerdev-workpaper-task.md')
requireInngestWorkpaperStepDiscovery({ inngestWorkpaperStep, llms, llmsFull, scopedWorkpaperPackageReadme })
await requireTemporalWorkpaperActivityDiscovery({ repoRoot, docsRoot, llmsFull, scopedWorkpaperPackageReadme })
requireIncludes(airflowWorkpaperDag, 'examples/airflow-workpaper-dag', 'docs/airflow-workpaper-dag.md')
requireIncludes(airflowWorkpaperDag, 'from airflow.sdk import dag, task', 'docs/airflow-workpaper-dag.md')
requireIncludes(airflowWorkpaperDag, 'npx --no-install tsx workpaper-quote.ts', 'docs/airflow-workpaper-dag.md')
requireIncludes(
  airflowWorkpaperDag,
  'Airflow owns scheduling, retries, dependency graph state, XCom summary',
  'docs/airflow-workpaper-dag.md',
)
requireIncludes(dagsterWorkpaperAsset, 'examples/dagster-workpaper-asset', 'docs/dagster-workpaper-asset.md')
requireIncludes(dagsterWorkpaperAsset, 'PipesSubprocessClient', 'docs/dagster-workpaper-asset.md')
requireIncludes(dagsterWorkpaperAsset, 'report_asset_materialization', 'docs/dagster-workpaper-asset.md')
requireIncludes(
  dagsterWorkpaperAsset,
  'Dagster owns orchestration, asset state, run history, and materialization',
  'docs/dagster-workpaper-asset.md',
)
requireIncludes(kestraWorkpaperFlow, 'examples/kestra-workpaper-flow', 'docs/kestra-workpaper-flow.md')
requireIncludes(kestraWorkpaperFlow, 'blueprint.yaml', 'docs/kestra-workpaper-flow.md')
requireIncludes(kestraWorkpaperFlow, 'io.kestra.plugin.scripts.node.Commands', 'docs/kestra-workpaper-flow.md')
requireIncludes(kestraWorkpaperFlow, 'Kestra owns orchestration and output-file routing.', 'docs/kestra-workpaper-flow.md')
requireIncludes(prefectWorkpaperFlow, 'examples/prefect-workpaper-flow', 'docs/prefect-workpaper-flow.md')
requireIncludes(prefectWorkpaperFlow, 'from prefect import flow, task', 'docs/prefect-workpaper-flow.md')
requireIncludes(prefectWorkpaperFlow, 'Prefect owns orchestration and task history.', 'docs/prefect-workpaper-flow.md')
requireIncludes(llmsFull, '@bilig/n8n-nodes-workpaper', 'docs/llms-full.txt')
requireIncludes(llmsFull, 'Directus WorkPaper Flow Operation', 'docs/llms-full.txt')
requireIncludes(llmsFull, 'Windmill WorkPaper TypeScript script', 'docs/llms-full.txt')
requireIncludes(llmsFull, 'Trigger.dev WorkPaper task', 'docs/llms-full.txt')
requireIncludes(llmsFull, 'Airflow WorkPaper DAG', 'docs/llms-full.txt')
requireIncludes(llmsFull, 'Dagster WorkPaper Asset', 'docs/llms-full.txt')
requireIncludes(llmsFull, 'Kestra WorkPaper Node flow', 'docs/llms-full.txt')
requireIncludes(llmsFull, 'Prefect WorkPaper flow', 'docs/llms-full.txt')
requireIncludes(llmsFull, '## Agent Adoption Kit', 'docs/llms-full.txt')
requireIncludes(readme, 'docs/.well-known/agent.json', 'README.md')
requireIncludes(headlessReadme, 'https://proompteng.github.io/bilig/.well-known/agent.json', 'packages/headless/README.md')
requireIncludes(scopedWorkpaperPackageReadme, '## Start Here', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Windmill TypeScript workflow fields', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Trigger.dev durable task fields', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Apache Airflow DAG task outputs', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Dagster asset materialization metadata', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Kestra Node Commands flow fields', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Prefect flow fields', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Directus Flow operation for persisted calculated fields', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'n8n formula readback for self-hosted workflows', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Dify formula readback', 'packages/workpaper/README.md')
requireIncludes(scopedWorkpaperPackageReadme, 'Pipedream formula readback', 'packages/workpaper/README.md')
requireIncludes(llmsFull, 'Pipedream WorkPaper Formula Readback', 'docs/llms-full.txt')
requireIncludes(readme, 'docs/pipedream-workpaper-formula-readback.md', 'README.md')
requireReadmeAgentWorkflowRecipeLinks(readme)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/triggerdev-workpaper-task.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/airflow-workpaper-dag.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/dagster-workpaper-asset.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/kestra-workpaper-flow.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/prefect-workpaper-flow.html',
  'packages/workpaper/README.md',
)
await requireFastMcpWorkpaperClientDiscovery({ repoRoot, docsRoot, index, llms, llmsFull, scopedWorkpaperPackageReadme })
await requireSmolagentsWorkpaperToolDiscovery({ repoRoot, docsRoot, index, llms, llmsFull, scopedWorkpaperPackageReadme })
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/directus-workpaper-flow-operation.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html',
  'packages/workpaper/README.md',
)
requireIncludes(
  scopedWorkpaperPackageReadme,
  'https://proompteng.github.io/bilig/pipedream-workpaper-formula-readback.html',
  'packages/workpaper/README.md',
)
requireIncludes(llms, 'Do not claim success from a write call alone.', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir bilig/examples/headless-workpaper install --ignore-workspace', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir bilig/examples/headless-workpaper run agent:framework-adapters', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir examples/headless-workpaper run agent:mcp-tools', 'docs/llms.txt')
requireNotIncludes(llms, 'cd bilig/examples/headless-workpaper', 'docs/llms.txt')
requireNotIncludes(llms, '\nnpm start\n', 'docs/llms.txt')
requireIncludes(headlessReadme, 'https://proompteng.github.io/bilig/why-use-bilig.html', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'The npm tarball also includes `AGENTS.md`', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'SKILL.md', 'packages/headless/README.md')
requireIncludes(headlessPackageJson, '"AGENTS.md"', 'packages/headless/package.json')
requireIncludes(headlessPackageJson, '"SKILL.md"', 'packages/headless/package.json')
await requireAgentInstructionDiscovery({ repoRoot, docsRoot, headlessPackageVersion })
if (agentJsonRoot !== agentJson) {
  throw new Error('docs/agent.json must match docs/.well-known/agent.json')
}
const parsedAgentJson: unknown = JSON.parse(agentJson)
if (typeof parsedAgentJson !== 'object' || parsedAgentJson === null || Array.isArray(parsedAgentJson)) {
  throw new Error('docs/.well-known/agent.json must be a JSON object')
}
requireAgentJsonDiscoveryContract({
  parsedAgentJson,
  repositoryUrl: 'https://github.com/proompteng/bilig',
  siteRoot,
  skillManifestUrl,
  workpaperPackageSpec,
})
const parsedAgentJsonMcp = Reflect.get(parsedAgentJson, 'mcp')
if (typeof parsedAgentJsonMcp !== 'object' || parsedAgentJsonMcp === null || Array.isArray(parsedAgentJsonMcp)) {
  throw new Error('docs/.well-known/agent.json must define an mcp object')
}
if (Reflect.get(parsedAgentJsonMcp, 'server_card') !== 'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json') {
  throw new Error('docs/.well-known/agent.json must point at the MCP server card')
}
if (Reflect.get(parsedAgentJsonMcp, 'remote_endpoint') !== 'https://bilig.proompteng.ai/mcp') {
  throw new Error('docs/.well-known/agent.json must advertise the hosted MCP endpoint')
}
const agentJsonMcpRemoteTransport = Reflect.get(parsedAgentJsonMcp, 'remote_transport')
if (
  typeof agentJsonMcpRemoteTransport !== 'object' ||
  agentJsonMcpRemoteTransport === null ||
  Reflect.get(agentJsonMcpRemoteTransport, 'type') !== 'streamable-http' ||
  Reflect.get(agentJsonMcpRemoteTransport, 'protocol_version') !== '2025-11-25'
) {
  throw new Error('docs/.well-known/agent.json must advertise the hosted Streamable HTTP MCP transport')
}
const agentJsonMcpTools = Reflect.get(parsedAgentJsonMcp, 'tools')
if (!Array.isArray(agentJsonMcpTools) || !agentJsonMcpTools.every((tool) => typeof tool === 'string')) {
  throw new Error('docs/.well-known/agent.json mcp.tools must be a string array')
}
for (const requiredTool of [
  'list_sheets',
  'set_cell_contents',
  'set_cell_contents_and_readback',
  'get_cell_display_value',
  'export_workpaper_document',
  'validate_formula',
]) {
  if (!agentJsonMcpTools.includes(requiredTool)) {
    throw new Error(`docs/.well-known/agent.json mcp.tools is missing ${requiredTool}`)
  }
}
const agentJsonMcpResources = Reflect.get(parsedAgentJsonMcp, 'resources')
if (!Array.isArray(agentJsonMcpResources) || !agentJsonMcpResources.every((resource) => typeof resource === 'string')) {
  throw new Error('docs/.well-known/agent.json mcp.resources must be a string array')
}
for (const requiredResource of ['bilig://workpaper/manifest', 'bilig://workpaper/agent-handoff', 'bilig://workpaper/current-document']) {
  if (!agentJsonMcpResources.includes(requiredResource)) {
    throw new Error(`docs/.well-known/agent.json mcp.resources is missing ${requiredResource}`)
  }
}
const agentJsonMcpPrompts = Reflect.get(parsedAgentJsonMcp, 'prompts')
if (!Array.isArray(agentJsonMcpPrompts) || !agentJsonMcpPrompts.every((prompt) => typeof prompt === 'string')) {
  throw new Error('docs/.well-known/agent.json mcp.prompts must be a string array')
}
for (const requiredPrompt of ['edit_and_verify_workpaper', 'debug_workpaper_formula']) {
  if (!agentJsonMcpPrompts.includes(requiredPrompt)) {
    throw new Error(`docs/.well-known/agent.json mcp.prompts is missing ${requiredPrompt}`)
  }
}
const agentJsonCapabilities = Reflect.get(parsedAgentJson, 'capabilities')
if (
  !Array.isArray(agentJsonCapabilities) ||
  !agentJsonCapabilities.some(
    (capability) =>
      typeof capability === 'object' &&
      capability !== null &&
      Reflect.get(capability, 'name') === 'file-backed-workpaper-mcp' &&
      Reflect.get(capability, 'server_card') === 'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
  )
) {
  throw new Error('docs/.well-known/agent.json must advertise the file-backed MCP capability')
}
if (
  !agentJsonCapabilities.some(
    (capability) =>
      typeof capability === 'object' &&
      capability !== null &&
      Reflect.get(capability, 'name') === 'remote-workpaper-mcp-demo' &&
      Reflect.get(capability, 'endpoint') === 'https://bilig.proompteng.ai/mcp',
  )
) {
  throw new Error('docs/.well-known/agent.json must advertise the remote MCP demo capability')
}
if (
  !agentJsonCapabilities.some(
    (capability) =>
      typeof capability === 'object' &&
      capability !== null &&
      Reflect.get(capability, 'name') === 'claude-desktop-mcpb' &&
      Reflect.get(capability, 'type') === 'mcpb-desktop-extension' &&
      Reflect.get(capability, 'download_url') === mcpbReleaseAssetUrl &&
      Reflect.get(capability, 'checksum_url') === mcpbReleaseChecksumUrl,
  )
) {
  throw new Error('docs/.well-known/agent.json must advertise the Claude Desktop MCPB release asset')
}
const agentJsonPublicEntrypoints = Reflect.get(parsedAgentJson, 'public_entrypoints')
if (!Array.isArray(agentJsonPublicEntrypoints) || !agentJsonPublicEntrypoints.every((entrypoint) => typeof entrypoint === 'string')) {
  throw new Error('docs/.well-known/agent.json public_entrypoints must be a string array')
}
for (const requiredEntrypoint of [
  'https://proompteng.github.io/bilig/openai-agents-sdk-workpaper-tool.html',
  'https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html',
  'https://proompteng.github.io/bilig/llamaindex-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/crewai-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/cloudflare-agents-workpaper-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/semantic-kernel-workpaper-mcp.html',
  'https://proompteng.github.io/bilig/gemini-cli-workpaper-extension.html',
  'https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html',
  'https://proompteng.github.io/bilig/dify-workpaper-formula-readback.html',
  'https://proompteng.github.io/bilig/flowise-workpaper-formula-readback.html',
  'https://proompteng.github.io/bilig/triggerdev-workpaper-task.html',
  'https://proompteng.github.io/bilig/temporal-workpaper-activity.html',
]) {
  if (!agentJsonPublicEntrypoints.includes(requiredEntrypoint)) {
    throw new Error(`docs/.well-known/agent.json public_entrypoints is missing ${requiredEntrypoint}`)
  }
}
requireIncludes(
  agentSkillsIndex,
  '"$schema": "https://schemas.agentskills.io/discovery/0.2.0/schema.json"',
  'docs/.well-known/agent-skills/index.json',
)
requireIncludes(agentSkillsIndex, '"type": "skill-md"', 'docs/.well-known/agent-skills/index.json')
requireIncludes(agentSkillsIndex, '"digest": "sha256:', 'docs/.well-known/agent-skills/index.json')
requireIncludes(
  agentSkillsIndex,
  '"url": "https://bilig.proompteng.ai/.well-known/agent-skills/bilig-workpaper/SKILL.txt"',
  'docs/.well-known/agent-skills/index.json',
)
requireIncludes(
  legacySkillsIndex,
  '"$schema": "https://schemas.agentskills.io/discovery/0.2.0/schema.json"',
  'docs/.well-known/skills/index.json',
)
requireIncludes(legacySkillsIndex, '"type": "skill-md"', 'docs/.well-known/skills/index.json')
requireIncludes(legacySkillsIndex, '"digest": "sha256:', 'docs/.well-known/skills/index.json')
requireIncludes(
  legacySkillsIndex,
  '"url": "https://bilig.proompteng.ai/.well-known/agent-skills/bilig-workpaper/SKILL.txt"',
  'docs/.well-known/skills/index.json',
)
requireIncludes(llmsFull, '## Generated Skill Manifest', 'docs/llms-full.txt')
requireIncludes(llmsFull, '## Headless WorkPaper Agent Handbook', 'docs/llms-full.txt')
requireIncludes(llmsFull, `npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge`, 'docs/llms-full.txt')
requireIncludes(llmsFull, `npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp`, 'docs/llms-full.txt')
requireIncludes(headlessReadme, '## Stay Connected', 'packages/headless/README.md')
requireIncludes(headlessReadme, '## More Guides', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'Pick a scoped first patch:', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'When the sanity check passes, these are the next useful pages.', 'packages/headless/README.md')

requireStarterIssueDiscovery({ contributing, llms, newContributorGuide, starterIssues })
requireIncludes(
  evaluateExcelFormulasInNodeTypescript,
  'npx tsx eval-node-formulas.ts',
  'docs/evaluate-excel-formulas-in-node-typescript.md',
)
requireIncludes(serverSideSpreadsheetAutomationNode, 'npx tsx eval.ts', 'docs/server-side-spreadsheet-automation-node.md')
for (const required of [
  'title: Google Sheets API alternative for local Node workbook execution',
  'That is the boundary. `bilig` is not trying to replace Google Sheets.',
  'npm install @bilig/headless',
  'npx tsx eval.ts',
  '"verified": true',
  'https://developers.google.com/workspace/sheets/api/guides/concepts',
  'https://developers.google.com/workspace/sheets/api/guides/values',
] as const) {
  requireIncludes(googleSheetsApiBoundaryDoc, required, 'docs/google-sheets-api-alternative-node-workpaper.md')
}
requireIncludes(readme, 'Google Sheets API boundary', 'README.md')
requireIncludes(headlessReadme, 'Google Sheets API boundary', 'packages/headless/README.md')
requireIncludes(index, './google-sheets-api-alternative-node-workpaper.html', 'docs/index.html')
requireIncludes(llms, 'https://proompteng.github.io/bilig/google-sheets-api-alternative-node-workpaper.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/google-sheets-api-alternative-node-workpaper.md', 'docs/llms.txt')

for (const required of [
  'title: Verify npm provenance for @bilig/headless',
  'npm view @bilig/headless@latest version dist.attestations dist.signatures --json',
  'npm audit signatures',
  'dist.attestations.provenance.predicateType',
  'npm publish ... --provenance',
  'https://docs.npmjs.com/trusted-publishers/',
  'https://docs.npmjs.com/viewing-package-provenance/',
  'https://scorecard.dev/',
  'official OpenSSF Scorecard action',
  'uploaded as SARIF to GitHub code',
] as const) {
  requireIncludes(npmProvenancePackageTrustDoc, required, 'docs/npm-provenance-package-trust.md')
}
requireIncludes(readme, `@bilig/headless@${headlessPackageVersion}`, 'README.md')
requireIncludes(readme, 'npm view @bilig/headless version dist.attestations dist.signatures --json', 'README.md')
requireIncludes(readme, 'npm provenance and package trust', 'README.md')
requireIncludes(readme, 'https://api.scorecard.dev/projects/github.com/proompteng/bilig/badge', 'README.md')
requireIncludes(readme, 'uploaded to GitHub code scanning on every `main` update', 'README.md')
requireIncludes(headlessReadme, `@bilig/headless@${headlessPackageVersion}`, 'packages/headless/README.md')
requireIncludes(
  headlessReadme,
  'npm view @bilig/headless@latest version dist.attestations dist.signatures --json',
  'packages/headless/README.md',
)
requireIncludes(headlessReadme, 'npm provenance and package trust guide', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'https://api.scorecard.dev/projects/github.com/proompteng/bilig/badge', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'uploaded to GitHub code scanning on every `main` update', 'packages/headless/README.md')
requirePackageCliSurfaceDiscovery({
  exceljsRecalcPackageAgentNotes,
  exceljsRecalcPackageJson,
  exceljsRecalcPackageReadme,
  exceljsRecalcPackageSkillNotes,
  scopedWorkpaperPackageJson,
  scopedWorkpaperPackageReadme,
  sheetjsRecalcPackageAgentNotes,
  sheetjsRecalcPackageJson,
  sheetjsRecalcPackageReadme,
  sheetjsRecalcPackageSkillNotes,
  xlsxRecalcPackageAgentNotes,
  xlsxRecalcPackageJson,
  xlsxRecalcPackageReadme,
  xlsxRecalcPackageSkillNotes,
})
requireIncludes(exceljsFormulaRecalculationNode, 'exceljs-recalc --demo --json', 'docs/exceljs-formula-recalculation-node.md')
requireIncludes(exceljsFormulaRecalculationNode, 'so:exceljs-44199441', 'docs/exceljs-formula-recalculation-node.md')
requireIncludes(
  exceljsFormulaRecalculationNode,
  'https://stackoverflow.com/questions/44199441/get-computed-value-of-excel-sheet-cell-in-node-js',
  'docs/exceljs-formula-recalculation-node.md',
)
requireIncludes(
  llms,
  'https://github.com/proompteng/bilig/blob/main/examples/recalc-bridge-workflows/stackoverflow-sheetjs-63085785.mjs',
  'docs/llms.txt',
)
requireIncludes(
  llms,
  'https://github.com/proompteng/bilig/blob/main/examples/recalc-bridge-workflows/stackoverflow-exceljs-44199441.mjs',
  'docs/llms.txt',
)
requireXlsxRecalcPublicDiscovery({
  agentXlsxFormulaRecalculationWithoutLibreOffice,
  headlessReadme,
  index,
  liveSheetjsRecalcCli,
  llms,
  readme,
  xlsxFormulaRecalculationNode,
  xlsxCacheDoctorCli,
  xlsxRecalcCli,
})
await requireExternalWorkbookRecalcProofDiscovery({ docsRoot, index, llms, xlsxRecalcPackageReadme })
requireIncludes(llms, 'https://proompteng.github.io/bilig/agent-xlsx-formula-recalculation-without-libreoffice.html', 'docs/llms.txt')
requireIncludes(
  llms,
  'https://github.com/proompteng/bilig/blob/main/docs/agent-xlsx-formula-recalculation-without-libreoffice.md',
  'docs/llms.txt',
)
requireIncludes(llms, 'gives spreadsheet agents a Node.js tool contract', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/excel-file-calculation-engine-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/excel-file-calculation-engine-node.md', 'docs/llms.txt')
requireIncludes(llms, 'covers backend routes that write request inputs into an XLSX workbook', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/exceljs-shared-formula-recalculation-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/exceljs-shared-formula-recalculation-node.md', 'docs/llms.txt')
requireIncludes(llms, 'documents the XLSX shared-formula expansion path', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-template-formula-recalculation-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-template-formula-recalculation-node.md', 'docs/llms.txt')
requireIncludes(llms, 'template substitution -> formula runtime -> verified readback', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-populate-formula-result-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-populate-formula-result-node.md', 'docs/llms.txt')
requireIncludes(llms, 'separates formula serialization from recalculation', 'docs/llms.txt')
requireIncludes(index, './npm-provenance-package-trust.html', 'docs/index.html')
requireIncludes(llms, 'https://proompteng.github.io/bilig/npm-provenance-package-trust.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/npm-provenance-package-trust.md', 'docs/llms.txt')
await requireFile(join(repoRoot, '.github', 'workflows', 'scorecard.yml'))

requireXlsxCorpusVerifierDiscovery(xlsxCorpusVerifierWalkthrough)
requireIncludes(index, './xlsx-corpus-verifier-walkthrough.html', 'docs/index.html')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-corpus-verifier-walkthrough.html', 'docs/llms.txt')

await requireAgentPublicSurfaceDiscovery({
  context: docsDiscoveryContext,
  headlessSpreadsheetEngineNodeServicesAgents,
  spreadsheetMcpServerComparison,
})

console.log(
  JSON.stringify(
    {
      ok: true,
      sitemapUrlCount: actualSitemapUrls.length,
      robots: 'ok',
      llms: 'ok',
    },
    null,
    2,
  ),
)
