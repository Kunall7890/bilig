import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { getBenchmarkDiscoveryEvidence } from './check-docs-discovery-benchmark-evidence.ts'
import { docsSiteSources } from './check-docs-discovery-site-sources.ts'

function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

function requireNotIncludes(haystack: string, needle: string, context: string): void {
  if (haystack.includes(needle)) {
    throw new Error(`${context} must not include ${needle}`)
  }
}

function requireAllIncludes(haystack: string, needles: string[], context: string): void {
  for (const needle of needles) {
    requireIncludes(haystack, needle, context)
  }
}

function requireNoIncludes(haystack: string, needles: string[], context: string): void {
  for (const needle of needles) {
    requireNotIncludes(haystack, needle, context)
  }
}

function requireHomepageLocalHtmlLinksHaveSources(index: string, docsRoot: string): void {
  const sourceFilesByPage = new Map(docsSiteSources.map(([urlPath, sourceFile]) => [urlPath, sourceFile]))
  const missingSources: string[] = []
  const localHtmlHrefs = Array.from(index.matchAll(/href="([^"]+\.html(?:#[^"]*)?)"/g))
    .map((match) => match[1])
    .filter((href) => href.startsWith('./') || href.startsWith('/bilig/'))

  for (const href of localHtmlHrefs) {
    const page = href
      .replace(/^\.\//, '')
      .replace(/^\/bilig\//, '')
      .split('#')[0]
    const sourceFile = sourceFilesByPage.get(page)

    if (sourceFile === undefined) {
      missingSources.push(`${href}: missing from docsSiteSources`)
      continue
    }

    const sourcePath = join(docsRoot, sourceFile)
    if (!existsSync(sourcePath)) {
      missingSources.push(`${href}: missing source docs/${sourceFile}`)
      continue
    }

    if (page.endsWith('.html') && sourceFile.endsWith('.md')) {
      const sourceText = readFileSync(sourcePath, 'utf8')
      if (!sourceText.startsWith('---\n')) {
        missingSources.push(`${href}: docs/${sourceFile} is missing YAML front matter, so GitHub Pages will not emit ${page}`)
      }
    }
  }

  if (missingSources.length > 0) {
    throw new Error(`docs/index.html has local HTML links without published sources:\n${missingSources.join('\n')}`)
  }
}

export function requireHomepageDiscovery(index: string, siteCss: string, productCss: string, docsRoot: string): void {
  const benchmarkEvidence = getBenchmarkDiscoveryEvidence()

  requireHomepageLocalHtmlLinksHaveSources(index, docsRoot)

  requireAllIncludes(
    index,
    [
      '<link rel="canonical" href="https://proompteng.github.io/bilig/" />',
      '<link rel="sitemap" type="application/xml" href="https://proompteng.github.io/bilig/sitemap.xml" />',
      '<link rel="help" href="https://context7.com/proompteng/bilig" title="Ask Bilig docs on Context7" />',
      '<link rel="alternate" type="text/plain" href="https://proompteng.github.io/bilig/llms.txt" title="llms.txt" />',
      '"@type": "SoftwareSourceCode"',
      '"codeRepository": "https://github.com/proompteng/bilig"',
      '<title>bilig - XLSX Cache Doctor for Node and CI</title>',
      '<meta name="robots" content="index, follow, max-image-preview:large" />',
      '<link rel="icon" type="image/svg+xml" href="./assets/favicon.svg" />',
      '<meta property="og:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-15-2" />',
      '<meta property="og:image:alt" content="bilig XLSX cache doctor formula readback preview" />',
      '<meta name="twitter:card" content="summary_large_image" />',
      '<meta name="twitter:image" content="https://proompteng.github.io/bilig/assets/github-social-preview.png?v=2026-05-15-2" />',
      '<meta name="twitter:image:alt" content="bilig XLSX cache doctor formula readback preview" />',
      '<link rel="stylesheet" href="./assets/fonts.css?v=2026-05-14-1" />',
      '<link rel="stylesheet" href="./assets/site.css?v=2026-05-30-10" />',
      '<link rel="stylesheet" href="./assets/product-demo.css?v=2026-05-15-3" />',
      '<script src="./assets/site-nav.js?v=2026-05-30-10"></script>',
      '<script type="module" src="./assets/hero-scene.js?v=2026-05-15-1"></script>',
      '<p class="eyebrow">XLSX cache doctor for Node and CI</p>',
      '<h1 id="hero-title" class="hero-title">Find stale Excel formula values before production reads them.</h1>',
      'Run a local check or a read-only GitHub Action',
      'https://github.com/marketplace/actions/xlsx-cache-doctor',
      './stale-xlsx-formula-cache-node.html',
      './eval-xlsx-cache-doctor.html',
      './eval-xlsx-recalc.html',
      './eval-workpaper-service.html',
      './eval-agent-mcp.html',
      './agent-adoption-kit.html',
      'npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor --demo --json',
      'See a stale cached value and the recalculated value without cloning the repo.',
      "workbooks: '**/*.xlsx'",
      'fail-on-stale',
      'Need a fresh workbook artifact after the detector finds the cells?',
      '<a href="#packages">Start</a>',
      '<a href="#install">Node API</a>',
      '<a href="./xlsx-cache-doctor-github-action.html">Action</a>',
      '<a href="./xlsx-formula-recalculation-node.html">XLSX</a>',
      '<a href="#docs">Docs</a>',
      'stale XLSX cache evaluator',
      '<a class="button primary" href="https://github.com/proompteng/xlsx-cache-doctor-demo/pull/1">See demo PR</a>',
      '<a class="button" href="./xlsx-cache-doctor-github-action.html#generate-the-workflow">Copy workflow</a>',
      '<a class="button" href="https://github.com/marketplace/actions/xlsx-cache-doctor">Use the Action</a>',
      '<strong>1 stale formula cache</strong>',
      '<dt>Cell</dt>',
      '<dd><code>Sheet1!B61</code></dd>',
      '<h2 id="packages-title">Start with the file before the runtime.</h2>',
      "const canonicalDocsRoot = 'https://proompteng.github.io/bilig/'",
      "const isRawFilePreview = window.location.protocol === 'file:'",
      'const isCanonicalGeneratedDocsHost',
      "window.location.hostname === 'proompteng.github.io' && window.location.pathname.startsWith('/bilig/')",
      'document.querySelectorAll',
      'assets/site-nav.js',
      'isGeneratedDocsHref',
      'markdownHrefForLocalSource',
      'canonicalHrefForGeneratedDocs',
      'rewriteLinks',
      'if (isRawFilePreview) {',
      'rewriteLinks(markdownHrefForLocalSource)',
      'if (!isCanonicalGeneratedDocsHost) {',
      'rewriteLinks(canonicalHrefForGeneratedDocs)',
      'new URL(href.slice(2), canonicalDocsRoot).toString()',
      "href.replace(/\\.html(?=$|[?#])/, '.md')",
      '<code>npm install @bilig/xlsx-formula-recalc</code>',
      '<code>xlsx-recalc --read Summary!B7</code>',
      'bilig-hero-ambient.png?v=2026-05-15-1',
      'One stale formula value should be obvious.',
      'The detector names the workbook, sheet, cell, cached value, recalculated value, and next readback command.',
      'List formula cells.',
      'Compare the saved cache.',
      'Name the exact cell.',
      'Call the detector from Node.',
      'Use the same stale-cache report as the CLI',
      'npm install @bilig/xlsx-formula-recalc',
      'inspectXlsxCache',
      'npx tsx quickstart.ts',
      'Coding agent? Do not drive Excel, LibreOffice, Google Sheets, or a browser grid.',
      'bilig-mcp-challenge --json',
      'Put the workbook behind the code that owns the workflow.',
      'Route handler, queue worker, CLI, or MCP server: load the workbook',
      'The benchmark is public. So is the caveat.',
      'The checked scorecard has bilig ahead across HyperFormula, TrueCalc, Univer, xlsx-calc, IronCalc Rust.',
      'Coverage tiers and unsupported rows stay visible.',
      'headless-performance-leadership-scorecard.json',
      '5 engines',
      'mean+p95 scorecard',
      'pnpm headless:performance:check',
      `<strong>${benchmarkEvidence.meanWinHeadline}</strong>`,
      `${benchmarkEvidence.allProviderMeanAndP95Headline} comparable workloads win on both mean and p95 in the checked leadership scorecard.`,
      `${benchmarkEvidence.meanWinSentencePrefix} comparable mean-latency rows are faster in the checked HyperFormula file.`,
      `${benchmarkEvidence.p95HoldoutWorkload} is the current worst p95 row:`,
      'rendering is not part of this benchmark.',
      `<code>${benchmarkEvidence.p95HoldoutRatio}</code>`,
      'UI rendering, Excel file compatibility, and workbook shapes this suite does not cover.',
      'packages/benchmarks/baselines/<wbr />headless-performance-leadership-scorecard.json',
      'Benchmark notes',
      'Compatibility gaps',
      './production-adoption-checklist-headless-workpaper.html',
      'Starter issues',
      'benchmark critique',
      'https://github.com/proompteng/bilig/blob/main/SECURITY.md',
      'https://github.com/proompteng/bilig/blob/main/SUPPORT.md',
      'Pick the path that matches the job.',
      'Try the package, put it behind a service route, expose it as an agent tool',
      'Take a starter issue that improves the examples.',
      'Start with code/test picks, example tasks, adapters, or focused docs coverage.',
      '<h3>Run</h3>',
      '<h3>Build</h3>',
      '<h3>Agents</h3>',
      'Context7 indexed docs',
      '<h3>Decide</h3>',
      'Try it on the spreadsheet that is already slowing you down.',
    ],
    'docs/index.html',
  )

  requireAllIncludes(
    siteCss,
    [
      "--font-body: 'Bilig Sans'",
      "--font-display: 'Bilig Sans'",
      "--font-mono: 'Bilig Mono'",
      'grid-template-columns: minmax(440px, 0.78fr) minmax(520px, 1fr);',
      'grid-template-columns: minmax(360px, 0.92fr) minmax(0, 1.08fr);',
      'min-width: 0;',
      'section[id] {\n  scroll-margin-top:',
      'scroll-padding-top: calc(var(--topbar-height) + 18px);',
      '@media (max-width: 389px) {\n  :root {\n    --topbar-height: 126px;',
      'flex-wrap: nowrap;',
      'overflow-x: auto;',
      '.nav a {\n  display: inline-flex;',
      '.hero-command',
      '.hero-note',
      'overflow-wrap: anywhere;',
      '@media (max-width: 1180px) {',
      'grid-template-columns: repeat(3, minmax(0, 1fr));',
      '.link-stack .path-link em {\n    grid-column: auto;',
      '.proof-result {\n  display: grid;\n  grid-template-columns: minmax(0, 1fr);',
      'width: fit-content;',
      '.proof-metric {\n  display: grid;',
      '.proof-result p {\n  min-width: 0;',
      '.proof-facts',
      '.page-main',
      '.markdown-page',
    ],
    'docs/assets/site.css',
  )

  requireAllIncludes(
    productCss,
    [
      '.hero-media {\n  position: relative;',
      '.hero-ambient,',
      '.hero-canvas,',
      '.hero-fallback {',
      '.action-report-card {',
      '.report-card-top strong {',
      '.hero-media::after',
      '.action-report-card dl {',
    ],
    'docs/assets/product-demo.css',
  )

  requireNoIncludes(
    index,
    [
      'bilig-hero-workbook-api.png?v=2026-05-08-2',
      'bilig-hero-workbook-api.png?v=2026-05-14-4',
      'bilig-hero-workbook-api.png?v=2026-05-14-6',
      'localPreviewHostnames',
      'isLocalHttpPreview',
      'Run the smoke test',
      '<h1 id="hero-title" class="hero-title" translate="no">bilig</h1>',
      'Headless workbooks for TypeScript services.',
      'Keep the spreadsheet shape when cells and formulas are still the clearest way to explain the calculation.',
      'Run the maintained eval before reading more docs.',
      'The smoke test changes one input',
      'TypeScript smoke test',
      '>eval.ts<',
      'curl -fsSLo eval.ts',
      'npx tsx eval.ts',
      'npm-only smoke test',
      'npm smoke test feedback',
      'If the smoke test matches your workflow',
      'TypeScript package',
      'Examples are real <code>.ts</code> files with public imports.',
      'The npm package ships a stdio server for agent tools.',
      'The HyperFormula comparison links to the JSON artifact and caveat.',
      'github-social-preview.png?v=2026-05-08-2',
      'github-social-preview.png?v=2026-05-14-5',
      "window.location.hostname === 'localhost'",
      "window.location.hostname === '127.0.0.1'",
      "window.location.hostname === '::1'",
      'localPreviewHosts',
      'isLocalStaticPreview',
      'const localHtmlLinks',
      'const probeAndRewriteMissingLocalHtml',
      'publicDocsHostnames',
      'isHostedSourcePreview',
      'isSourcePreview',
      'Run workbook formulas in TypeScript.',
      'Use a WorkPaper when a calculation is easier to review as cells and formulas.',
      'Runnable .ts examples',
      'The docs link to real TypeScript files.',
      'Set an input cell and read the recalculated total.',
      'no browser required',
      'Use it when the spreadsheet is the model.',
      'Pricing tables, import checks, payouts, and budget rules often make the most sense as rows and formulas.',
      'Run the smoke test in a blank Node project.',
      'Put the workbook where the calculation happens.',
      'Here is the benchmark. Here is the caveat.',
      'The checked JSON shows bilig ahead on every comparable mean-latency row.',
      'It does not claim every workbook will be faster.',
      'Pick the path that matches your codebase.',
      'Docs by job, not package name.',
      'Spreadsheet logic for TypeScript services.',
      'When a pricing rule, budget check, or payout model is easiest to express as cells and formulas',
      'The docs link to runnable TypeScript files.',
      'Edit, then read',
      'Change one input and inspect the dependent total.',
      'Small docs, examples, adapters, and tests are open.',
      'Run the TypeScript example',
      'Build a workbook in Node, change inputs through code',
      'For pricing, budgets, imports, payouts, and agent tools that still work like a workbook.',
      'Build the sheet, write the input,',
      'Use it for pricing calculators, budget checks, imports, payouts, and agent tools when the formula should stay in a workbook.',
      'Build the sheet, edit the input, read the cell, save JSON.',
      'Examples are `.ts` files with real imports.',
      'Change B2 and read the dependent total.',
      'Small first PRs are kept open.',
      'The workbook is data plus formulas. The code edits it and reads the result.',
      'Workbook formulas for TypeScript services.',
      'Keep a calculation as cells and formulas when that is the clearest model.',
      'Copy the `.ts` files and run them.',
      'Change an input cell and read the dependent total.',
      'One workbook object. TypeScript changes it and reads the value it produced.',
      'Use a workbook when formulas are the clearest source of truth.',
      'Use a workbook when the formula is the thing you ship.',
      'If a model is already easier to review in rows and formulas',
      'Change the input.',
      'Run this before reading more.',
      'The smoke test changes one customer count',
      'Use it from the place that already owns the workflow.',
      'Node route, queue worker, CLI, or MCP server: load the workbook',
      'The benchmark claim is narrow on purpose.',
      'The checked-in WorkPaper vs HyperFormula artifact says bilig is faster on mean latency',
      'It does not say every spreadsheet, p95 row, or Excel file is faster.',
      'The number is public. So is the caveat.',
      'The current WorkPaper vs HyperFormula artifact has WorkPaper ahead on mean latency',
      'One approximate-lookup duplicate case is slower at p95, and that stays visible.',
      'Every comparable mean-latency row in the checked-in artifact is green.',
      'not a blanket',
      'promise for every workbook.',
      'Lower mean latency on every comparable workload in the current artifact.',
      'The approximate-lookup duplicate case is slower',
      'Mean latency is green for the suite; one p95 row is not.',
      'Browser grid performance, Excel compatibility, and your own workbook shape.',
      'That is the current WorkPaper vs HyperFormula file in the repo.',
      'It covers shared headless formula workloads,',
      'compatibility, browser rendering, or your own workbook.',
      'The benchmark artifact includes the caveat.',
      'No hidden browser, no screen scraping',
      'Run the benchmark, then decide.',
      'bilig is ahead on this checked-in WorkPaper suite.',
      'good news for this suite, not a blank check',
      'Run the benchmark before you depend on it.',
      'Fast where the benchmark says so. Clear where it does not.',
      'Speed claims are cheap. Run the benchmark.',
      'The benchmark command and JSON artifact are in the repo.',
      'Before you quote the number, run it.',
      'The result, baseline JSON, and p95 caveat are all in the repo.',
      'The benchmark lives in the repo.',
      'Run the command, read the JSON, and check the slower row before you use the number in your own decision.',
      'The checked-in benchmark compares WorkPaper against HyperFormula-style workloads.',
      'One p95 caveat is listed beside it.',
      'Comparable benchmark rows only, measured by mean latency.',
      '<dd>46 / 46 ahead</dd>',
      'The JSON artifact changes in review when the benchmark moves.',
      'run the benchmark locally before depending on the headline row.',
      'Coverage and caveats',
      'Excel behavior not covered yet',
      'Small TypeScript tasks',
      'The speed claim is deliberately narrow',
      'Run it yourself.',
      'the slower p95 row named beside the result.',
      'Comparable benchmark rows only. Treat the number like a benchmark, not a slogan.',
      'Mean latency only. This is not a promise that every workbook, formula, or p95 row is faster.',
      '<dd>46 / 46 comparable mean rows</dd>',
      'The checked-in JSON is the thing to inspect in review, not a screenshot of the number.',
      'The artifact lives in the repo so benchmark drift shows up in review.',
      'That slower row stays visible because it matters if your workload looks like it.',
      'Rerun the command before using the number in your own docs.',
      'Artifact summary, caveat, and rerun instructions.',
      'Run the benchmark yourself.',
      'Do not use this as a blanket speed claim.',
      'Pick the closest starting point.',
      'Docs, examples, issues.',
      'Small docs, examples, and integration tasks for first PRs.',
      'Known gaps are documented in public before they become surprises.',
      'Everything public, including the rough edges.',
      'Try the package before you trust the page.',
      'Public project numbers',
      '<span>Stars</span>',
      '<strong>24</strong>',
      'No trust-me homepage claims',
      'trust-me',
      'What this benchmark says.',
      '<dd>46 / 46</dd>',
      'Mean latency wins on the comparable rows in the current WorkPaper vs HyperFormula suite.',
      'CI checks this file, so benchmark drift shows up as a normal review diff.',
      'The approximate-lookup duplicate case is slower at p95. Benchmark your own workbook if that pattern matters.',
      'Read the benchmark notes',
      'Pick a starter issue',
      'One claim, with the caveat beside it.',
      'The benchmark is public. So are the gaps.',
      'The current WorkPaper vs HyperFormula artifact puts WorkPaper ahead on mean latency for this suite.',
      'Current checked-in result',
      '<strong>46 / 46</strong>',
      'Lower mean latency on every comparable row in the current artifact.',
      'Comparable headless workloads',
      'No UI-performance claim. No Excel-compatibility claim.',
      'Committed JSON, not a screenshot. Benchmark changes show up as normal diffs.',
      'Read what the numbers mean',
      'Check the Excel gaps',
      'Good first issues',
      'Public project signals',
      '<strong>40 starter tasks</strong>',
      '90 starter issues',
      '105 starter issues',
      '<strong>0.13.9</strong>',
      'Read those before you depend on the package.',
      'launch essay',
      '<a class="button" href="./eval-workpaper-service.html">Build an API</a>',
      '<a class="button" href="./eval-agent-mcp.html">Use with agents</a>',
      '<title>bilig - Formulas for TypeScript</title>',
      '<h1 id="hero-title" class="hero-title">Formulas for TypeScript.</h1>',
      '<title>bilig - Formula workbooks for Node services</title>',
      '<h1 id="hero-title" class="hero-title">Formula workbooks for Node services.</h1>',
      '"https://dev.to/gregkonush/why-agents-need-workbook-apis-instead-of-spreadsheet-screenshots-3d61"',
    ],
    'docs/index.html',
  )

  requireNotIncludes(siteCss, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'grid-template-columns: minmax(150px, 0.34fr) minmax(0, 1fr);', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'grid-template-columns: minmax(210px, max-content) minmax(0, 1fr);', 'docs/assets/site.css')
  requireNotIncludes(siteCss, 'border-left: 1px solid rgba(255, 250, 240, 0.16);', 'docs/assets/site.css')
}
