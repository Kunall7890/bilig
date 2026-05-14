import type { Buffer } from 'node:buffer'
import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const previewWidth = 1280
const previewHeight = 640
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsRoot = join(repoRoot, 'docs', 'assets')
const outputPath = join(assetsRoot, 'github-social-preview.png')
const checkMode = process.argv.includes('--check')

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function fontFace(family: string, weight: number, data: Buffer): string {
  return String.raw`
    @font-face {
      font-family: '${family}';
      font-weight: ${weight};
      src: url(data:font/woff2;base64,${data.toString('base64')}) format('woff2');
    }`
}

function execFileBuffer(file: string, args: readonly string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      [...args],
      {
        encoding: 'buffer',
        maxBuffer: 16 * 1024 * 1024,
      },
      (error, stdout, stderr) => {
        if (error !== null) {
          const message = Buffer.isBuffer(stderr) ? stderr.toString('utf8') : String(stderr)
          reject(new Error(`${file} failed: ${message.trim() || error.message}`))
          return
        }
        resolve(Buffer.from(stdout))
      },
    )
  })
}

async function buildSvg(): Promise<string> {
  const sansRegular = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-400.woff2'))
  const sansMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-500.woff2'))
  const sansSemiBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-600.woff2'))
  const sansBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-700.woff2'))
  const monoMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-mono-500.woff2'))

  const codeLines = [
    "import { WorkPaper } from '@bilig/headless';",
    '',
    'const workbook = WorkPaper.build({',
    "  sheet: 'Revenue',",
    '  rows: [',
    "    ['West', 32, 1200, '=B2*C2'],",
    "    ['East', 30, 250, '=B3*C3'],",
    '  ],',
    '});',
    '',
    'workbook.setCellContents(cell, 32);',
    'const total = workbook.getCellValue(sum);',
  ]
  const code = codeLines
    .map(
      (line, index) =>
        `<text x="958" y="${194 + index * 22}" class="mono" font-size="11" font-weight="500" fill="${line.startsWith('const') ? '#d8f5df' : '#d9d2c3'}">${escapeXml(line)}</text>`,
    )
    .join('\n')

  return String.raw`<svg xmlns="http://www.w3.org/2000/svg" width="${previewWidth}" height="${previewHeight}" viewBox="0 0 ${previewWidth} ${previewHeight}">
  <defs>
    <style>
      ${fontFace('Bilig Sans', 400, sansRegular)}
      ${fontFace('Bilig Sans', 500, sansMedium)}
      ${fontFace('Bilig Sans', 600, sansSemiBold)}
      ${fontFace('Bilig Sans', 700, sansBold)}
      ${fontFace('Bilig Mono', 500, monoMedium)}
      text {
        font-family: 'Bilig Sans', Arial, sans-serif;
        letter-spacing: 0;
      }
      .mono {
        font-family: 'Bilig Mono', Menlo, monospace;
      }
    </style>
    <linearGradient id="page" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#15140f"/>
      <stop offset="1" stop-color="#202719"/>
    </linearGradient>
    <linearGradient id="sheet" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#f7f5ef"/>
      <stop offset="1" stop-color="#edf4ea"/>
    </linearGradient>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="22" stdDeviation="28" flood-color="#000000" flood-opacity="0.28"/>
    </filter>
  </defs>

  <rect width="${previewWidth}" height="${previewHeight}" fill="url(#page)"/>
  <g opacity="0.24" stroke="#4c493e" stroke-width="1">
    <path d="M0 78 H1280"/>
    <path d="M0 156 H1280"/>
    <path d="M0 234 H1280"/>
    <path d="M0 312 H1280"/>
    <path d="M0 390 H1280"/>
    <path d="M0 468 H1280"/>
    <path d="M0 546 H1280"/>
    <path d="M96 0 V640"/>
    <path d="M192 0 V640"/>
    <path d="M288 0 V640"/>
    <path d="M384 0 V640"/>
    <path d="M480 0 V640"/>
    <path d="M576 0 V640"/>
    <path d="M672 0 V640"/>
    <path d="M768 0 V640"/>
    <path d="M864 0 V640"/>
    <path d="M960 0 V640"/>
    <path d="M1056 0 V640"/>
    <path d="M1152 0 V640"/>
  </g>

  <g transform="translate(72 74)">
    <rect x="0" y="0" width="58" height="58" rx="14" fill="#f5f0e6"/>
    <rect x="14" y="14" width="12" height="12" rx="3" fill="#1f8a54"/>
    <rect x="32" y="14" width="12" height="12" rx="3" fill="#1f8a54"/>
    <rect x="14" y="32" width="12" height="12" rx="3" fill="#1f8a54"/>
    <rect x="32" y="32" width="12" height="12" rx="3" fill="#1f8a54"/>
    <text x="76" y="39" fill="#f5f0e6" font-size="34" font-weight="700">bilig</text>
  </g>

  <text x="72" y="190" fill="#32d179" font-size="19" font-weight="700">@bilig/headless</text>
  <text x="72" y="254" fill="#f5f0e6" font-size="50" font-weight="700">Spreadsheet formulas</text>
  <text x="72" y="309" fill="#f5f0e6" font-size="50" font-weight="700">for TypeScript services</text>
  <text x="72" y="360" fill="#cfc7b7" font-size="23" font-weight="400">Build a workbook, change cells, read the result,</text>
  <text x="72" y="392" fill="#cfc7b7" font-size="23" font-weight="400">and persist the same model as JSON.</text>
  <text x="72" y="424" fill="#cfc7b7" font-size="23" font-weight="400">No grid required.</text>

  <g transform="translate(72 462)">
    <rect x="0" y="0" width="492" height="54" fill="#201f1a" stroke="#615b4e"/>
    <rect x="0" y="0" width="58" height="54" fill="#25241f" stroke="#615b4e"/>
    <text x="25" y="35" fill="#32d179" class="mono" font-size="21" font-weight="500">$</text>
    <text x="78" y="35" fill="#f5f0e6" class="mono" font-size="21" font-weight="500">npm install @bilig/headless</text>
  </g>

  <g transform="translate(72 548)" fill="#cfc7b7">
    <text x="0" y="0" fill="#f5f0e6" font-size="20" font-weight="700">TypeScript examples</text>
    <text x="0" y="30" font-size="17" font-weight="400">real .ts files</text>
    <path d="M188 -24 V38" stroke="#615b4e"/>
    <text x="212" y="0" fill="#f5f0e6" font-size="20" font-weight="700">MCP ready</text>
    <text x="212" y="30" font-size="17" font-weight="400">stdio server</text>
    <path d="M354 -24 V38" stroke="#615b4e"/>
    <text x="378" y="0" fill="#f5f0e6" font-size="20" font-weight="700">46/46 mean rows</text>
    <text x="378" y="30" font-size="17" font-weight="400">checked benchmark</text>
  </g>

  <g filter="url(#shadow)">
    <rect x="650" y="82" width="556" height="468" rx="16" fill="#12110e" stroke="#696154"/>
  </g>
  <rect x="650" y="82" width="556" height="42" rx="16" fill="#252119"/>
  <path d="M650 108 H1206" stroke="#696154"/>
  <text x="672" y="109" fill="#d9d1c1" font-size="16" font-weight="700">Revenue.workpaper</text>
  <text x="1042" y="109" fill="#d9d1c1" font-size="15" font-weight="700">verified readback</text>

  <rect x="672" y="150" width="250" height="260" fill="url(#sheet)" stroke="#d1d8cd"/>
  <rect x="672" y="150" width="250" height="43" fill="#e9f0e5" stroke="#d1d8cd"/>
  <text x="698" y="177" fill="#687064" font-size="16" font-weight="700">A</text>
  <text x="777" y="177" fill="#687064" font-size="16" font-weight="700">B</text>
  <text x="856" y="177" fill="#687064" font-size="16" font-weight="700">C</text>
  <path d="M748 150 V410 M828 150 V410 M672 236 H922 M672 279 H922 M672 322 H922 M672 365 H922" stroke="#d1d8cd"/>
  <text x="690" y="222" fill="#1f201c" font-size="18">Region</text>
  <text x="762" y="222" fill="#1f201c" font-size="18">ARPA</text>
  <text x="843" y="222" fill="#1f201c" font-size="18">Revenue</text>
  <text x="690" y="265" fill="#1f201c" font-size="19">West</text>
  <text x="774" y="265" fill="#1f201c" font-size="19">1200</text>
  <text x="846" y="265" fill="#12794a" font-size="19" font-weight="700">38,400</text>
  <text x="690" y="308" fill="#1f201c" font-size="19">East</text>
  <text x="786" y="308" fill="#1f201c" font-size="19">250</text>
  <text x="858" y="308" fill="#1f201c" font-size="19">7,500</text>
  <text x="690" y="351" fill="#1f201c" font-size="19">Total</text>
  <text x="848" y="351" fill="#12794a" font-size="20" font-weight="700">51,300</text>

  <rect x="944" y="150" width="240" height="260" rx="10" fill="#151f27" stroke="#354657"/>
  <rect x="944" y="150" width="240" height="38" rx="10" fill="#1b2834"/>
  <text x="962" y="175" fill="#8ab4f8" class="mono" font-size="14" font-weight="500">tool.ts</text>
  ${code}

  <g transform="translate(672 438)">
    <text x="0" y="0" fill="#cfc7b7" font-size="19">after restore</text>
    <text x="0" y="48" fill="#32d179" font-size="47" font-weight="700">51,300</text>
    <text x="198" y="48" fill="#32d179" font-size="25" font-weight="700">verified</text>
  </g>

  <text x="1206" y="594" text-anchor="end" fill="#cfc7b7" font-size="19">github.com/proompteng/bilig</text>
</svg>`
}

async function renderPreview(): Promise<Buffer> {
  const svg = await buildSvg()
  const tempRoot = await mkdtemp(join(tmpdir(), 'bilig-social-preview-'))
  const svgPath = join(tempRoot, 'preview.svg')
  try {
    await writeFile(svgPath, svg)
    return await execFileBuffer('rsvg-convert', [
      '--format=png',
      '--width',
      String(previewWidth),
      '--height',
      String(previewHeight),
      svgPath,
    ])
  } finally {
    await rm(tempRoot, { recursive: true, force: true })
  }
}

const image = await renderPreview()

if (checkMode) {
  const existing = await readFile(outputPath)
  if (!existing.equals(image)) {
    throw new Error(`${outputPath} is stale. Run pnpm docs:social-preview:generate.`)
  }
  console.log(`social preview is current: ${outputPath}`)
} else {
  await mkdir(dirname(outputPath), { recursive: true })
  await writeFile(outputPath, image)
  console.log(`wrote ${outputPath}`)
}
