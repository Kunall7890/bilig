import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

export function requirePackageCliSurfaceDiscovery(args: {
  readonly scopedXlsxRecalcPackageAgentNotes: string
  readonly scopedXlsxRecalcPackageJson: string
  readonly scopedXlsxRecalcPackageReadme: string
  readonly scopedXlsxRecalcPackageSkillNotes: string
  readonly exceljsRecalcPackageAgentNotes: string
  readonly exceljsRecalcPackageJson: string
  readonly exceljsRecalcPackageReadme: string
  readonly exceljsRecalcPackageSkillNotes: string
  readonly sheetjsRecalcPackageAgentNotes: string
  readonly sheetjsRecalcPackageJson: string
  readonly sheetjsRecalcPackageReadme: string
  readonly sheetjsRecalcPackageSkillNotes: string
  readonly scopedWorkpaperPackageJson: string
  readonly scopedWorkpaperPackageReadme: string
  readonly xlsxRecalcPackageAgentNotes: string
  readonly xlsxRecalcPackageJson: string
  readonly xlsxRecalcPackageReadme: string
  readonly xlsxRecalcPackageSkillNotes: string
}): void {
  requireIncludes(args.scopedWorkpaperPackageJson, '"bilig-agent-start": "./bin/bilig-agent-start.js"', 'packages/workpaper/package.json')
  requireIncludes(args.scopedWorkpaperPackageJson, '"bilig-evaluate": "./bin/bilig-evaluate.js"', 'packages/workpaper/package.json')
  requireIncludes(args.scopedWorkpaperPackageJson, '"@bilig/xlsx-formula-recalc": "workspace:*"', 'packages/workpaper/package.json')
  requireIncludes(args.scopedWorkpaperPackageJson, 'Run workbook-shaped business rules in Node services', 'packages/workpaper/package.json')
  requireIncludes(args.scopedWorkpaperPackageJson, 'workbook-runtime', 'packages/workpaper/package.json')
  requireIncludes(args.scopedWorkpaperPackageJson, 'workpaper-json', 'packages/workpaper/package.json')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'bilig-evaluate --door workpaper-service --json', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'bilig-evaluate --door agent-mcp --json', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'bilig-agent-start --json', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'compact routing card with proof commands', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'workbook-shaped business logic', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'bilig-evaluator.v1', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, '## What Success Looks Like', 'packages/workpaper/README.md')
  requireIncludes(
    args.scopedWorkpaperPackageReadme,
    'bilig-evaluate --door agent-mcp --scenario revenue-plan --json',
    'packages/workpaper/README.md',
  )
  requireIncludes(args.scopedWorkpaperPackageReadme, '"door": "workpaper-service"', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, '"editedCell": "Inputs!B2"', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, '"dependentCell": "Summary!B2"', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, '"afterRestore": 38400', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'https://github.com/proompteng/bilig', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'https://github.com/proompteng/bilig/subscription', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'Use `bilig-mcp-challenge --json` only', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'analyze_workbook_risk', 'packages/workpaper/README.md')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'does not certify Excel compatibility', 'packages/workpaper/README.md')

  requireIncludes(
    args.scopedXlsxRecalcPackageJson,
    '"workbook-compatibility-report": "./bin/workbook-compatibility-report.js"',
    'packages/bilig-xlsx-formula-recalc/package.json',
  )
  requireIncludes(args.scopedXlsxRecalcPackageJson, '"AGENTS.md"', 'packages/bilig-xlsx-formula-recalc/package.json')
  requireIncludes(args.scopedXlsxRecalcPackageJson, '"SKILL.md"', 'packages/bilig-xlsx-formula-recalc/package.json')
  requireIncludes(
    args.scopedXlsxRecalcPackageAgentNotes,
    'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json',
    'packages/bilig-xlsx-formula-recalc/AGENTS.md',
  )
  requireIncludes(
    args.scopedXlsxRecalcPackageAgentNotes,
    'workbook-compatibility-report workbook.xlsx --json',
    'packages/bilig-xlsx-formula-recalc/AGENTS.md',
  )
  requireIncludes(
    args.scopedXlsxRecalcPackageSkillNotes,
    'bilig-evaluate --door workbook-compatibility --json',
    'packages/bilig-xlsx-formula-recalc/SKILL.md',
  )
  requireIncludes(
    args.scopedXlsxRecalcPackageSkillNotes,
    "from '@bilig/xlsx-formula-recalc'",
    'packages/bilig-xlsx-formula-recalc/SKILL.md',
  )
  requireIncludes(
    args.scopedXlsxRecalcPackageReadme,
    'workbook-compatibility-report pricing.xlsx --json',
    'packages/bilig-xlsx-formula-recalc/README.md',
  )

  requireIncludes(args.xlsxRecalcPackageJson, '"bilig-evaluate": "./bin/bilig-evaluate.js"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(
    args.xlsxRecalcPackageJson,
    '"workbook-compatibility-report": "./bin/workbook-compatibility-report.js"',
    'packages/xlsx-formula-recalc/package.json',
  )
  requireIncludes(args.xlsxRecalcPackageJson, '"xlsx-recalc": "./bin/xlsx-recalc.js"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(
    args.xlsxRecalcPackageJson,
    '"xlsx-cache-doctor": "./bin/xlsx-cache-doctor.js"',
    'packages/xlsx-formula-recalc/package.json',
  )
  requireIncludes(args.xlsxRecalcPackageJson, '"sheetjs-recalc": "./bin/sheetjs-recalc.js"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(args.xlsxRecalcPackageJson, '"./workbook-compatibility-report"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(args.xlsxRecalcPackageJson, '"./evaluator"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(args.xlsxRecalcPackageJson, '"./cli-api"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(
    args.xlsxRecalcPackageReadme,
    'bilig-evaluate --door workbook-compatibility --json',
    'packages/xlsx-formula-recalc/README.md',
  )
  requireIncludes(
    args.xlsxRecalcPackageReadme,
    'workbook-compatibility-report pricing.xlsx --json',
    'packages/xlsx-formula-recalc/README.md',
  )
  requireIncludes(args.xlsxRecalcPackageReadme, 'bilig-evaluate --door xlsx-cache --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'bilig-evaluator.v1', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-cache-doctor --demo --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-recalc --demo --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-cache-doctor pricing.xlsx --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-recalc pricing.xlsx --inspect --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'sheetjs-recalc --demo --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'If You Arrived From SheetJS or xlsx-populate', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'SheetJS formula result not updating', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'examples/recalc-bridge-workflows', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(
    args.xlsxRecalcPackageAgentNotes,
    'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json',
    'packages/xlsx-formula-recalc/AGENTS.md',
  )
  requireIncludes(
    args.xlsxRecalcPackageAgentNotes,
    'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
    'packages/xlsx-formula-recalc/AGENTS.md',
  )
  requireIncludes(
    args.xlsxRecalcPackageAgentNotes,
    'package remains a compatibility and search alias',
    'packages/xlsx-formula-recalc/AGENTS.md',
  )
  requireIncludes(args.xlsxRecalcPackageAgentNotes, 'xlsx-cache-doctor workbook.xlsx --json', 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(
    args.xlsxRecalcPackageAgentNotes,
    'workbook-compatibility-report workbook.xlsx --json',
    'packages/xlsx-formula-recalc/AGENTS.md',
  )
  requireIncludes(args.xlsxRecalcPackageAgentNotes, 'xlsx-recalc workbook.xlsx --inspect --json', 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(args.xlsxRecalcPackageAgentNotes, 'sheetjs-recalc --demo --json', 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(args.xlsxRecalcPackageAgentNotes, "from '@bilig/xlsx-formula-recalc'", 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'bilig-evaluator.v1', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(
    args.xlsxRecalcPackageSkillNotes,
    'bilig-evaluate --door workbook-compatibility --json',
    'packages/xlsx-formula-recalc/SKILL.md',
  )
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'Summary!B2', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'xlsx-cache-doctor workbook.xlsx --json', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(
    args.xlsxRecalcPackageSkillNotes,
    'workbook-compatibility-report workbook.xlsx --json',
    'packages/xlsx-formula-recalc/SKILL.md',
  )
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'xlsx-recalc workbook.xlsx --inspect --json', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'sheetjs-recalc --demo --json', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, "from '@bilig/xlsx-formula-recalc'", 'packages/xlsx-formula-recalc/SKILL.md')

  requireIncludes(
    args.sheetjsRecalcPackageJson,
    '"sheetjs-recalc": "./bin/sheetjs-recalc.js"',
    'packages/sheetjs-formula-recalc/package.json',
  )
  requireIncludes(args.sheetjsRecalcPackageReadme, 'sheetjs-recalc --demo --json', 'packages/sheetjs-formula-recalc/README.md')
  requireIncludes(
    args.sheetjsRecalcPackageReadme,
    'If You Arrived From a SheetJS Formula Issue',
    'packages/sheetjs-formula-recalc/README.md',
  )
  requireIncludes(args.sheetjsRecalcPackageReadme, 'SheetJS formula result not updating', 'packages/sheetjs-formula-recalc/README.md')
  requireIncludes(args.sheetjsRecalcPackageReadme, 'examples/recalc-bridge-workflows', 'packages/sheetjs-formula-recalc/README.md')
  requireIncludes(args.sheetjsRecalcPackageAgentNotes, 'recalculateSheetjsWorkbook', 'packages/sheetjs-formula-recalc/AGENTS.md')
  requireIncludes(args.sheetjsRecalcPackageSkillNotes, 'sheetjs-recalc --demo --json', 'packages/sheetjs-formula-recalc/SKILL.md')

  requireIncludes(
    args.exceljsRecalcPackageJson,
    '"exceljs-recalc": "./bin/exceljs-recalc.js"',
    'packages/exceljs-formula-recalc/package.json',
  )
  requireIncludes(args.exceljsRecalcPackageReadme, 'exceljs-recalc --demo --json', 'packages/exceljs-formula-recalc/README.md')
  requireIncludes(
    args.exceljsRecalcPackageReadme,
    'If You Arrived From an ExcelJS Formula Issue',
    'packages/exceljs-formula-recalc/README.md',
  )
  requireIncludes(args.exceljsRecalcPackageReadme, 'ExcelJS formula result not updating', 'packages/exceljs-formula-recalc/README.md')
  requireIncludes(args.exceljsRecalcPackageReadme, 'examples/recalc-bridge-workflows', 'packages/exceljs-formula-recalc/README.md')
  requireIncludes(args.exceljsRecalcPackageAgentNotes, 'recalculateExceljsWorkbook', 'packages/exceljs-formula-recalc/AGENTS.md')
  requireIncludes(args.exceljsRecalcPackageSkillNotes, 'exceljs-recalc --demo --json', 'packages/exceljs-formula-recalc/SKILL.md')
  requireIncludes(args.exceljsRecalcPackageSkillNotes, 'commandSucceeded: true', 'packages/exceljs-formula-recalc/SKILL.md')
  requireIncludes(args.exceljsRecalcPackageSkillNotes, 'expectedValueMatched: true', 'packages/exceljs-formula-recalc/SKILL.md')
  requireNotIncludes(args.exceljsRecalcPackageSkillNotes, 'verified: true', 'packages/exceljs-formula-recalc/SKILL.md')
}
