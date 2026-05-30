import { requireIncludes } from './check-docs-discovery-core.ts'

export function requirePackageCliSurfaceDiscovery(args: {
  readonly exceljsRecalcPackageAgentNotes: string
  readonly exceljsRecalcPackageJson: string
  readonly exceljsRecalcPackageReadme: string
  readonly exceljsRecalcPackageSkillNotes: string
  readonly sheetjsRecalcPackageAgentNotes: string
  readonly sheetjsRecalcPackageJson: string
  readonly sheetjsRecalcPackageReadme: string
  readonly sheetjsRecalcPackageSkillNotes: string
  readonly xlsxRecalcPackageAgentNotes: string
  readonly xlsxRecalcPackageJson: string
  readonly xlsxRecalcPackageReadme: string
  readonly xlsxRecalcPackageSkillNotes: string
}): void {
  requireIncludes(args.xlsxRecalcPackageJson, '"xlsx-recalc": "./bin/xlsx-recalc.js"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(
    args.xlsxRecalcPackageJson,
    '"xlsx-cache-doctor": "./bin/xlsx-cache-doctor.js"',
    'packages/xlsx-formula-recalc/package.json',
  )
  requireIncludes(args.xlsxRecalcPackageJson, '"sheetjs-recalc": "./bin/sheetjs-recalc.js"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(args.xlsxRecalcPackageJson, '"./cli-api"', 'packages/xlsx-formula-recalc/package.json')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-cache-doctor --demo --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-recalc --demo --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-cache-doctor pricing.xlsx --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'xlsx-recalc pricing.xlsx --inspect --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'sheetjs-recalc --demo --json', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'If You Arrived From SheetJS or xlsx-populate', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'SheetJS formula result not updating', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageReadme, 'examples/recalc-bridge-workflows', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.xlsxRecalcPackageAgentNotes, 'xlsx-recalc --demo --json', 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(args.xlsxRecalcPackageAgentNotes, 'xlsx-cache-doctor workbook.xlsx --json', 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(args.xlsxRecalcPackageAgentNotes, 'xlsx-recalc workbook.xlsx --inspect --json', 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(args.xlsxRecalcPackageAgentNotes, 'sheetjs-recalc --demo --json', 'packages/xlsx-formula-recalc/AGENTS.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'Summary!B2', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'xlsx-cache-doctor workbook.xlsx --json', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'xlsx-recalc workbook.xlsx --inspect --json', 'packages/xlsx-formula-recalc/SKILL.md')
  requireIncludes(args.xlsxRecalcPackageSkillNotes, 'sheetjs-recalc --demo --json', 'packages/xlsx-formula-recalc/SKILL.md')

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
}
