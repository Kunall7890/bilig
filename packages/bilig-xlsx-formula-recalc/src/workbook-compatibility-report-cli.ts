#!/usr/bin/env node

import { runWorkbookCompatibilityReportCli } from 'xlsx-formula-recalc/workbook-compatibility-report'

process.exitCode = runWorkbookCompatibilityReportCli(process.argv.slice(2))
