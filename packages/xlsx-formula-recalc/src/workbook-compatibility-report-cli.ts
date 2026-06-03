#!/usr/bin/env node

import { runWorkbookCompatibilityReportCli } from './workbook-compatibility-report.js'

process.exitCode = runWorkbookCompatibilityReportCli(process.argv.slice(2))
