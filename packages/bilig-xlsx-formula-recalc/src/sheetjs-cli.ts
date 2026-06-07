#!/usr/bin/env node

import { runXlsxFormulaRecalcCliAsync } from './cli-api.js'

process.exitCode = await runXlsxFormulaRecalcCliAsync(process.argv.slice(2), {
  commandName: 'sheetjs-recalc',
})
