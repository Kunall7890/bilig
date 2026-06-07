#!/usr/bin/env node

import { basename } from 'node:path'

import { runXlsxFormulaRecalcCliAsync } from './cli-api.js'

process.exitCode = await runXlsxFormulaRecalcCliAsync(process.argv.slice(2), {
  commandName: basename(process.argv[1] ?? 'xlsx-recalc'),
})
