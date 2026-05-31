#!/usr/bin/env node
const { runBiligEvaluatorCli } = await import('@bilig/xlsx-formula-recalc/evaluator')

process.exitCode = runBiligEvaluatorCli(process.argv.slice(2))
