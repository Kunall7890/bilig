#!/usr/bin/env node
const { runBiligEvaluatorCli } = await import('@bilig/xlsx-formula-recalc/evaluator')

process.exitCode = await Promise.resolve(runBiligEvaluatorCli(process.argv.slice(2)))
