import { runBiligEvaluatorCli } from 'xlsx-formula-recalc/evaluator'

process.exitCode = await runBiligEvaluatorCli(process.argv.slice(2))
