import { runBiligEvaluatorCli } from 'xlsx-formula-recalc/evaluator'

process.exitCode = runBiligEvaluatorCli(process.argv.slice(2))
